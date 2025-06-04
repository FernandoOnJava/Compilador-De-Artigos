using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using GongSolutions.Wpf.DragDrop;
using System.Globalization;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.ComponentModel;
using System.Runtime.CompilerServices;

// Aliases para resolver conflitos de nomes
using WpfStyle = System.Windows.Style;
using OpenXmlStyle = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace DocumentUploader
{
    // REPLACE the existing constructor and add these new properties in CompileDocumentsWindow.xaml.cs

    public partial class CompileDocumentsWindow : Window, IDropTarget, INotifyPropertyChanged
    {
        public ObservableCollection<string> SelectedFiles { get; set; }

        // Ficheiros recebidos do formulário anterior
        private string capaFilePath;
        private string conselhoFilePath;
        private string editorialFilePath;

        // NEW: Novos campos recebidos do formulário anterior
        private string magazineTitle;
        private string magazineISSN;

        // Path da ficha técnica (automaticamente detectado)
        private string GetFichaTecnicaPath()
        {
            // Procura na pasta Resources do projeto
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string resourcesPath = Path.Combine(baseDir, "Resources", "ficha_tecnica.docx");

            // Se não existir na pasta Resources, procura na pasta raiz da aplicação
            if (!File.Exists(resourcesPath))
            {
                resourcesPath = Path.Combine(baseDir, "ficha_tecnica.docx");
            }

            return resourcesPath;
        }

        private Dictionary<string, List<Author>> articleAuthors;
        private DispatcherTimer progressTimer;

        private bool _isCompiling;
        public bool IsCompiling
        {
            get => _isCompiling;
            set
            {
                _isCompiling = value;
                OnPropertyChanged();
            }
        }

        private double _progressValue;
        public double ProgressValue
        {
            get => _progressValue;
            set
            {
                _progressValue = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // UPDATED: Construtor atualizado que recebe os novos parâmetros
        public CompileDocumentsWindow(string capa, string conselho, string editorial, string title, string issn)
        {
            InitializeComponent();

            // Guardar os ficheiros e novos dados recebidos
            capaFilePath = capa;
            conselhoFilePath = conselho;
            editorialFilePath = editorial;
            magazineTitle = title ?? "TMQ - Técnicas, Metodologias e Qualidade";
            magazineISSN = issn ?? "1647-9440";

            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            articleAuthors = new Dictionary<string, List<Author>>();
            DataContext = this;

            progressTimer = new DispatcherTimer();
            progressTimer.Interval = TimeSpan.FromMilliseconds(50);
            progressTimer.Tick += ProgressTimer_Tick;

            UpdateStatus($"Ficheiros base carregados. Revista: '{magazineTitle}' (ISSN: {magazineISSN}). Adicione os artigos para compilar.");

            // Mostrar os ficheiros já carregados
            UpdateLoadedFilesDisplay();
        }

        // NEW: Propriedades públicas para acesso aos novos dados
        public string MagazineTitle => magazineTitle;
        public string MagazineISSN => magazineISSN;

        private void UpdateLoadedFilesDisplay()
        {
            if (capaStatus != null)
                capaStatus.Text = $"✅ {Path.GetFileName(capaFilePath)}";

            if (conselhoStatus != null)
                conselhoStatus.Text = $"✅ {Path.GetFileName(conselhoFilePath)}";

            if (editorialStatus != null)
                editorialStatus.Text = $"✅ {Path.GetFileName(editorialFilePath)}";
        }

        #region Drag and Drop Implementation
        public void DragOver(IDropInfo dropInfo)
        {
            if (dropInfo.Data is string && dropInfo.TargetCollection is ObservableCollection<string>)
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                dropInfo.Effects = DragDropEffects.Move;
            }
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.Copy;
            }
        }

        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.Data is string sourceItem && dropInfo.TargetCollection is ObservableCollection<string> targetCollection)
            {
                int sourceIndex = SelectedFiles.IndexOf(sourceItem);
                int targetIndex = dropInfo.InsertIndex;

                if (sourceIndex != targetIndex)
                {
                    if (sourceIndex < targetIndex)
                    {
                        targetIndex--;
                    }

                    SelectedFiles.RemoveAt(sourceIndex);
                    SelectedFiles.Insert(targetIndex, sourceItem);

                    filesListBox.SelectedIndex = targetIndex;
                    UpdateStatus("Item reordenado");
                }
            }
            else if (dropInfo.Data is IDataObject dataObject && dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])dataObject.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if (File.Exists(file) && Path.GetExtension(file).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        SelectedFiles.Add(file);
                    }
                }
                UpdateStatus($"{files.Length} ficheiro(s) adicionado(s)");
                CheckCompileEnabled();
            }
        }
        #endregion

        // SUBSTITUA o método AddFiles_Click no CompileDocumentsWindow.xaml.cs por esta versão limpa:
        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                debugTerminal?.WriteLine("🔘 Botão 'Adicionar Artigos' clicado");

                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Ficheiros Word (*.docx)|*.docx|Todos os ficheiros (*.*)|*.*",
                    Multiselect = true,
                    Title = "Selecionar Artigos para Adicionar",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    RestoreDirectory = true,
                    CheckFileExists = true,
                    CheckPathExists = true
                };

                debugTerminal?.WriteLine("📁 Abrindo diálogo de seleção de ficheiros...");

                bool? result = openFileDialog.ShowDialog(this);

                if (result == true)
                {
                    debugTerminal?.WriteLine($"✅ Utilizador selecionou {openFileDialog.FileNames.Length} ficheiro(s)");

                    int addedCount = 0;
                    int duplicateCount = 0;
                    int invalidCount = 0;

                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        debugTerminal?.WriteLine($"📄 Processando: {Path.GetFileName(fileName)}");

                        // Verificar se o ficheiro existe
                        if (!File.Exists(fileName))
                        {
                            debugTerminal?.WriteLineError($"   Ficheiro não encontrado: {fileName}");
                            invalidCount++;
                            continue;
                        }

                        // Verificar se já está na lista
                        if (SelectedFiles.Contains(fileName))
                        {
                            debugTerminal?.WriteLineWarning($"   Ficheiro já adicionado: {Path.GetFileName(fileName)}");
                            duplicateCount++;
                            continue;
                        }

                        // Verificar extensão
                        if (!Path.GetExtension(fileName).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            debugTerminal?.WriteLineWarning($"   Ficheiro não é .docx: {Path.GetFileName(fileName)}");
                            invalidCount++;
                            continue;
                        }

                        // Adicionar à lista
                        SelectedFiles.Add(fileName);
                        addedCount++;
                        debugTerminal?.WriteLineSuccess($"   ✅ Adicionado: {Path.GetFileName(fileName)}");
                    }

                    // Relatório final
                    debugTerminal?.WriteSeparator();
                    debugTerminal?.WriteLine($"📊 Relatório da adição:");
                    debugTerminal?.WriteLine($"   • Adicionados: {addedCount}");
                    debugTerminal?.WriteLine($"   • Duplicados: {duplicateCount}");
                    debugTerminal?.WriteLine($"   • Inválidos: {invalidCount}");
                    debugTerminal?.WriteLine($"   • Total na lista: {SelectedFiles.Count}");

                    if (addedCount > 0)
                    {
                        UpdateStatus($"{addedCount} artigo(s) adicionado(s) com sucesso!");
                        CheckCompileEnabled();

                        // Forçar refresh da ListBox
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            filesListBox.Items.Refresh();
                            filesListBox.UpdateLayout();
                        }), DispatcherPriority.Background);
                    }
                    else
                    {
                        string message = duplicateCount > 0 ?
                            "Todos os ficheiros já estavam na lista" :
                            "Nenhum ficheiro válido foi selecionado";
                        UpdateStatus(message);
                    }
                }
                else
                {
                    debugTerminal?.WriteLine("❌ Utilizador cancelou a seleção");
                    UpdateStatus("Seleção cancelada pelo utilizador");
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Erro ao adicionar ficheiros: {ex.Message}";
                debugTerminal?.WriteLineError(errorMsg);
                debugTerminal?.WriteLineError($"Stack trace: {ex.StackTrace}");

                MessageBox.Show(errorMsg, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus($"Erro: {ex.Message}");
            }
        }

        public void TestButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("✅ Botão de teste funcionou perfeitamente!", "Teste de Sucesso");

            try
            {
                if (SelectedFiles == null)
                {
                    MessageBox.Show("❌ SelectedFiles é null!", "Erro");
                    return;
                }

                string testFile = $"TESTE_{DateTime.Now:HHmmss}.docx";
                SelectedFiles.Add(testFile);

                // Forçar atualização da interface
                if (filesListBox != null)
                {
                    filesListBox.Items.Refresh();
                    filesListBox.UpdateLayout();
                }

                MessageBox.Show($"✅ Ficheiro de teste adicionado!\nTotal na lista: {SelectedFiles.Count}\nÚltimo item: {testFile}", "Sucesso");

                // Atualizar botão de compilação
                CheckCompileEnabled();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Erro no teste: {ex.Message}\n\nStack trace:\n{ex.StackTrace}", "Erro no Teste");
            }
        }

        // Método melhorado para verificar o estado da compilação
        private void CheckCompileEnabled()
        {
            try
            {
                bool shouldEnable = SelectedFiles != null && SelectedFiles.Count > 0;

                Console.WriteLine($"🔘 CheckCompileEnabled: {SelectedFiles?.Count ?? 0} ficheiros, botão será {(shouldEnable ? "ATIVADO" : "DESATIVADO")}");

                if (btnCompile != null)
                {
                    btnCompile.IsEnabled = shouldEnable;
                    Console.WriteLine($"✅ Botão compilar atualizado: {btnCompile.IsEnabled}");
                }
                else
                {
                    Console.WriteLine("❌ btnCompile é null!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro em CheckCompileEnabled: {ex.Message}");
            }
        }

        // Método melhorado para atualizar status
        private void UpdateStatus(string message)
        {
            try
            {
                Console.WriteLine($"📢 Status: {message}");

                Dispatcher.Invoke(() =>
                {
                    if (statusTextBlock != null)
                    {
                        statusTextBlock.Text = message;
                        Console.WriteLine("✅ Status atualizado na UI");
                    }
                    else
                    {
                        Console.WriteLine("❌ statusTextBlock é null!");
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Erro ao atualizar status: {ex.Message}");
            }
        }

        private void AddTableOfContents(Body body, List<ArticleInfo> articles)
        {
            // Title
            Paragraph titleParagraph = new Paragraph();
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            Run titleRun = new Run(new Text("Índice"));
            titleParagraph.Append(titleRun);
            body.AppendChild(titleParagraph);

            // Editorial entry
            Paragraph editorialPara = new Paragraph();
            Run editorialRun = new Run(new Text("Editorial"));
            editorialRun.RunProperties = new RunProperties(new Bold());
            editorialPara.Append(editorialRun);

            // Add page number (placeholder)
            editorialPara.Append(new Run(new Text(" ................... ")));
            editorialPara.Append(new Run(new Text("XX")));
            body.AppendChild(editorialPara);

            // Article entries
            foreach (var article in articles)
            {
                // Article title with page number
                Paragraph articlePara = new Paragraph();
                Run articleRun = new Run(new Text(article.Title));
                articleRun.RunProperties = new RunProperties(new Bold());
                articlePara.Append(articleRun);

                // Add page number (placeholder)
                articlePara.Append(new Run(new Text(" ................... ")));
                articlePara.Append(new Run(new Text("XX")));
                body.AppendChild(articlePara);

                // Authors below title
                if (article.Authors.Count > 0)
                {
                    Paragraph authorsPara = new Paragraph();
                    authorsPara.ParagraphProperties = new ParagraphProperties(
                        new Indentation() { Left = "720" } // Indent authors
                    );
                    string authorNames = string.Join(", ", article.Authors.Select(a => a.Nome));
                    Run authorsRun = new Run(new Text(authorNames));
                    authorsRun.RunProperties = new RunProperties(new Italic());
                    authorsPara.Append(authorsRun);
                    body.AppendChild(authorsPara);
                }
            }
        }

        private void AddArticleWithHeading(Body body, string filePath, string title, List<Author> authors)
        {
            // Article title with Heading1 style
            Paragraph titleParagraph = new Paragraph();
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            Run titleRun = new Run(new Text(title));
            titleParagraph.Append(titleRun);
            body.AppendChild(titleParagraph);

            // Authors in italic (nome 1, nome 2, nome 3)
            if (authors.Count > 0)
            {
                Paragraph authorParagraph = new Paragraph();
                Run authorRun = new Run();
                authorRun.RunProperties = new RunProperties(new Italic());
                string authorNames = string.Join(", ", authors.Select(a => a.Nome));
                authorRun.Append(new Text(authorNames));
                authorParagraph.Append(authorRun);
                authorParagraph.ParagraphProperties = new ParagraphProperties(
                    new SpacingBetweenLines() { After = "240" }
                );
                body.AppendChild(authorParagraph);
            }

            // Article content
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                try
                {
                    using (WordprocessingDocument articleDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        if (articleDoc.MainDocumentPart != null && articleDoc.MainDocumentPart.Document.Body != null)
                        {
                            var elements = articleDoc.MainDocumentPart.Document.Body.Elements().ToList();

                            // Skip the original title and author paragraphs when copying content
                            int startIndex = Math.Min(authors.Count + 1, elements.Count);

                            for (int i = startIndex; i < elements.Count; i++)
                            {
                                body.AppendChild(elements[i].CloneNode(true));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Paragraph errorPara = new Paragraph(new Run(new Text($"Erro ao carregar conteúdo: {ex.Message}")));
                    body.AppendChild(errorPara);
                }
            }
        }

        private void AddPageBreak(Body body)
        {
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
        }

        private void AddSettingsToDocument(MainDocumentPart mainPart)
        {
            DocumentSettingsPart settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            Settings settings = new Settings();
            settings.Append(new UpdateFieldsOnOpen() { Val = true });
            settingsPart.Settings = settings;
        }

        // Adicione este método para teste manual (pode remover depois)
        private void TestAddFiles()
        {
            Console.WriteLine("🧪 TESTE: Adicionando ficheiro manualmente...");

            if (SelectedFiles == null)
            {
                Console.WriteLine("❌ SelectedFiles é null no teste!");
                return;
            }

            // Adicionar um ficheiro de teste
            string testFile = @"C:\temp\teste.docx"; // Coloque aqui um caminho válido para teste

            if (File.Exists(testFile))
            {
                SelectedFiles.Add(testFile);
                Console.WriteLine($"✅ Ficheiro de teste adicionado: {SelectedFiles.Count} items na lista");

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    filesListBox.Items.Refresh();
                    Console.WriteLine($"✅ Interface atualizada - Items visíveis: {filesListBox.Items.Count}");
                }));
            }
            else
            {
                Console.WriteLine($"❌ Ficheiro de teste não existe: {testFile}");
            }
        }

        // Método auxiliar para verificar se o SelectedFiles está inicializado corretamente
        private void VerifySelectedFilesInitialization()
        {
            if (SelectedFiles == null)
            {
                debugTerminal?.WriteLineError("❌ SelectedFiles é null! Reinicializando...");
                SelectedFiles = new ObservableCollection<string>();
                filesListBox.ItemsSource = SelectedFiles;
            }

            if (filesListBox.ItemsSource == null)
            {
                debugTerminal?.WriteLineError("❌ filesListBox.ItemsSource é null! Reconectando...");
                filesListBox.ItemsSource = SelectedFiles;
            }
        }

        // ADICIONE este método de teste temporário tt
        private void TestAddButton()
        {
            try
            {
                MessageBox.Show("Teste: Adicionando ficheiro manualmente...", "Debug");

                // Teste básico da coleção
                if (SelectedFiles == null)
                {
                    MessageBox.Show("SelectedFiles é null!", "Erro");
                    return;
                }

                SelectedFiles.Add("TESTE_FICHEIRO.docx");

                Dispatcher.BeginInvoke(() =>
                {
                    filesListBox.Items.Refresh();
                    MessageBox.Show($"Total de items: {SelectedFiles.Count}", "Sucesso");
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro no teste: {ex.Message}", "Erro");
            }
        }

        private void RemoveFile_Click(object sender, RoutedEventArgs e)
        {
            if (filesListBox.SelectedIndex != -1)
            {
                SelectedFiles.RemoveAt(filesListBox.SelectedIndex);
                UpdateStatus("Ficheiro removido");
                CheckCompileEnabled();
            }
            else
            {
                UpdateStatus("Nenhum ficheiro selecionado");
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = filesListBox.SelectedIndex;
            if (selectedIndex > 0)
            {
                string item = SelectedFiles[selectedIndex];
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex - 1, item);
                filesListBox.SelectedIndex = selectedIndex - 1;
                UpdateStatus("Ficheiro movido para cima");
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            int selectedIndex = filesListBox.SelectedIndex;
            if (selectedIndex != -1 && selectedIndex < SelectedFiles.Count - 1)
            {
                string item = SelectedFiles[selectedIndex];
                SelectedFiles.RemoveAt(selectedIndex);
                SelectedFiles.Insert(selectedIndex + 1, item);
                filesListBox.SelectedIndex = selectedIndex + 1;
                UpdateStatus("Ficheiro movido para baixo");
            }
        }

        

        // ADICIONE estes métodos à classe CompileDocumentsWindow no arquivo .xaml.cs

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Janela carregada! Verificando componentes...", "Debug");

            // Verificar componentes críticos
            if (filesListBox == null)
                MessageBox.Show("❌ filesListBox é NULL!", "Erro");
            else
                MessageBox.Show("✅ filesListBox OK", "Debug");

            if (SelectedFiles == null)
                MessageBox.Show("❌ SelectedFiles é NULL!", "Erro");
            else
                MessageBox.Show($"✅ SelectedFiles OK - Count: {SelectedFiles.Count}", "Debug");

            if (btnCompile == null)
                MessageBox.Show("❌ btnCompile é NULL!", "Erro");
            else
                MessageBox.Show("✅ btnCompile OK", "Debug");

            if (btnAddFiles == null)
                MessageBox.Show("❌ btnAddFiles é NULL!", "Erro");
            else
                MessageBox.Show("✅ btnAddFiles OK", "Debug");

            // Verificar conexão
            if (filesListBox != null && filesListBox.ItemsSource != SelectedFiles)
            {
                MessageBox.Show("⚠️ Reconectando ItemsSource...", "Aviso");
                filesListBox.ItemsSource = SelectedFiles;
            }
        }

        private async void Compile_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedFiles.Count == 0)
            {
                MessageBox.Show("Por favor, adicione pelo menos um artigo.", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Inicializar terminal de debug se ativado
            if (SHOW_DEBUG_TERMINAL)
            {
                debugTerminal = new DebugTerminalWindow();
                debugTerminal.Owner = this;
                debugTerminal.Show();
                debugTerminal.WriteLine("🚀 Iniciando compilação da revista...");
                debugTerminal.WriteLine($"📁 Total de artigos: {SelectedFiles.Count}");
                debugTerminal.WriteSeparator();
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Ficheiro Word (*.docx)|*.docx",
                Title = "Guardar Revista Compilada",
                FileName = $"Revista_TMQ_{DateTime.Now:yyyyMMdd}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                IsCompiling = true;
                ProgressValue = 0;
                progressTimer.Start();

                debugTerminal?.WriteLine($"💾 Documento será salvo em: {saveFileDialog.FileName}");
                debugTerminal?.UpdateStatus("Compilando...");

                try
                {
                    await Task.Run(() => CreateRevistaDocument(saveFileDialog.FileName));

                    progressTimer.Stop();
                    ProgressValue = 100;
                    await Task.Delay(500);

                    UpdateStatus($"Revista compilada com sucesso!");
                    debugTerminal?.WriteLineSuccess("Revista compilada com sucesso!");
                    debugTerminal?.UpdateStatus("Concluído");

                    MessageBox.Show($"Revista guardada com sucesso!\nLocalização: {saveFileDialog.FileName}",
                        "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    progressTimer.Stop();
                    UpdateStatus("Erro: " + ex.Message);
                    debugTerminal?.WriteLineError($"Erro na compilação: {ex.Message}");
                    debugTerminal?.UpdateStatus("Erro");

                    MessageBox.Show("Erro ao compilar revista: " + ex.Message,
                        "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    IsCompiling = false;
                    ProgressValue = 0;
                }
            }
            else
            {
                // Utilizador cancelou - fechar terminal se aberto
                debugTerminal?.Hide();
            }
        }

        private void ProgressTimer_Tick(object sender, EventArgs e)
        {
            if (ProgressValue < 90)
            {
                ProgressValue += 2;
            }
        }

        // SUBSTITUA o método CreateRevistaDocument e métodos relacionados

        // Código original do método CreateRevistaDocument restaurado

        private void CreateRevistaDocument(string outputPath)
        {
            debugTerminal?.WriteLine("📄 Criando documento Word...");

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                debugTerminal?.WriteLine("🎨 Definindo estilos do documento...");

                // Define styles
                StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new Styles();
                stylesPart.Styles = styles;
                AddCustomStyles(styles);

                debugTerminal?.WriteLine("🔍 Iniciando extração de informações dos artigos...");
                debugTerminal?.WriteSeparator();

                // Extract all authors and article info
                articleAuthors.Clear();
                var allAuthors = new List<Author>();
                var articleInfoList = new List<ArticleInfo>();

                int articleCount = 1;
                foreach (var article in SelectedFiles)
                {
                    debugTerminal?.WriteLine($"📖 Processando artigo {articleCount}/{SelectedFiles.Count}: {Path.GetFileName(article)}");

                    var articleInfo = ExtractArticleInfo(article);
                    if (articleInfo != null)
                    {
                        articleInfoList.Add(articleInfo);
                        articleAuthors[article] = articleInfo.Authors;
                        allAuthors.AddRange(articleInfo.Authors);

                        debugTerminal?.WriteLineInfo($"   Título: {articleInfo.Title}");
                        debugTerminal?.WriteLineInfo($"   Autores encontrados: {articleInfo.Authors.Count}");

                        foreach (var author in articleInfo.Authors)
                        {
                            debugTerminal?.WriteLine($"     👤 {author.Nome}" +
                                (!string.IsNullOrEmpty(author.Email) ? $" ({author.Email})" : "") +
                                (!string.IsNullOrEmpty(author.Escola) ? $" - {author.Escola}" : ""));
                        }
                    }
                    else
                    {
                        debugTerminal?.WriteLineWarning($"   Falhou ao extrair informações do artigo");
                    }

                    debugTerminal?.WriteLine("");
                    articleCount++;
                }

                debugTerminal?.WriteSeparator();
                debugTerminal?.WriteLine("👥 Processando lista de autores...");

                // Remove duplicates from author list
                int totalAuthors = allAuthors.Count;
                allAuthors = allAuthors.GroupBy(a => a.Email ?? a.Nome)
                    .Select(g => g.First())
                    .OrderBy(a => a.Nome)
                    .ToList();

                debugTerminal?.WriteLineSuccess($"Lista de autores criada: {allAuthors.Count} autores únicos (de {totalAuthors} totais)");

                foreach (var author in allAuthors)
                {
                    debugTerminal?.WriteLine($"  📝 {author.Nome}" +
                        (!string.IsNullOrEmpty(author.Email) ? $" - {author.Email}" : "") +
                        (!string.IsNullOrEmpty(author.Escola) ? $" - {author.Escola}" : ""));
                }

                debugTerminal?.WriteSeparator();
                debugTerminal?.WriteLine("📑 Montando documento na ordem específica...");

                // ORDEM SOLICITADA:
                // 1. Capa
                debugTerminal?.WriteLine("1️⃣ Adicionando Capa...");
                AddDocument(body, capaFilePath, "Capa");
                AddPageBreak(body);

                // 2. Folha em Branco
                debugTerminal?.WriteLine("2️⃣ Adicionando Página em Branco...");
                AddBlankPage(body);
                AddPageBreak(body);

                // 3. Ficha Técnica
                debugTerminal?.WriteLine("3️⃣ Adicionando Ficha Técnica...");
                AddDocument(body, GetFichaTecnicaPath(), "Ficha Técnica");
                AddPageBreak(body);

                // 4. Conselho Editorial
                debugTerminal?.WriteLine("4️⃣ Adicionando Conselho Editorial...");
                AddDocument(body, conselhoFilePath, "Conselho Editorial");
                AddPageBreak(body);

                // 5. Lista de Autores
                debugTerminal?.WriteLine("5️⃣ Adicionando Lista de Autores...");
                AddAuthorList(body, allAuthors);
                AddPageBreak(body);

                // 6. Índice
                debugTerminal?.WriteLine("6️⃣ Adicionando Índice...");
                AddTableOfContents(body, articleInfoList);
                AddPageBreak(body);

                // 7. Editorial
                debugTerminal?.WriteLine("7️⃣ Adicionando Editorial...");
                AddArticleWithHeading(body, editorialFilePath, "Editorial", new List<Author>());
                AddPageBreak(body);

                // 8. Artigos
                debugTerminal?.WriteLine("8️⃣ Adicionando Artigos...");
                int artCount = 1;
                foreach (var articleInfo in articleInfoList)
                {
                    debugTerminal?.WriteLine($"   📄 Artigo {artCount}/{articleInfoList.Count}: {articleInfo.Title}");
                    AddArticleWithHeading(body, articleInfo.FilePath, articleInfo.Title, articleInfo.Authors);
                    AddPageBreak(body);
                    artCount++;
                }

                debugTerminal?.WriteLine("⚙️ Configurando atualização automática de campos...");
                // Update fields (for TOC)
                AddSettingsToDocument(mainPart);

                debugTerminal?.WriteLine("💾 Salvando documento...");
                mainPart.Document.Save();

                debugTerminal?.WriteSeparator();
                debugTerminal?.WriteLineSuccess("✨ Documento compilado com sucesso!");
                debugTerminal?.WriteLine($"📊 Estatísticas finais:");
                debugTerminal?.WriteLine($"   • Total de artigos: {articleInfoList.Count}");
                debugTerminal?.WriteLine($"   • Total de autores únicos: {allAuthors.Count}");
                debugTerminal?.WriteLine($"   • Localização: {outputPath}");
            }
        }

        // Método seguro para extrair informações do artigo
        private ArticleInfo ExtractArticleInfoSafe(string filePath)
        {
            var articleInfo = new ArticleInfo
            {
                FilePath = filePath,
                Title = Path.GetFileNameWithoutExtension(filePath),
                Authors = new List<Author>()
            };

            try
            {
                string content = ReadFileContent(filePath);
                if (!string.IsNullOrEmpty(content))
                {
                    // Extrair título da primeira linha não vazia
                    var lines = content.Split('\n');
                    foreach (var line in lines)
                    {
                        string cleanLine = line.Trim();
                        if (!string.IsNullOrEmpty(cleanLine))
                        {
                            articleInfo.Title = cleanLine;
                            break;
                        }
                    }

                    // Tentativa simples de extrair autores
                    foreach (var line in lines.Take(10)) // Verificar apenas as primeiras 10 linhas
                    {
                        var author = ParseAuthorSimple(line);
                        if (author != null)
                        {
                            articleInfo.Authors.Add(author);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineWarning($"Aviso ao processar {Path.GetFileName(filePath)}: {ex.Message}");
            }

            return articleInfo;
        }

        // Método simples para ler conteúdo de arquivo
        private string ReadFileContent(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return string.Empty;
            }

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                {
                    if (doc.MainDocumentPart?.Document?.Body != null)
                    {
                        return doc.MainDocumentPart.Document.Body.InnerText;
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineWarning($"Erro ao ler {Path.GetFileName(filePath)}: {ex.Message}");
                return $"Erro ao carregar conteúdo de {Path.GetFileName(filePath)}";
            }

            return string.Empty;
        }

        // Método para parsing simples de autor
        private Author ParseAuthorSimple(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;

            var emailMatch = Regex.Match(text, @"\b[\w\.-]+@[\w\.-]+\.\w+\b");
            if (emailMatch.Success)
            {
                return new Author
                {
                    Nome = text.Replace(emailMatch.Value, "").Trim(),
                    Email = emailMatch.Value
                };
            }

            // Se tiver mais de 50 caracteres, provavelmente não é um autor
            if (text.Length > 50) return null;

            // Se contiver palavras-chave comuns de não-autor
            if (text.ToLower().Contains("resumo") || text.ToLower().Contains("abstract") ||
                text.ToLower().Contains("introdução") || text.ToLower().Contains("palavras-chave"))
                return null;

            return new Author { Nome = text.Trim() };
        }

        // Adicionar seção simples
        private void AddSimpleSection(Body body, string title, string content)
        {
            // Título
            if (!string.IsNullOrEmpty(title))
            {
                Paragraph titlePara = new Paragraph();
                Run titleRun = new Run();
                RunProperties titleProps = new RunProperties();
                titleProps.Append(new Bold());
                titleProps.Append(new FontSize() { Val = "28" });
                titleRun.RunProperties = titleProps;
                titleRun.Append(new Text(title));
                titlePara.Append(titleRun);
                body.Append(titlePara);

                // Espaço após título
                body.Append(new Paragraph());
            }

            // Conteúdo
            if (!string.IsNullOrEmpty(content))
            {
                var lines = content.Split('\n');
                foreach (var line in lines)
                {
                    string cleanLine = line.Trim();
                    if (!string.IsNullOrEmpty(cleanLine))
                    {
                        AddSimpleParagraph(body, cleanLine);
                    }
                }
            }
            else
            {
                AddSimpleParagraph(body, "Conteúdo não disponível");
            }
        }

        // Adicionar parágrafo simples
        private void AddSimpleParagraph(Body body, string text)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            run.Append(new Text(text ?? ""));
            para.Append(run);
            body.Append(para);
        }

        // Adicionar quebra de página simples
        private void AddSimplePageBreak(Body body)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            run.Append(new Break() { Type = BreakValues.Page });
            para.Append(run);
            body.Append(para);
        }

        // Adicionar lista de autores simples
        private void AddAuthorListSimple(Body body, List<Author> authors)
        {
            AddSimpleSection(body, "LISTA DE AUTORES", "");

            foreach (var author in authors)
            {
                string authorText = author.Nome;
                if (!string.IsNullOrEmpty(author.Email))
                    authorText += $" - {author.Email}";
                if (!string.IsNullOrEmpty(author.Escola))
                    authorText += $" - {author.Escola}";

                AddSimpleParagraph(body, $"• {authorText}");
            }
        }

        // Adicionar índice simples
        private void AddIndexSimple(Body body, List<ArticleInfo> articles)
        {
            AddSimpleSection(body, "ÍNDICE", "");

            AddSimpleParagraph(body, "Editorial ................................. XX");

            foreach (var article in articles)
            {
                AddSimpleParagraph(body, $"{article.Title} ................................. XX");

                if (article.Authors.Count > 0)
                {
                    string authors = string.Join(", ", article.Authors.Select(a => a.Nome));
                    AddSimpleParagraph(body, $"    {authors}");
                }
            }
        }

        // Adicionar artigo simples
        private void AddArticleSimple(Body body, ArticleInfo article)
        {
            // Título do artigo
            Paragraph titlePara = new Paragraph();
            Run titleRun = new Run();
            RunProperties titleProps = new RunProperties();
            titleProps.Append(new Bold());
            titleProps.Append(new FontSize() { Val = "24" });
            titleRun.RunProperties = titleProps;
            titleRun.Append(new Text(article.Title));
            titlePara.Append(titleRun);
            body.Append(titlePara);

            // Autores
            if (article.Authors.Count > 0)
            {
                Paragraph authorPara = new Paragraph();
                Run authorRun = new Run();
                RunProperties authorProps = new RunProperties();
                authorProps.Append(new Italic());
                authorRun.RunProperties = authorProps;
                string authorNames = string.Join(", ", article.Authors.Select(a => a.Nome));
                authorRun.Append(new Text(authorNames));
                authorPara.Append(authorRun);
                body.Append(authorPara);
            }

            // Espaço
            body.Append(new Paragraph());

            // Conteúdo do artigo
            string content = ReadFileContent(article.FilePath);
            if (!string.IsNullOrEmpty(content))
            {
                var lines = content.Split('\n');
                // Pular título e autores (primeiras linhas)
                var contentLines = lines.Skip(Math.Min(3, lines.Length)).ToArray();

                foreach (var line in contentLines)
                {
                    string cleanLine = line.Trim();
                    if (!string.IsNullOrEmpty(cleanLine))
                    {
                        AddSimpleParagraph(body, cleanLine);
                    }
                }
            }
            else
            {
                AddSimpleParagraph(body, "Conteúdo do artigo não disponível.");
            }
        }

        // Método auxiliar para criar estilos de forma segura
        private void CreateDocumentStyles(MainDocumentPart mainPart)
        {
            try
            {
                StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new Styles();

                // Estilo para títulos
                var headingStyle = new DocumentFormat.OpenXml.Wordprocessing.Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Heading1"
                };
                headingStyle.Append(new StyleName() { Val = "Heading 1" });

                var headingPPr = new StyleParagraphProperties();
                headingPPr.Append(new SpacingBetweenLines() { Before = "240", After = "120" });
                headingStyle.Append(headingPPr);

                var headingRPr = new StyleRunProperties();
                headingRPr.Append(new Bold());
                headingRPr.Append(new FontSize() { Val = "32" });
                headingStyle.Append(headingRPr);

                styles.Append(headingStyle);
                stylesPart.Styles = styles;
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineWarning($"Aviso ao criar estilos: {ex.Message}");
            }
        }

        // Método seguro para adicionar documentos
        private void AddDocumentSafely(Body body, string filePath, string title)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                debugTerminal?.WriteLineWarning($"   Ficheiro não encontrado: {filePath}");
                AddPlaceholderSection(body, title, $"{title} - Ficheiro não encontrado");
                return;
            }

            try
            {
                using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(filePath, false))
                {
                    if (sourceDoc.MainDocumentPart?.Document?.Body != null)
                    {
                        foreach (var element in sourceDoc.MainDocumentPart.Document.Body.Elements())
                        {
                            var clonedElement = element.CloneNode(true);
                            body.AppendChild(clonedElement);
                        }
                        debugTerminal?.WriteLineSuccess($"   ✅ {title} adicionado com sucesso");
                    }
                    else
                    {
                        debugTerminal?.WriteLineWarning($"   Documento {title} está vazio ou corrompido");
                        AddPlaceholderSection(body, title, $"{title} - Documento vazio ou corrompido");
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   Erro ao processar {title}: {ex.Message}");
                AddPlaceholderSection(body, title, $"Erro ao carregar {title}: {ex.Message}");
            }
        }

        // Método para adicionar seção placeholder
        private void AddPlaceholderSection(Body body, string title, string message)
        {
            var titlePara = new Paragraph();
            titlePara.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            titlePara.Append(new Run(new Text(title)));
            body.AppendChild(titlePara);

            var messagePara = new Paragraph();
            messagePara.Append(new Run(new Text(message)));
            body.AppendChild(messagePara);
        }

        // Método seguro para adicionar lista de autores
        private void AddAuthorListSafely(Body body, List<Author> authors)
        {
            try
            {
                // Título
                var titleParagraph = new Paragraph();
                titleParagraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Heading1" }
                );
                titleParagraph.Append(new Run(new Text("Lista de Autores")));
                body.AppendChild(titleParagraph);

                // Autores
                foreach (var author in authors)
                {
                    string authorText = author.Nome;
                    if (!string.IsNullOrEmpty(author.Email))
                        authorText += $" - {author.Email}";
                    if (!string.IsNullOrEmpty(author.Escola))
                        authorText += $" - {author.Escola}";

                    var authorParagraph = new Paragraph();
                    authorParagraph.Append(new Run(new Text(authorText)));
                    body.AppendChild(authorParagraph);
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   Erro ao criar lista de autores: {ex.Message}");
                AddPlaceholderSection(body, "Lista de Autores", "Erro ao criar lista de autores");
            }
        }

        // Método seguro para adicionar índice
        private void AddTableOfContentsSafely(Body body, List<ArticleInfo> articles)
        {
            try
            {
                // Título
                var titleParagraph = new Paragraph();
                titleParagraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Heading1" }
                );
                titleParagraph.Append(new Run(new Text("Índice")));
                body.AppendChild(titleParagraph);

                // Editorial
                var editorialPara = new Paragraph();
                var editorialRun = new Run(new Text("Editorial"));
                editorialRun.RunProperties = new RunProperties(new Bold());
                editorialPara.Append(editorialRun);
                editorialPara.Append(new Run(new Text(" ................... XX")));
                body.AppendChild(editorialPara);

                // Artigos
                foreach (var article in articles)
                {
                    var articlePara = new Paragraph();
                    var articleRun = new Run(new Text(article.Title));
                    articleRun.RunProperties = new RunProperties(new Bold());
                    articlePara.Append(articleRun);
                    articlePara.Append(new Run(new Text(" ................... XX")));
                    body.AppendChild(articlePara);

                    // Autores
                    if (article.Authors.Count > 0)
                    {
                        var authorsPara = new Paragraph();
                        authorsPara.ParagraphProperties = new ParagraphProperties(
                            new Indentation() { Left = "720" }
                        );
                        string authorNames = string.Join(", ", article.Authors.Select(a => a.Nome));
                        var authorsRun = new Run(new Text(authorNames));
                        authorsRun.RunProperties = new RunProperties(new Italic());
                        authorsPara.Append(authorsRun);
                        body.AppendChild(authorsPara);
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   Erro ao criar índice: {ex.Message}");
                AddPlaceholderSection(body, "Índice", "Erro ao criar índice");
            }
        }

        // Método seguro para adicionar artigos com cabeçalho
        private void AddArticleWithHeadingSafely(Body body, string filePath, string title, List<Author> authors)
        {
            try
            {
                // Título do artigo
                var titleParagraph = new Paragraph();
                titleParagraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Heading1" }
                );
                titleParagraph.Append(new Run(new Text(title)));
                body.AppendChild(titleParagraph);

                // Autores
                if (authors.Count > 0)
                {
                    var authorParagraph = new Paragraph();
                    var authorRun = new Run();
                    authorRun.RunProperties = new RunProperties(new Italic());
                    string authorNames = string.Join(", ", authors.Select(a => a.Nome));
                    authorRun.Append(new Text(authorNames));
                    authorParagraph.Append(authorRun);
                    authorParagraph.ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines() { After = "240" }
                    );
                    body.AppendChild(authorParagraph);
                }

                // Conteúdo do artigo
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    using (WordprocessingDocument articleDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        if (articleDoc.MainDocumentPart?.Document?.Body != null)
                        {
                            var elements = articleDoc.MainDocumentPart.Document.Body.Elements().ToList();

                            // Pular título e autores originais
                            int startIndex = Math.Min(authors.Count + 1, elements.Count);

                            for (int i = startIndex; i < elements.Count; i++)
                            {
                                var clonedElement = elements[i].CloneNode(true);
                                body.AppendChild(clonedElement);
                            }
                        }
                    }
                }
                else
                {
                    var errorPara = new Paragraph();
                    errorPara.Append(new Run(new Text($"Conteúdo do artigo não disponível: {Path.GetFileName(filePath)}")));
                    body.AppendChild(errorPara);
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   Erro ao adicionar artigo {title}: {ex.Message}");
                AddPlaceholderSection(body, title, $"Erro ao carregar artigo: {ex.Message}");
            }
        }


        private void AddCustomStyles(Styles styles)
        {
            // Heading1 style for articles - usando OpenXmlStyle explicitamente
            OpenXmlStyle heading1Style = new OpenXmlStyle()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
                StyleName = new StyleName() { Val = "Heading 1" }
            };

            StyleParagraphProperties heading1PPr = new StyleParagraphProperties();
            heading1PPr.Append(new OutlineLevel() { Val = 0 });
            heading1PPr.Append(new SpacingBetweenLines() { Before = "240", After = "120" });
            heading1Style.Append(heading1PPr);

            StyleRunProperties heading1RPr = new StyleRunProperties();
            heading1RPr.Append(new Bold());
            heading1RPr.Append(new FontSize() { Val = "32" });
            heading1Style.Append(heading1RPr);

            styles.Append(heading1Style);

            // Author style - usando OpenXmlStyle explicitamente
            OpenXmlStyle authorStyle = new OpenXmlStyle()
            {
                Type = StyleValues.Paragraph,
                StyleId = "ArticleAuthor",
                StyleName = new StyleName() { Val = "Article Author" }
            };

            StyleParagraphProperties authorPPr = new StyleParagraphProperties();
            authorPPr.Append(new SpacingBetweenLines() { Before = "0", After = "240" });
            authorStyle.Append(authorPPr);

            StyleRunProperties authorRPr = new StyleRunProperties();
            authorRPr.Append(new Italic());
            authorRPr.Append(new FontSize() { Val = "24" });
            authorStyle.Append(authorRPr);

            styles.Append(authorStyle);
        }

        private Author ParseAuthor(string text)
        {
            var emailMatch = Regex.Match(text, @"\b[\w\.-]+@[\w\.-]+\.\w+\b");
            string email = emailMatch.Success ? emailMatch.Value : string.Empty;

            var idMatch = Regex.Match(text, @"\b\d{5,}\b");
            string id = idMatch.Success ? idMatch.Value : string.Empty;

            string remaining = text;
            if (emailMatch.Success) remaining = remaining.Replace(email, "");
            if (idMatch.Success) remaining = remaining.Replace(id, "");

            remaining = Regex.Replace(remaining, @"Email|E-mail|ID|Id|^\d+\s*[-–]\s*", "", RegexOptions.IgnoreCase).Trim();
            remaining = remaining.Trim(' ', '-', ',', '.', '–');

            var parts = remaining.Split(new[] { '-', '–', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(p => p.Trim()).ToArray();

            string name = parts.Length > 0 ? parts[0] : string.Empty;
            string school = parts.Length > 1 ? string.Join(" - ", parts.Skip(1)) : string.Empty;

            if (string.IsNullOrEmpty(name) && string.IsNullOrEmpty(email)) return null;

            return new Author
            {
                Nome = name,
                Email = email,
                Escola = school,
                Id = id
            };
        }

        private void AddArticleWithFullFormatting(WordprocessingDocument targetDoc, string sourceFilePath, string title, List<Author> authors)
        {
            if (string.IsNullOrEmpty(sourceFilePath) || !File.Exists(sourceFilePath))
            {
                debugTerminal?.WriteLineWarning($"   Arquivo não encontrado: {sourceFilePath}");
                return;
            }

            try
            {
                // Adicionar título do artigo
                var titleParagraph = new Paragraph();
                titleParagraph.ParagraphProperties = new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Heading1" }
                );
                var titleRun = new Run(new Text(title));
                titleParagraph.Append(titleRun);
                targetDoc.MainDocumentPart.Document.Body.AppendChild(titleParagraph);

                // Adicionar autores
                if (authors.Count > 0)
                {
                    var authorParagraph = new Paragraph();
                    var authorRun = new Run();
                    authorRun.RunProperties = new RunProperties(new Italic());
                    string authorNames = string.Join(", ", authors.Select(a => a.Nome));
                    authorRun.Append(new Text(authorNames));
                    authorParagraph.Append(authorRun);
                    authorParagraph.ParagraphProperties = new ParagraphProperties(
                        new SpacingBetweenLines() { After = "240" }
                    );
                    targetDoc.MainDocumentPart.Document.Body.AppendChild(authorParagraph);
                }

                // Copiar conteúdo do artigo com formatação completa
                using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceFilePath, false))
                {
                    debugTerminal?.WriteLineInfo($"   Copiando formatação do artigo {Path.GetFileName(sourceFilePath)}...");

                    // Copiar estilos específicos do artigo
                    CopyStyles(sourceDoc, targetDoc);

                    // Copiar partes relacionadas (imagens, gráficos, etc.)
                    CopyRelatedParts(sourceDoc, targetDoc);

                    // Copiar conteúdo do artigo (pulando título e autores originais)
                    if (sourceDoc.MainDocumentPart?.Document?.Body != null)
                    {
                        var sourceElements = sourceDoc.MainDocumentPart.Document.Body.Elements().ToList();

                        // Pular os primeiros parágrafos (título e autores)
                        int startIndex = Math.Min(authors.Count + 1, sourceElements.Count);

                        for (int i = startIndex; i < sourceElements.Count; i++)
                        {
                            var clonedElement = sourceElements[i].CloneNode(true);
                            targetDoc.MainDocumentPart.Document.Body.AppendChild(clonedElement);
                        }
                    }

                    debugTerminal?.WriteLineSuccess($"   ✅ Artigo {Path.GetFileName(sourceFilePath)} copiado com formatação completa");
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   ❌ Erro ao copiar artigo {title}: {ex.Message}");
                var errorPara = new Paragraph(new Run(new Text($"Erro ao carregar artigo: {ex.Message}")));
                targetDoc.MainDocumentPart.Document.Body.AppendChild(errorPara);
            }
        }

        // Método para copiar estilos do documento fonte
        private void CopyStyles(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
        {
            if (sourceDoc.MainDocumentPart.StyleDefinitionsPart == null) return;

            // Garantir que o documento de destino tem StyleDefinitionsPart
            if (targetDoc.MainDocumentPart.StyleDefinitionsPart == null)
            {
                targetDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                targetDoc.MainDocumentPart.StyleDefinitionsPart.Styles = new Styles();
            }

            var sourceStyles = sourceDoc.MainDocumentPart.StyleDefinitionsPart.Styles;
            var targetStyles = targetDoc.MainDocumentPart.StyleDefinitionsPart.Styles;

            if (sourceStyles?.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>() != null)
            {
                foreach (var style in sourceStyles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>())
                {
                    // Verificar se o estilo já existe no documento de destino
                    string styleId = style.StyleId?.Value ?? "";
                    if (!string.IsNullOrEmpty(styleId))
                    {
                        var existingStyle = targetStyles.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
                            .FirstOrDefault(s => s.StyleId?.Value == styleId);

                        if (existingStyle == null)
                        {
                            // Clonar e adicionar o estilo
                            var clonedStyle = (DocumentFormat.OpenXml.Wordprocessing.Style)style.CloneNode(true);
                            targetStyles.AppendChild(clonedStyle);
                        }
                    }
                }
            }
        }

        // Método para copiar temas
        private void CopyTheme(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
        {
            if (sourceDoc.MainDocumentPart.ThemePart != null && targetDoc.MainDocumentPart.ThemePart == null)
            {
                var targetThemePart = targetDoc.MainDocumentPart.AddNewPart<ThemePart>();
                targetThemePart.Theme = (DocumentFormat.OpenXml.Drawing.Theme)sourceDoc.MainDocumentPart.ThemePart.Theme.CloneNode(true);
            }
        }

        // Método para copiar partes relacionadas (imagens, gráficos, etc.)
        private void CopyRelatedParts(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
        {
            try
            {
                // Copiar imagens
                foreach (var imagePart in sourceDoc.MainDocumentPart.ImageParts)
                {
                    var targetImagePart = targetDoc.MainDocumentPart.AddImagePart(imagePart.ContentType);
                    using (var sourceStream = imagePart.GetStream())
                    using (var targetStream = targetImagePart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        sourceStream.CopyTo(targetStream);
                    }

                    // Atualizar referências de imagem
                    UpdateImageReferences(targetDoc, imagePart.Uri.ToString(), targetImagePart.Uri.ToString());
                }

                // Copiar outras partes relacionadas conforme necessário
                // Pode adicionar aqui outros tipos de parte como ChartParts, DiagramParts, etc.
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineWarning($"   ⚠️ Aviso ao copiar partes relacionadas: {ex.Message}");
            }
        }

        // Método para atualizar referências de imagem
        private void UpdateImageReferences(WordprocessingDocument targetDoc, string oldImageId, string newImageId)
        {
            try
            {
                var body = targetDoc.MainDocumentPart.Document.Body;

                // Procurar e atualizar todas as referências de imagem
                var blips = body.Descendants<DocumentFormat.OpenXml.Drawing.Blip>();
                foreach (var blip in blips)
                {
                    if (blip.Embed?.Value == oldImageId)
                    {
                        blip.Embed = newImageId;
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineWarning($"   ⚠️ Aviso ao atualizar referências de imagem: {ex.Message}");
            }
        }

        // Método para copiar fontes personalizadas (se necessário)
        private void CopyFonts(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
        {
            if (sourceDoc.MainDocumentPart.FontTablePart != null && targetDoc.MainDocumentPart.FontTablePart == null)
            {
                var targetFontTablePart = targetDoc.MainDocumentPart.AddNewPart<FontTablePart>();
                targetFontTablePart.Fonts = (DocumentFormat.OpenXml.Wordprocessing.Fonts)sourceDoc.MainDocumentPart.FontTablePart.Fonts.CloneNode(true);
            }
        }

        private void AddDocumentWithFullFormatting(WordprocessingDocument targetDoc, string sourceFilePath, string title)
        {
            if (string.IsNullOrEmpty(sourceFilePath) || !File.Exists(sourceFilePath))
            {
                debugTerminal?.WriteLineWarning($"   Arquivo não encontrado: {sourceFilePath}");
                var errorPara = new Paragraph(new Run(new Text($"{title} - Ficheiro não encontrado")));
                targetDoc.MainDocumentPart.Document.Body.AppendChild(errorPara);
                return;
            }

            try
            {
                using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(sourceFilePath, false))
                {
                    debugTerminal?.WriteLineInfo($"   Copiando estilos e formatação de {Path.GetFileName(sourceFilePath)}...");

                    // 1. Copiar estilos
                    CopyStyles(sourceDoc, targetDoc);

                    // 2. Copiar temas
                    CopyTheme(sourceDoc, targetDoc);

                    // 3. Copiar partes relacionadas (imagens, gráficos, etc.)
                    CopyRelatedParts(sourceDoc, targetDoc);

                    // 4. Copiar conteúdo do corpo do documento
                    if (sourceDoc.MainDocumentPart?.Document?.Body != null)
                    {
                        var sourceElements = sourceDoc.MainDocumentPart.Document.Body.Elements().ToList();
                        foreach (var element in sourceElements)
                        {
                            var clonedElement = element.CloneNode(true);
                            targetDoc.MainDocumentPart.Document.Body.AppendChild(clonedElement);
                        }
                    }

                    debugTerminal?.WriteLineSuccess($"   ✅ Formatação preservada para {Path.GetFileName(sourceFilePath)}");
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"   ❌ Erro ao copiar {title}: {ex.Message}");
                var errorPara = new Paragraph(new Run(new Text($"Erro ao carregar {title}: {ex.Message}")));
                targetDoc.MainDocumentPart.Document.Body.AppendChild(errorPara);
            }
        }

        private void AddDocument(Body body, string filePath, string title)
        {
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                try
                {
                    using (WordprocessingDocument sourceDoc = WordprocessingDocument.Open(filePath, false))
                    {
                        if (sourceDoc.MainDocumentPart != null && sourceDoc.MainDocumentPart.Document.Body != null)
                        {
                            foreach (var element in sourceDoc.MainDocumentPart.Document.Body.Elements())
                            {
                                body.AppendChild(element.CloneNode(true));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Se houver erro, adicionar uma mensagem
                    Paragraph errorPara = new Paragraph(new Run(new Text($"Erro ao carregar {title}: {ex.Message}")));
                    body.AppendChild(errorPara);
                }
            }
            else
            {
                // Se o ficheiro não existir
                Paragraph missingPara = new Paragraph(new Run(new Text($"{title} - Ficheiro não encontrado")));
                body.AppendChild(missingPara);
            }
        }

        private void AddBlankPage(Body body)
        {
            // Adicionar página em branco
            Paragraph blankPara = new Paragraph();
            body.AppendChild(blankPara);
        }

        private void AddAuthorList(Body body, List<Author> authors)
        {
            // Title
            Paragraph titleParagraph = new Paragraph();
            titleParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId() { Val = "Heading1" }
            );
            Run titleRun = new Run(new Text("Lista de Autores"));
            titleParagraph.Append(titleRun);
            body.AppendChild(titleParagraph);

            // Authors
            foreach (var author in authors)
            {
                string authorText = author.Nome;
                if (!string.IsNullOrEmpty(author.Email))
                    authorText += $" - {author.Email}";
                if (!string.IsNullOrEmpty(author.Escola))
                    authorText += $" - {author.Escola}";

                Paragraph authorParagraph = new Paragraph(new Run(new Text(authorText)));
                body.AppendChild(authorParagraph);
            }
        }
        

        private DebugTerminalWindow? debugTerminal;

        // Se necessário para desenvolvimento trocar o falso para true
        private const bool SHOW_DEBUG_TERMINAL = false;

        private void filesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }

    // Support classes
    public class Author
    {
        public string Nome { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Escola { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
    }

    public class ArticleInfo
    {
        public string FilePath { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public List<Author> Authors { get; set; } = new List<Author>();
    }

    // Converters
    public class FileNameConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string filePath)
            {
                return Path.GetFileName(filePath);
            }
            return value?.ToString() ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class FileIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Retorna null - sem ícone
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}