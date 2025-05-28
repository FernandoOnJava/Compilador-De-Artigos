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
    public partial class CompileDocumentsWindow : Window, IDropTarget, INotifyPropertyChanged
    {
        public ObservableCollection<string> SelectedFiles { get; set; }

        // Ficheiros recebidos do formulário anterior
        private string capaFilePath;
        private string conselhoFilePath;
        private string editorialFilePath;

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

        // Construtor que recebe os ficheiros do formulário anterior
        public CompileDocumentsWindow(string capa, string conselho, string editorial)
        {
            InitializeComponent();

            // Guardar os ficheiros recebidos
            capaFilePath = capa;
            conselhoFilePath = conselho;
            editorialFilePath = editorial;

            SelectedFiles = new ObservableCollection<string>();
            filesListBox.ItemsSource = SelectedFiles;
            articleAuthors = new Dictionary<string, List<Author>>();
            DataContext = this;

            progressTimer = new DispatcherTimer();
            progressTimer.Interval = TimeSpan.FromMilliseconds(50);
            progressTimer.Tick += ProgressTimer_Tick;

            UpdateStatus($"Ficheiros base carregados. Adicione os artigos para compilar a revista.");

            // Mostrar os ficheiros já carregados
            UpdateLoadedFilesDisplay();
        }

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

        private void CheckCompileEnabled()
        {
            btnCompile.IsEnabled = SelectedFiles.Count > 0;
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

        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() => statusTextBlock.Text = message);
        }

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

        private ArticleInfo ExtractArticleInfo(string filePath)
        {
            var articleInfo = new ArticleInfo
            {
                FilePath = filePath,
                Title = Path.GetFileNameWithoutExtension(filePath),
                Authors = new List<Author>()
            };

            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    if (wordDoc.MainDocumentPart != null && wordDoc.MainDocumentPart.Document.Body != null)
                    {
                        var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();
                        debugTerminal?.WriteLineInfo($"     📄 Total de parágrafos encontrados: {paragraphs.Count}");

                        // First paragraph is usually the title
                        if (paragraphs.Count > 0)
                        {
                            string extractedTitle = paragraphs[0].InnerText.Trim();
                            if (!string.IsNullOrEmpty(extractedTitle))
                            {
                                articleInfo.Title = extractedTitle;
                                debugTerminal?.WriteLineInfo($"     📝 Título extraído: {extractedTitle}");
                            }
                        }

                        // Find authors (before Abstract/Resumo)
                        int abstractIndex = -1;
                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (text.StartsWith("Resumo", StringComparison.OrdinalIgnoreCase) ||
                                text.StartsWith("Abstract", StringComparison.OrdinalIgnoreCase))
                            {
                                abstractIndex = i;
                                debugTerminal?.WriteLineInfo($"     🔍 Abstract/Resumo encontrado no parágrafo {i}");
                                break;
                            }
                        }

                        if (abstractIndex < 0)
                        {
                            abstractIndex = paragraphs.Count;
                            debugTerminal?.WriteLineWarning($"     ⚠️ Abstract/Resumo não encontrado, processando até o final");
                        }

                        // Extract authors from paragraphs 1 to abstractIndex-1
                        debugTerminal?.WriteLineInfo($"     👥 Procurando autores nos parágrafos 1 a {abstractIndex - 1}");
                        for (int i = 1; i < abstractIndex && i < paragraphs.Count; i++)
                        {
                            string text = paragraphs[i].InnerText.Trim();
                            if (!string.IsNullOrEmpty(text))
                            {
                                debugTerminal?.WriteLine($"       Parágrafo {i}: \"{text.Substring(0, Math.Min(text.Length, 50))}...\"");
                                Author author = ParseAuthor(text);
                                if (author != null)
                                {
                                    articleInfo.Authors.Add(author);
                                    debugTerminal?.WriteLineSuccess($"       ✅ Autor extraído: {author.Nome}");
                                }
                                else
                                {
                                    debugTerminal?.WriteLine($"       ❌ Não foi possível extrair autor deste parágrafo");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                debugTerminal?.WriteLineError($"     ❌ Erro ao processar artigo: {ex.Message}");
                MessageBox.Show($"Erro ao ler artigo {Path.GetFileName(filePath)}: {ex.Message}",
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            return articleInfo;
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

        private DebugTerminalWindow? debugTerminal;
        private const bool SHOW_DEBUG_TERMINAL = true; // ou false, conforme desejado

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