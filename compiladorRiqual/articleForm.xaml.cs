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
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocumentUploader
{
    public partial class CompileDocumentsWindow : Window, IDropTarget, INotifyPropertyChanged
    {
        public ObservableCollection<string> SelectedFiles { get; set; }

        // Ficheiros recebidos do formulário anterior
        private string capaFilePath;
        private string conselhoFilePath;
        private string editorialFilePath;

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
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Ficheiros Word (*.docx)|*.docx",
                Multiselect = true,
                Title = "Selecionar Artigos"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string fileName in openFileDialog.FileNames)
                {
                    SelectedFiles.Add(fileName);
                }
                UpdateStatus($"{openFileDialog.FileNames.Length} artigo(s) adicionado(s)");
                CheckCompileEnabled();
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

                try
                {
                    await Task.Run(() => CreateRevistaDocument(saveFileDialog.FileName));

                    progressTimer.Stop();
                    ProgressValue = 100;
                    await Task.Delay(500);

                    UpdateStatus($"Revista compilada com sucesso!");

                    MessageBox.Show($"Revista guardada com sucesso!\nLocalização: {saveFileDialog.FileName}",
                        "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    progressTimer.Stop();
                    UpdateStatus("Erro: " + ex.Message);
                    MessageBox.Show("Erro ao compilar revista: " + ex.Message,
                        "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    IsCompiling = false;
                    ProgressValue = 0;
                }
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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());
                Body body = mainPart.Document.Body;

                // Define styles
                StyleDefinitionsPart stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new Styles();
                stylesPart.Styles = styles;
                AddCustomStyles(styles);

                // Add header/footer definitions
                AddHeaderFooterDefinitions(mainPart);

                // Extract all authors and article info
                articleAuthors.Clear();
                var allAuthors = new List<Author>();
                var articleInfoList = new List<ArticleInfo>();

                foreach (var article in SelectedFiles)
                {
                    var articleInfo = ExtractArticleInfo(article);
                    if (articleInfo != null)
                    {
                        articleInfoList.Add(articleInfo);
                        articleAuthors[article] = articleInfo.Authors;
                        allAuthors.AddRange(articleInfo.Authors);
                    }
                }

                // Remove duplicates from author list
                allAuthors = allAuthors.GroupBy(a => a.Email ?? a.Nome)
                    .Select(g => g.First())
                    .OrderBy(a => a.Nome)
                    .ToList();

                // Create sections with different headers/footers
                SectionProperties firstSectionProps = new SectionProperties();

                // ORDEM SOLICITADA:
                // 1. Capa
                AddCapa(body, capaFilePath);
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

                // 2. Página em Branco
                AddBlankPage(body);
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

                // 3. Conselho Editorial
                AddConselhoEditorial(body, conselhoFilePath);

                // 4. Lista de Autores
                AddAuthorList(body, allAuthors);

                // 5. Índice
                AddTableOfContents(body);

                // End first section (roman numerals)
                SetupFirstSectionProperties(firstSectionProps, mainPart);
                body.AppendChild(new Paragraph(new ParagraphProperties(firstSectionProps)));

                // 6. Editorial (start of arabic numerals)
                SectionProperties secondSectionProps = new SectionProperties();
                AddEditorial(body, editorialFilePath);

                // 7. Artigos
                foreach (var articleInfo in articleInfoList)
                {
                    AddArticle(body, articleInfo, mainPart);
                }

                // Setup second section properties
                SetupSecondSectionProperties(secondSectionProps, mainPart);
                body.AppendChild(new Paragraph(new ParagraphProperties(secondSectionProps)));

                // Update fields (for TOC)
                AddSettingsToDocument(mainPart);

                mainPart.Document.Save();
            }
        }

        private void AddCapa(Body body, string capaPath)
        {
            if (!string.IsNullOrEmpty(capaPath) && File.Exists(capaPath))
            {
                try
                {
                    using (WordprocessingDocument capaDoc = WordprocessingDocument.Open(capaPath, false))
                    {
                        if (capaDoc.MainDocumentPart != null && capaDoc.MainDocumentPart.Document.Body != null)
                        {
                            // Copiar todo o conteúdo da capa
                            foreach (var element in capaDoc.MainDocumentPart.Document.Body.Elements())
                            {
                                var clonedElement = element.CloneNode(true);

                                // Se houver imagens, precisamos copiá-las também
                                if (element.Descendants<Drawing>().Any())
                                {
                                    CopyImages(capaDoc.MainDocumentPart, body.GetFirstChild<Document>().MainDocumentPart, clonedElement);
                                }

                                body.AppendChild(clonedElement);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Se houver erro, criar uma capa simples
                    AddDefaultCoverPage(body);
                }
            }
            else
            {
                // Capa padrão se o ficheiro não existir
                AddDefaultCoverPage(body);
            }
        }

        private void CopyImages(MainDocumentPart sourcePart, MainDocumentPart targetPart, OpenXmlElement element)
        {
            // Implementação simplificada para copiar imagens
            // Em produção, seria necessário um código mais robusto
            foreach (var drawing in element.Descendants<Drawing>())
            {
                var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                if (blip != null && blip.Embed != null)
                {
                    try
                    {
                        var imagePart = sourcePart.GetPartById(blip.Embed.Value);
                        if (imagePart is ImagePart sourceImagePart)
                        {
                            ImagePart targetImagePart = targetPart.AddImagePart(sourceImagePart.ContentType);
                            using (var stream = sourceImagePart.GetStream())
                            {
                                targetImagePart.FeedData(stream);
                            }
                            blip.Embed = targetPart.GetIdOfPart(targetImagePart);
                        }
                    }
                    catch { }
                }
            }
        }

        private void AddDefaultCoverPage(Body body)
        {
            // Capa padrão
            Paragraph titlePara = new Paragraph();
            titlePara.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "5000" }
            );
            Run titleRun = new Run();
            titleRun.RunProperties = new RunProperties(
                new Bold(),
                new FontSize() { Val = "48" }
            );
            titleRun.Append(new Text("TMQ"));
            titlePara.Append(titleRun);
            body.AppendChild(titlePara);

            // Subtitle
            Paragraph subtitlePara = new Paragraph();
            subtitlePara.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { After = "2000" }
            );
            Run subtitleRun = new Run();
            subtitleRun.RunProperties = new RunProperties(
                new FontSize() { Val = "28" }
            );
            subtitleRun.Append(new Text("TECHNIQUES, METHODOLOGIES AND QUALITY"));
            subtitlePara.Append(subtitleRun);
            body.AppendChild(subtitlePara);

            // Date
            Paragraph datePara = new Paragraph();
            datePara.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            );
            Run dateRun = new Run(new Text(DateTime.Now.ToString("MMMM yyyy", new CultureInfo("pt-PT"))));
            dateRun.RunProperties = new RunProperties(new FontSize() { Val = "24" });
            datePara.Append(dateRun);
            body.AppendChild(datePara);
        }

        private void AddHeaderFooterDefinitions(MainDocumentPart mainPart)
        {
            // Create header parts
            HeaderPart headerPart1 = mainPart.AddNewPart<HeaderPart>();
            HeaderPart headerPartOdd = mainPart.AddNewPart<HeaderPart>();
            HeaderPart headerPartEven = mainPart.AddNewPart<HeaderPart>();

            // Header for first section (all pages até ao editorial)
            Header header1 = new Header();
            Paragraph headerPara1 = new Paragraph();
            headerPara1.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Right }
            );
            headerPara1.Append(new Run(new Text("TMQ – TECHNIQUES, METHODOLOGIES AND QUALITY")));
            header1.Append(headerPara1);
            headerPart1.Header = header1;

            // Headers for articles section
            // Odd pages - article title
            Header headerOdd = new Header();
            Paragraph headerParaOdd = new Paragraph();
            headerParaOdd.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Right }
            );
            headerParaOdd.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            headerParaOdd.Append(new Run(new FieldCode(" STYLEREF \"Heading1\" \\* MERGEFORMAT ")));
            headerParaOdd.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            headerOdd.Append(headerParaOdd);
            headerPartOdd.Header = headerOdd;

            // Even pages - TMQ
            Header headerEven = new Header();
            Paragraph headerParaEven = new Paragraph();
            headerParaEven.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Right }
            );
            headerParaEven.Append(new Run(new Text("TMQ – TECHNIQUES, METHODOLOGIES AND QUALITY")));
            headerEven.Append(headerParaEven);
            headerPartEven.Header = headerEven;

            // Create footer parts
            FooterPart footerPart1 = mainPart.AddNewPart<FooterPart>();
            FooterPart footerPartOdd = mainPart.AddNewPart<FooterPart>();
            FooterPart footerPartEven = mainPart.AddNewPart<FooterPart>();

            // Footer for first section (roman numerals)
            Footer footer1 = new Footer();
            Paragraph footerPara1 = new Paragraph();
            footerPara1.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            );
            footerPara1.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            footerPara1.Append(new Run(new FieldCode(" PAGE \\* ROMAN ")));
            footerPara1.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            footer1.Append(footerPara1);
            footerPart1.Footer = footer1;

            // Footers for articles section
            // Odd pages - page number and authors
            Footer footerOdd = new Footer();

            // Page number
            Paragraph pageNumPara = new Paragraph();
            pageNumPara.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            );
            pageNumPara.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            pageNumPara.Append(new Run(new FieldCode(" PAGE ")));
            pageNumPara.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            footerOdd.Append(pageNumPara);

            // Authors - this would need to be dynamic per article
            Paragraph authorsPara = new Paragraph();
            authorsPara.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center },
                new SpacingBetweenLines() { Before = "120" }
            );
            Run authorsRun = new Run();
            authorsRun.RunProperties = new RunProperties(
                new FontSize() { Val = "18" },
                new Italic()
            );
            authorsRun.Append(new Text(""));  // Will be filled dynamically
            authorsPara.Append(authorsRun);
            footerOdd.Append(authorsPara);

            footerPartOdd.Footer = footerOdd;

            // Even pages - just page number
            Footer footerEven = new Footer();
            Paragraph footerParaEven = new Paragraph();
            footerParaEven.ParagraphProperties = new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            );
            footerParaEven.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));
            footerParaEven.Append(new Run(new FieldCode(" PAGE ")));
            footerParaEven.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
            footerEven.Append(footerParaEven);
            footerPartEven.Footer = footerEven;
        }

        private void SetupFirstSectionProperties(SectionProperties props, MainDocumentPart mainPart)
        {
            props.Append(new PageSize() { Width = 11906, Height = 16838 });
            props.Append(new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 });

            var headerParts = mainPart.HeaderParts.ToList();
            if (headerParts.Count > 0)
            {
                props.Append(new HeaderReference() { Type = HeaderFooterValues.Default, Id = mainPart.GetIdOfPart(headerParts[0]) });
            }

            var footerParts = mainPart.FooterParts.ToList();
            if (footerParts.Count > 0)
            {
                props.Append(new FooterReference() { Type = FooterValues.Default, Id = mainPart.GetIdOfPart(footerParts[0]) });
            }

            props.Append(new PageNumberType() { Format = NumberFormatValues.LowerRoman });
        }

        private void SetupSecondSectionProperties(SectionProperties props, MainDocumentPart mainPart)
        {
            props.Append(new PageSize() { Width = 11906, Height = 16838 });
            props.Append(new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 });

            var headerParts = mainPart.HeaderParts.ToList();
            if (headerParts.Count > 1)
            {
                props.Append(new HeaderReference() { Type = HeaderFooterValues.Odd, Id = mainPart.GetIdOfPart(headerParts[1]) });
                if (headerParts.Count > 2)
                {
                    props.Append(new HeaderReference() { Type = HeaderFooterValues.Even, Id = mainPart.GetIdOfPart(headerParts[2]) });
                }
            }

            var footerParts = mainPart.FooterParts.ToList();
            if (footerParts.Count > 1)
            {
                props.Append(new FooterReference() { Type = FooterValues.Odd, Id = mainPart.GetIdOfPart(footerParts[1]) });
                if (footerParts.Count > 2)
                {