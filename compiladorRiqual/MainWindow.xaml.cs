using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace DocumentUploader
{
    public partial class MainWindow : Window
    {
        // Propriedades públicas para acesso aos ficheiros
        public string File1Path { get; private set; } // Capa
        public string File2Path { get; private set; } // Conselho Editorial
        public string File3Path { get; private set; } // Editorial

        public MainWindow()
        {
            InitializeComponent();
            UpdateStatus();
        }

        private void SelectFile1_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile("Selecionar Capa da Revista");
            if (!string.IsNullOrEmpty(filePath))
            {
                File1Path = filePath;
                txtFile1.Text = $"✅ {Path.GetFileName(filePath)}";
                txtFile1.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                UpdateStatus();
            }
        }

        private void SelectFile2_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile("Selecionar Conselho Editorial");
            if (!string.IsNullOrEmpty(filePath))
            {
                File2Path = filePath;
                txtFile2.Text = $"✅ {Path.GetFileName(filePath)}";
                txtFile2.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                UpdateStatus();
            }
        }

        private void SelectFile3_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile("Selecionar Editorial");
            if (!string.IsNullOrEmpty(filePath))
            {
                File3Path = filePath;
                txtFile3.Text = $"✅ {Path.GetFileName(filePath)}";
                txtFile3.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                UpdateStatus();
            }
        }

        private string SelectDocxFile(string title)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = title,
                Filter = "Documentos Word (*.docx)|*.docx",
                FilterIndex = 1,
                RestoreDirectory = true,
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }
            return null;
        }

        private void UpdateStatus()
        {
            int filesSelected = 0;
            var missingFiles = new List<string>();

            if (!string.IsNullOrEmpty(File1Path))
                filesSelected++;
            else
                missingFiles.Add("Capa");

            if (!string.IsNullOrEmpty(File2Path))
                filesSelected++;
            else
                missingFiles.Add("Conselho Editorial");

            if (!string.IsNullOrEmpty(File3Path))
                filesSelected++;
            else
                missingFiles.Add("Editorial");

            if (filesSelected == 3)
            {
                // Todos os documentos selecionados
                txtStatus.Text = "✅ Todos os documentos foram selecionados!";
                txtStatusDetail.Text = "Pronto para adicionar artigos e compilar a revista";
                statusBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5E8"));
                statusBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                statusIcon.Text = "✅";
                txtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                btnProceed.IsEnabled = true;
            }
            else
            {
                // Documentos em falta
                txtStatus.Text = $"Selecione todos os 3 documentos para prosseguir ({filesSelected}/3)";
                txtStatusDetail.Text = $"Faltam: {string.Join(", ", missingFiles)}";
                statusBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFE6E6"));
                statusBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C"));
                statusIcon.Text = "⚠️";
                txtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C"));
                btnProceed.IsEnabled = false;
            }
        }

        private void Proceed_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Abrir o formulário de compilação passando os ficheiros selecionados
                var compileWindow = new CompileDocumentsWindow(File1Path, File2Path, File3Path);
                compileWindow.Show();

                // Fechar esta janela
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ Erro ao abrir o formulário de compilação:\n{ex.Message}",
                    "Erro",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }
    }
}