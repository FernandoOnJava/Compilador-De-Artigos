using System.IO;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace DocumentUploader
{
    public partial class MainWindow : Window
    {
        // Propriedades públicas para acesso aos ficheiros e pasta temporária
        public string TempFolderPath { get; private set; }
        public string File1Path { get; private set; }
        public string File2Path { get; private set; }
        public string File3Path { get; private set; }

        // Caminhos originais dos ficheiros selecionados
        private string _originalFile1Path;
        private string _originalFile2Path;
        private string _originalFile3Path;

        public MainWindow()
        {
            InitializeComponent();
            UpdateStatus();
        }

        private void SelectFile1_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                _originalFile1Path = filePath;
                txtFile1.Text = Path.GetFileName(filePath);
                txtFile1.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2C3E50"));
                UpdateStatus();
            }
        }

        private void SelectFile2_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                _originalFile2Path = filePath;
                txtFile2.Text = Path.GetFileName(filePath);
                txtFile2.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2C3E50"));
                UpdateStatus();
            }
        }

        private void SelectFile3_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                _originalFile3Path = filePath;
                txtFile3.Text = Path.GetFileName(filePath);
                txtFile3.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2C3E50"));
                UpdateStatus();
            }
        }

        private string SelectDocxFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Selecionar Documento DOCX",
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

            if (!string.IsNullOrEmpty(_originalFile1Path))
                filesSelected++;
            else
                missingFiles.Add("Capa");

            if (!string.IsNullOrEmpty(_originalFile2Path))
                filesSelected++;
            else
                missingFiles.Add("Conselho Editorial");

            if (!string.IsNullOrEmpty(_originalFile3Path))
                filesSelected++;
            else
                missingFiles.Add("Editorial");

            if (filesSelected == 3)
            {
                // Todos os documentos selecionados
                txtStatus.Text = "✅ Todos os documentos obrigatórios foram selecionados!";
                txtStatusDetail.Text = "Pronto para criar sua revista digital";
                statusBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F5E8"));
                statusBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                statusIcon.Text = "✅";
                txtStatus.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                btnProceed.IsEnabled = true;
            }
            else
            {
                // Documentos em falta
                txtStatus.Text = $"Selecione todos os 3 documentos obrigatórios para prosseguir ({filesSelected}/3)";
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
                // Criar pasta temporária
                CreateTempFolder();

                // Copiar ficheiros para a pasta temporária
                CopyFilesToTempFolder();

                // Mostrar mensagem de sucesso
                MessageBox.Show(
                    $"Ficheiros processados com sucesso!\n\nPasta temporária criada em:\n{TempFolderPath}\n\nOs ficheiros estão prontos para utilização.",
                    "Sucesso",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );

                // Aqui podes abrir outro formulário ou fazer outras operações
                // Exemplo: 
                // var nextWindow = new NextWindow(this.TempFolderPath);
                // nextWindow.Show();
                // this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Erro ao processar os ficheiros:\n{ex.Message}",
                    "Erro",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }

        private void CreateTempFolder()
        {
            // Criar pasta temporária única
            string tempBasePath = Path.GetTempPath();
            string folderName = $"DocumentUploader_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N")[..8]}";
            TempFolderPath = Path.Combine(tempBasePath, folderName);

            Directory.CreateDirectory(TempFolderPath);
        }

        private void CopyFilesToTempFolder()
        {
            if (!string.IsNullOrEmpty(_originalFile1Path))
            {
                string fileName = Path.GetFileName(_originalFile1Path);
                File1Path = Path.Combine(TempFolderPath, $"Documento1_{fileName}");
                File.Copy(_originalFile1Path, File1Path, true);
            }

            if (!string.IsNullOrEmpty(_originalFile2Path))
            {
                string fileName = Path.GetFileName(_originalFile2Path);
                File2Path = Path.Combine(TempFolderPath, $"Documento2_{fileName}");
                File.Copy(_originalFile2Path, File2Path, true);
            }

            if (!string.IsNullOrEmpty(_originalFile3Path))
            {
                string fileName = Path.GetFileName(_originalFile3Path);
                File3Path = Path.Combine(TempFolderPath, $"Documento3_{fileName}");
                File.Copy(_originalFile3Path, File3Path, true);
            }
        }

        // Método para limpeza da pasta temporária (opcional)
        public void CleanupTempFolder()
        {
            try
            {
                if (!string.IsNullOrEmpty(TempFolderPath) && Directory.Exists(TempFolderPath))
                {
                    Directory.Delete(TempFolderPath, true);
                }
            }
            catch (Exception ex)
            {
                // Log do erro se necessário
                System.Diagnostics.Debug.WriteLine($"Erro ao limpar pasta temporária: {ex.Message}");
            }
        }

        // Chamado quando a janela é fechada
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            // Opcional: limpar pasta temporária ao fechar
            // CleanupTempFolder();
        }
    }
}