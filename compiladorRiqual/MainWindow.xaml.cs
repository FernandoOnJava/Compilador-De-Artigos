using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace DocumentUploader
{
    public partial class MainWindow : Window
    {
        // Propriedades públicas para acesso aos ficheiros e novos campos
        public string File1Path { get; private set; }
        public string File2Path { get; private set; }
        public string File3Path { get; private set; }
        public string Title { get; private set; }
        public string ISSN { get; private set; }

        public MainWindow()
        {
            InitializeComponent();

            // Inicializar propriedades
            File1Path = string.Empty;
            File2Path = string.Empty;
            File3Path = string.Empty;
            Title = string.Empty;
            ISSN = string.Empty;

            // Call validation and status update after UI is fully loaded
            this.Loaded += (s, e) =>
            {
                UpdateButtonState();
            };
        }

        private void SelectFile1_Click(object sender, RoutedEventArgs e)
        {
            var filePath = SelectDocxFile("Selecionar Capa da Revista");
            if (!string.IsNullOrEmpty(filePath))
            {
                File1Path = filePath;
                txtFile1.Text = $"✅ {Path.GetFileName(filePath)}";
                txtFile1.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
                UpdateButtonState();
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
                UpdateButtonState();
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
                UpdateButtonState();
            }
        }

        private void TxtTitle_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                if (txtTitle != null)
                {
                    Title = txtTitle.Text.Trim();
                    UpdateButtonState();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TxtTitle_TextChanged: {ex.Message}");
            }
        }

        private void TxtISSN_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                if (txtISSN != null)
                {
                    ISSN = txtISSN.Text.Trim();
                    ValidateISSN();
                    UpdateButtonState();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TxtISSN_TextChanged: {ex.Message}");
            }
        }

        private void TxtISSN_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            try
            {
                // Permitir apenas números e hífen
                var regex = new Regex("[^0-9-]");
                e.Handled = regex.IsMatch(e.Text);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TxtISSN_PreviewTextInput: {ex.Message}");
                e.Handled = false; // Allow input if there's an error
            }
        }

        private void ValidateISSN()
        {
            // Check if UI element exists
            if (txtISSNValidation == null)
                return;

            if (string.IsNullOrEmpty(ISSN))
            {
                txtISSNValidation.Text = "";
                txtISSNValidation.Foreground = new SolidColorBrush(Colors.Gray);
                return;
            }

            // Verificar se o formato está correto: xxxx-xxxx
            var issnRegex = new Regex(@"^\d{4}-\d{4}$");

            if (issnRegex.IsMatch(ISSN))
            {
                txtISSNValidation.Text = "✅ Válido";
                txtISSNValidation.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#27AE60"));
            }
            else
            {
                txtISSNValidation.Text = "❌ Inválido";
                txtISSNValidation.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E74C3C"));
            }
        }

        private bool IsISSNValid()
        {
            if (string.IsNullOrEmpty(ISSN))
                return false;

            var issnRegex = new Regex(@"^\d{4}-\d{4}$");
            return issnRegex.IsMatch(ISSN);
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
            return string.Empty;
        }

        // Simplified method to just update the button state
        private void UpdateButtonState()
        {
            try
            {
                // Check if button element exists
                if (btnProceed == null)
                    return;

                // Check if all required fields are completed
                bool allFieldsComplete = !string.IsNullOrEmpty(File1Path) &&
                                       !string.IsNullOrEmpty(File2Path) &&
                                       !string.IsNullOrEmpty(File3Path) &&
                                       !string.IsNullOrEmpty(Title) &&
                                       IsISSNValid();

                btnProceed.IsEnabled = allFieldsComplete;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateButtonState: {ex.Message}");
            }
        }

        private void Proceed_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Verificar se todos os ficheiros existem
                if (string.IsNullOrEmpty(File1Path) || !File.Exists(File1Path))
                {
                    MessageBox.Show("❌ O ficheiro da capa não foi encontrado!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (string.IsNullOrEmpty(File2Path) || !File.Exists(File2Path))
                {
                    MessageBox.Show("❌ O ficheiro do conselho editorial não foi encontrado!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (string.IsNullOrEmpty(File3Path) || !File.Exists(File3Path))
                {
                    MessageBox.Show("❌ O ficheiro do editorial não foi encontrado!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Verificar título
                if (string.IsNullOrEmpty(Title))
                {
                    MessageBox.Show("❌ Por favor, insira o título da revista!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    txtTitle.Focus();
                    return;
                }

                // Verificar ISSN
                if (!IsISSNValid())
                {
                    MessageBox.Show("❌ Por favor, insira um ISSN válido no formato xxxx-xxxx!", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
                    txtISSN.Focus();
                    return;
                }

                // Abrir o formulário de compilação passando todos os parâmetros
                var compileWindow = new CompileDocumentsWindow(File1Path, File2Path, File3Path, Title, ISSN);
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