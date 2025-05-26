using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace DocumentUploader
{
    public partial class DebugTerminalWindow : Window
    {
        private readonly ScrollViewer scrollViewer;

        public DebugTerminalWindow()
        {
            InitializeComponent();
            UpdateTimestamp();

            // Encontrar o ScrollViewer para auto-scroll
            scrollViewer = FindScrollViewer();
        }

        private ScrollViewer FindScrollViewer()
        {
            // Encontra o ScrollViewer no template
            return txtOutput.Parent as ScrollViewer;
        }

        public void WriteLine(string message)
        {
            Dispatcher.Invoke(() =>
            {
                string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                txtOutput.Text += $"[{timestamp}] {message}\n";

                // Força o scroll para o final
                txtOutput.ScrollToEnd();

                UpdateTimestamp();
            });
        }

        public void WriteLineSuccess(string message)
        {
            WriteLine($"✅ {message}");
        }

        public void WriteLineError(string message)
        {
            WriteLine($"❌ {message}");
        }

        public void WriteLineInfo(string message)
        {
            WriteLine($"ℹ️  {message}");
        }

        public void WriteLineWarning(string message)
        {
            WriteLine($"⚠️  {message}");
        }

        public void WriteSeparator()
        {
            WriteLine("═══════════════════════════════════════════════════════════════");
        }

        public void UpdateStatus(string status)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = status;
                UpdateTimestamp();
            });
        }

        private void UpdateTimestamp()
        {
            txtTimestamp.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        }

        private void ClearOutput_Click(object sender, RoutedEventArgs e)
        {
            txtOutput.Text = "Terminal limpo.\n";
            UpdateTimestamp();
            WriteLine("Terminal reiniciado");
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // Não fechar, apenas esconder
            e.Cancel = true;
            this.Hide();
        }
    }
}