using System.Windows;

namespace DocumentUploader
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Configurações iniciais da aplicação podem ser definidas aqui
            // Por exemplo, definir o tema da aplicação, configurações de cultura, etc.
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Limpeza de recursos quando a aplicação termina
            base.OnExit(e);
        }
    }
}