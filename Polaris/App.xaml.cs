using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Polaris
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private MainWindow main = new MainWindow();
        private Login login = new Login();

        //Set login window = to MainWindow so that the user can terminate the program through the login window
        private void applicationStartup(object sender, StartupEventArgs e)
        {
            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;
            Application.Current.MainWindow = login;

            login.loginSuccessful += main.startupMainWindow;
            login.Show();
        }
    }
}
