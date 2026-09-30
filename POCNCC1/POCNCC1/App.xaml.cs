using SasControls;
using SasLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace POCNCC1
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            try
            {
                AppDomain.CurrentDomain.SetupInformation.LoaderOptimization = LoaderOptimization.MultiDomain;
                RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj);
            }
            catch (Exception ex)
            {
            }
            if (StartupBase.SasObj == null)
            {
                int num = (int)MessageBox.Show("Can not connect to server. Please login again", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                Application.Current.Shutdown(0);
            }
            else
            {
                StartUp startUp = new StartUp();
                StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "23.01.05";
                startUp.Run();
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                this.Shutdown(0);
            }
        }
    }
}
