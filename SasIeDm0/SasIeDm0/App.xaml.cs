using SasControls;
using SasLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace SasIeDm0
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
                RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj);
                if (StartupBase.SasObj == null)
                {
                    int num = (int)MessageBox.Show("Can not connect to server. Please login again", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    Application.Current.Shutdown();
                }
                else
                {
                    StartUp startUp = new StartUp();
                    StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "01.08.14";
                    startUp.Run();
                    if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                        return;
                    Application.Current.Shutdown();
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
