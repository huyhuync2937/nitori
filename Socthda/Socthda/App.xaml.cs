using SasControls;
using SasVoucherLib;
using SasLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Diagnostics;

namespace Socthda
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj); 
            Thread.CurrentThread.CurrentCulture = StartupBase.SasObj.SasCultureInfo;
            if (StartupBase.SasObj == null)
            {
                int num = (int)MessageBox.Show("Can not connect to server. Please login again", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                Application.Current.Shutdown();
            }
            StartUp startUp = new StartUp();
            StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "21.02.05";
            StartUpTrans.Editing_Stt_Rec = ((IEnumerable<string>)e.Args).Count<string>() > 1 ? e.Args[1].ToString() : string.Empty;
            startUp.Run();
            if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                return;
            Application.Current.Shutdown();
        }
    }
}
