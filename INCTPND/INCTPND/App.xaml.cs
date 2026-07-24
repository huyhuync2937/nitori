using SasControls;
using SasLib;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace INCTPND
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj);
            if (StartupBase.SasObj == null)
            {
                int num = (int)MessageBox.Show("Can not connect to server. Please login again", "Connection Error", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            StartUp startUp = new StartUp();
            StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "25.02.01";
            StartUpTrans.Editing_Stt_Rec = ((IEnumerable<string>)e.Args).Count<string>() > 1 ? e.Args[1].ToString() : string.Empty;
            startUp.Run();
        }
    }
}
