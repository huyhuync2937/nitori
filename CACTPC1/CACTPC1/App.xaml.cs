using SasLib;
using SasControls;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Threading;

namespace CACTPC1
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
            }
            StartUp startUp = new StartUp();
            StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "04.01.04";
            StartUpTrans.Editing_Stt_Rec = ((IEnumerable<string>)e.Args).Count<string>() > 1 ? e.Args[1].ToString() : string.Empty;
            startUp.Run();
            this.Shutdown();
        }
    }
}
