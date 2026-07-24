using SasControls;
using SasLib;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace COSXKSX.KSXS
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
                StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "23.02.32";
                StartUpTrans.Editing_Stt_Rec = ((IEnumerable<string>)e.Args).Count<string>() > 1 ? e.Args[1].ToString() : string.Empty;
                startUp.Run();
            }
        }
    }
}
