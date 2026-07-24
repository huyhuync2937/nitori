using SasControls;
using SasLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace Intc
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj);
            StartUp startUp = new StartUp();
            StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "50.08.05";
            startUp.Run();
        }
    }
}
