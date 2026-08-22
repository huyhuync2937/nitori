using SasControls;
using SasLib;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace SasIeDm
{
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            RemotingClient.InitClientRemoteObject(ref StartupBase.SasObj);
            StartUp startUp = new StartUp();
            StartupBase.Menu_Id = ((IEnumerable<string>)e.Args).Count<string>() > 0 ? e.Args[0].ToString() : "01.08.14";
            startUp.Run();
        }
    }
}
