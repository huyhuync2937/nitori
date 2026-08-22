using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Threading;
using System.Windows.Threading;
using SasControls;
using SasFormBrowes;

namespace SasIeDm
{   
    public partial class FrmWaiting : Form
    {
        private BackgroundWorker mWorker;
		public double pgValue
		{
			get
			{
				return PBar.Maximum;
			}
			set
			{
				PBar.Maximum = value;
			}
		}

		public FrmWaiting(double maximum)
		{
			InitializeComponent();
			base.ShowInTaskbar = false;
			PBar.Maximum = maximum;
			PBar.Value = 0.0;
			base.Topmost = true;
			SysFunc.LoadIcon(this);
			base.ActiveMainForm = false;
			lblMessage.Text = ((StartupBase.M_LAN == "V") ? "Đang thực hiện..." : "Processing...");
			txtVer.Text = StartupBase.SasObj.VersionInfo.Rows[0]["product"].ToString();
		}

		public void Set(double value)
		{
			mWorker = new BackgroundWorker();
			mWorker.WorkerReportsProgress = true;
			mWorker.WorkerSupportsCancellation = true;
			mWorker.RunWorkerAsync();
			if (!mWorker.CancellationPending)
			{
				PBar.Value = value;
			}
			Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)delegate
			{
			});
		}

		protected override void OnClosed(EventArgs e)
		{
		}

	}
}
