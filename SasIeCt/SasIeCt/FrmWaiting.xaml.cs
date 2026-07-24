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

namespace SasIeCt
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
			lblMessage.Text = ((StartupBase.M_LAN == "V") ? "Đang kiểm tra dữ liệu..." : "Processing...");
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
				int num = 0;
				if (C_ImportVoucher.tb_Post_Error != null)
				{
					num = C_ImportVoucher.tb_Post_Error.Rows.Count;
				}
				PBar.Value = value;
				lblMessage.Text = ((StartupBase.M_LAN == "V") ? "Đang thực hiện...  " : "Processing...  ") + PBar.Value.ToString() + "/" + PBar.Maximum.ToString() + (num.Equals(0) ? "" : (" (Lỗi " + num.ToString() + ")"));
			}
			Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)delegate
			{
			});
		}

		public void Set_Delete(double value)
		{
			mWorker = new BackgroundWorker();
			mWorker.WorkerReportsProgress = true;
			mWorker.WorkerSupportsCancellation = true;
			mWorker.RunWorkerAsync();
			if (!mWorker.CancellationPending)
			{
				PBar.Value = value;
				lblMessage.Text = ((StartupBase.M_LAN == "V") ? "Đang xóa dữ liệu trùng...  " : "Processing...  ") + PBar.Value.ToString() + "/" + PBar.Maximum.ToString();
			}
			Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)delegate
			{
			});
		}

		public void Set_Post_Error(double value)
		{
			mWorker = new BackgroundWorker();
			mWorker.WorkerReportsProgress = true;
			mWorker.WorkerSupportsCancellation = true;
			mWorker.RunWorkerAsync();
			if (!mWorker.CancellationPending)
			{
				PBar.Value = value;
				lblMessage.Text = ((StartupBase.M_LAN == "V") ? "Chương trình hủy số liệu đưa vào...  " : "Processing...  ") + PBar.Value.ToString() + "/" + PBar.Maximum.ToString();
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
