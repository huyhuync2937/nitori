using SasControls;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Threading;

namespace SasVoucherLib
{
    /// <summary>Interaction logic for Window1.xaml</summary>
    /// <summary>SelectTime</summary>
    public partial class SelectTime : FormFilter
    {
        private bool bResult;

        /// <summary>Ctor</summary>
        public SelectTime()
        {
            this.InitializeComponent();
        }

        private void LoadNgay_ct()
        {
            object obj = StartupBase.SasObj.ExcuteScalar(new SqlCommand("SELECT ngay_ks FROM dmct WHERE ma_ct = '" + StartUpTrans.Ma_ct + "'"));
            if (obj == null || obj == DBNull.Value)
                obj = StartupBase.SasObj.GetSysvar("M_NGAY_KS");
            this.txtNgay_ct1.Value = (object)((DateTime)obj).AddDays(1.0);
            this.txtNgay_ct2.Value = (object)DateTime.Today.Date;
        }

        private void FrmSelectTime_Loaded(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtNgay_ct1.Focus()), DispatcherPriority.Background);
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct1.Value is DBNull || this.txtNgay_ct1.Value == null || (!this.txtNgay_ct1.IsValueValid || this.txtNgay_ct1.Value.ToString() == ""))
            {
                int num = (int)ExMessageBox.Show(-1165, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
            }
            else if (this.txtNgay_ct2.Value is DBNull || this.txtNgay_ct2.Value == null || (!this.txtNgay_ct2.IsValueValid || this.txtNgay_ct2.Value.ToString() == ""))
            {
                int num = (int)ExMessageBox.Show(-1165, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct2.Focus();
            }
            else if ((DateTime)this.txtNgay_ct1.Value > (DateTime)this.txtNgay_ct2.Value)
            {
                int num = (int)ExMessageBox.Show(-1165, StartupBase.SasObj, "Chứng từ từ ngày phải nhỏ hơn đến ngày!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct2.Focus();
            }
            else
            {
                this.bResult = true;
                this.Close();
            }
        }

        public bool IsOK
        {
            get
            {
                return this.bResult;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            if (this.IsOK)
                return;
            Application.Current.Shutdown();
        }

    }
}
