using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System.Collections;
using System.Data;
using System.Windows;

namespace Inctpxd
{
    public partial class FrmInctpxd_Pn : Form
    {
        private CodeValueBindingObject Voucher_Ma_nt0;
        public DataRowView drvFrmINCTPXD_PN;
        public FrmInctpxd_Pn(DataTable tbSource, string ten_vt)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.Loaded += new RoutedEventHandler(this.FrmPoctpxf_PN_Loaded);
            this.Title = SysFunc.Cat_Dau(ten_vt.ToString());
            this.EscToClose = true;
            this.GrdINCTPXD_PN.DataSource = (IEnumerable)tbSource.DefaultView;
            if (StartUpTrans.M_LAN.Equals("V"))
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["ten_ct2"].Visibility = Visibility.Collapsed;
            else
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["ten_ct"].Visibility = Visibility.Collapsed;
        }

        private void FrmPoctpxf_PN_Loaded(object sender, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FindResource((object)"Voucher_Ma_nt0");
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            if (this.GrdINCTPXD_PN.Records.Count > 0)
            {
                this.GrdINCTPXD_PN.Focus();
                this.GrdINCTPXD_PN.ActiveRecord = this.GrdINCTPXD_PN.Records[0];
            }
            this.isVisibleField();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.GrdINCTPXD_PN.ActiveRecord == null)
                return;
            this.drvFrmINCTPXD_PN = (this.GrdINCTPXD_PN.ActiveRecord as DataRecord).DataItem as DataRowView;
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void isVisibleField()
        {
            if (this.Voucher_Ma_nt0.Text.Trim().Equals(StartUpTrans.M_ma_nt0))
            {
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Visible;
                this.GrdINCTPXD_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = this.GrdINCTPXD_PN.FieldLayouts[0].Fields["gia_nt"].Width.Value.Value;
            }
        }
    }
}
