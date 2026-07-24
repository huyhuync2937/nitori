using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;

namespace INCTPND
{
    public partial class FrmINCTPND_PX : Form
    {
        private CodeValueBindingObject Voucher_Ma_nt0;
        public DataRowView drvFrmINCTPND_PX;

        public FrmINCTPND_PX(DataTable tbSource, string ten_vt)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.Loaded += new RoutedEventHandler(this.FrmINCTPND_PX_Loaded);
            this.frmINCTPND_PX.Title = SysFunc.Cat_Dau(ten_vt.ToString());
            this.GrdINCTPND_PX.DataSource = (IEnumerable)tbSource.DefaultView;
        }

        private void FrmINCTPND_PX_Loaded(object sender, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.frmINCTPND_PX.FindResource((object)"Voucher_Ma_nt0");
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            if (this.GrdINCTPND_PX.Records.Count > 0)
                this.GrdINCTPND_PX.ActiveRecord = this.GrdINCTPND_PX.Records[0];
            this.isVisibleField();
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.GrdINCTPND_PX.ActiveRecord == null)
                return;
            this.drvFrmINCTPND_PX = (this.GrdINCTPND_PX.ActiveRecord as DataRecord).DataItem as DataRowView;
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
                this.GrdINCTPND_PX.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                this.GrdINCTPND_PX.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdINCTPND_PX.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Visible;
                this.GrdINCTPND_PX.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = this.GrdINCTPND_PX.FieldLayouts[0].Fields["gia_nt"].Width.Value.Value;
            }
        }
    }
}
