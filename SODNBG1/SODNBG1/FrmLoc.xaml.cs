using SasControls;
using SasDataLib;
using SasDefine;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace SODNBG1
{
    public partial class FrmLoc : Form
    {
        public FrmLoc()
        {
            this.InitializeComponent();
        }
        private void FrmLoc_Loaded(object sender, RoutedEventArgs e)
        {
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtTungay.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            this.txtDenngay.Value = DateTime.Now.Date;
            this.txtTungay.Focus();
            this.txtsodonhang.Filter = " isnull(status2,'') <> 2 and ngay_ct between '" + Convert.ToDateTime(this.txtTungay.Value).ToString("yyyy-MM-dd") + "' and '" + Convert.ToDateTime(this.txtDenngay.Value).ToString("yyyy-MM-dd") + "'";
        }
        private void _confirmGridview_OnOk(object sender, RoutedEventArgs e)
        {
            Checkvalid();
            string filter = " 1=1 ";
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ncc"].ToString()))
            {
                filter += " And v_ph109.ma_hd in (select ma_hd from ct109 where ncc =''" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ncc"].ToString().Trim() + "'')";
            }
            if (!string.IsNullOrEmpty(this.txtsodonhang.Text))
            {
                filter += " And v_ph109.ma_hd like ''" + this.txtsodonhang.Text.Trim() + "%''";
            }
            filter += " And  isnull(v_ph109.status2,'''') <> ''2''";
            FrmLoaddonhang Frmlocdon = new FrmLoaddonhang((DateTime)this.txtTungay.Value, (DateTime)this.txtDenngay.Value, filter);
            Frmlocdon.ShowDialog();
            this.Close();
        }
        bool Checkvalid()
        {
            bool flag = false;
            if (Convert.ToDateTime(txtTungay.Value) > Convert.ToDateTime(txtDenngay.Value))
            {
                flag = true;
                MessageBox.Show("Từ ngày phải <= đến ngày.", "Thông báo");
                this.txtTungay.Focus();
            }
            return flag;
        }
        private void _confirmGridview_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BCBKPBG_PreviewKeyUp(object sender, KeyEventArgs e)
        {

        }

        private void BCBKPBG_Closed(object sender, EventArgs e)
        {

        }

        private void txtsodonhang_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }

        private void txtDenngay_LostFocus(object sender, RoutedEventArgs e)
        {
            this.txtsodonhang.Filter = " isnull(status2,'') <> 2 and ngay_ct between '" + Convert.ToDateTime(this.txtTungay.Value).ToString("yyyy-MM-dd") + "' and '" + Convert.ToDateTime(this.txtDenngay.Value).ToString("yyyy-MM-dd") + "'";
        }

        private void txtTungay_LostFocus(object sender, RoutedEventArgs e)
        {
            this.txtsodonhang.Filter = " isnull(status2,'') <> 2 and ngay_ct between '" + Convert.ToDateTime(this.txtTungay.Value).ToString("yyyy-MM-dd") + "' and '" + Convert.ToDateTime(this.txtDenngay.Value).ToString("yyyy-MM-dd") + "'";
        }
    }
}
