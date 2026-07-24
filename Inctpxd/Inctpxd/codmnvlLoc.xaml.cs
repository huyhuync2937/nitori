using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace Inctpxd
{
    public partial class codmnvlLoc : FormFilter
    {
        public bool isClose = false;
        private SasFormBrowes.FormBrowse2 oBrowse;

        public codmnvlLoc()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmcodmnvlLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.BindingSasObj = StartupBase.SasObj;
            this.txtSp.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_sp"].ToString();
            this.txtMa_bpht.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_bpht"].ToString();
            this.txtSo_lsx.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsx"].ToString();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            string so_lsx = this.txtSo_lsx.Text.Trim().ToString();
            string ma_sp = this.txtSp.Text.Trim().ToString();
            string ma_bpht = this.txtMa_bpht.Text.Trim().ToString();
            DateTime? ngay_ct1 = new DateTime?();
            DateTime? ngay_ct2 = new DateTime?();
            if (this.TxtNgay_ct1.IsValueValid && this.TxtNgay_ct1.Value != null)
                ngay_ct1 = new DateTime?(this.TxtNgay_ct1.dValue);
            if (this.TxtNgay_ct2.IsValueValid && this.TxtNgay_ct2.Value != null)
                ngay_ct2 = new DateTime?(this.TxtNgay_ct2.dValue);
            this.CallGridVouchers(so_lsx, ma_sp, ma_bpht, ngay_ct1, ngay_ct2);
        }

        public void CallGridVouchers(
          string so_lsx,
          string ma_sp,
          string ma_bpht,
          DateTime? ngay_ct1,
          DateTime? ngay_ct2)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("[INCTPXD-CODMNVL]");
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Parameters.Add("@so_lsx", SqlDbType.VarChar).Value = (object)so_lsx;
                sqlcmd.Parameters.Add("@Ma_sp", SqlDbType.VarChar).Value = (object)ma_sp;
                sqlcmd.Parameters.Add("@Ma_bpht", SqlDbType.VarChar).Value = (object)ma_bpht;
                sqlcmd.Parameters.Add("@Ngay_ct1", SqlDbType.SmallDateTime).Value = !ngay_ct1.HasValue ? (object)DBNull.Value : (object)ngay_ct1;
                sqlcmd.Parameters.Add("@Ngay_ct2", SqlDbType.SmallDateTime).Value = !ngay_ct2.HasValue ? (object)DBNull.Value : (object)ngay_ct2;
                sqlcmd.Parameters.Add("@Ngay_ct", SqlDbType.SmallDateTime).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                DataTable table1 = dataSet.Tables[0];
                DataTable table2 = dataSet.Tables[1];
                table1.TableName = "tbMain";
                table2.TableName = "tbDetail";
                new FrmView(table1, table2).ShowDialog();
                this.Close();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtSo_lsx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtSo_lsx.SearchInit();
            this.lblTenSo_lsx.Text = this.txtSo_lsx.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtSo_lsx.RowResult["ten_lsx"].ToString() : this.txtSo_lsx.RowResult["ten_lsx2"].ToString());
            if (this.txtSo_lsx.RowResult == null)
                return;
            this.lblNgayLsx.Value = this.txtSo_lsx.RowResult["ngay_lkh"];
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isClose = true;
            this.Close();
        }

        private void txtSp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtSp.SearchInit();
            this.lblTenSp.Text = this.txtSp.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtSp.RowResult["ten_vt"].ToString() : this.txtSp.RowResult["ten_vt2"].ToString());
        }

        private void txtbpht_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtMa_bpht.SearchInit();
            this.lblTen_bpht.Text = this.txtMa_bpht.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMa_bpht.RowResult["ten_bpht"].ToString() : this.txtMa_bpht.RowResult["ten_bpht2"].ToString());
        }
    }
}
