using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace Socthda
{
    public partial class FrmTaoPXNVL : Form
    {
        public bool isOk = false;
        public int kind = 1;
        public string Ma_nt_ht = "";
        public string so_hd = "";
        public string ngay_hd = "";
        public DataTable tbInfoPT = (DataTable)null;
        public FrmTaoPXNVL()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtso_ct_pt.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            this.txtnguoi_nop.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("ong_ba");
            this.txtlydo_nop.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("dien_giai");
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMa_qs_pt.Filter = "ma_cts like '%PXD%' and status=1 \r\n                                        AND ((NOT EXISTS(SELECT 1 FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs) \r\n                                        OR EXISTS (SELECT ma_qs FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs \r\n                                        AND [user_id] = " + StartUpTrans.M_User_Id + ")))";
            if (this.tbInfoPT == null || this.tbInfoPT != null && this.tbInfoPT.Rows.Count == 0)
            {
                this.txtKind.Value = this.kind;
                this.txtMa_gd.Text = "4";
                this.txtMa_gd.SearchInit();
                this.txtMa_gd_PreviewLostFocus(this.txtMa_gd, (KeyboardFocusChangedEventArgs)null);
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "if(select count(1) from dmqs where ma_cts LIKE '%{0}%') = 1";
                sqlcmd.CommandText += " select ma_qs from dmqs where ma_cts LIKE '%PXD%'";
                DataSet dataSet = this.BindingSasObj.ExcuteReader(sqlcmd);
                if (dataSet.Tables.Count == 1)
                    this.txtMa_qs_pt.Text = dataSet.Tables[0].Rows[0]["ma_qs"].ToString();
                this.txtlydo_nop.Text = !StartUpTrans.M_LAN.Equals("V") ? string.Format("Export NVL {0}, invoice date {1}", this.so_hd, this.ngay_hd) : string.Format("Xuất NVL hóa đơn số {0}, ngày {1}", this.so_hd, this.ngay_hd);
            }
            else
            {
                this.txtMa_gd.Text = this.tbInfoPT.Rows[0]["ma_gd"].ToString();
                this.txtMa_gd.SearchInit();
                this.txtMa_gd_PreviewLostFocus(this.txtMa_gd, (KeyboardFocusChangedEventArgs)null);
                this.txtMa_qs_pt.Text = this.tbInfoPT.Rows[0]["ma_qs"].ToString();
                this.txtso_ct_pt.Text = this.tbInfoPT.Rows[0]["so_ct"].ToString().Trim();
                this.txtnguoi_nop.Text = this.tbInfoPT.Rows[0]["ong_ba"].ToString();
                this.txtlydo_nop.Text = this.tbInfoPT.Rows[0]["dien_giai"].ToString();
                this.txtMa_kho.Text = this.tbInfoPT.Rows[0]["ma_kho"].ToString();
                this.txtMa_kho.SearchInit();
                this.txtMa_kho_PreviewLostFocus(this.txtMa_kho, (KeyboardFocusChangedEventArgs)null);
            }
            this.txtMa_nt.Text = this.Ma_nt_ht;
            if (this.Ma_nt_ht != StartupBase.M_MA_NT0)
            {
                this.txtMa_nt.Filter = "ma_nt IN ('" + StartupBase.M_MA_NT0 + "','" + this.Ma_nt_ht + "')";
            }
            else
            {
                this.txtMa_nt.IsReadOnly = true;
                this.txtMa_nt.IsTabStop = false;
            }
            this.txtKind.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            FormTrans owner = this.Owner as FormTrans;
            if (string.IsNullOrEmpty(this.txtMa_kho.Text.Trim()) || !this.txtMa_kho.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(696, StartupBase.SasObj, "Chưa vào mã kho phiếu xuất", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_kho.IsFocus = true;
            }
            else if (string.IsNullOrEmpty(this.txtMa_qs_pt.Text.Trim()) || !this.txtMa_qs_pt.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(697, StartupBase.SasObj, "Chưa vào mã quyển sổ phiếu xuất", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_qs_pt.IsFocus = true;
            }
            else if (string.IsNullOrEmpty(this.txtso_ct_pt.Text.Trim()))
            {
                int num = (int)ExMessageBox.Show(698, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtso_ct_pt.Focus();
            }
            else
            {
                string stt_rec = "";
                if (this.tbInfoPT != null && this.tbInfoPT.Rows.Count == 1)
                    stt_rec = this.tbInfoPT.Rows[0]["stt_rec"].ToString();
                if (owner.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text, this.txtso_ct_pt.Text.PadLeft(this.txtso_ct_pt.MaxLength, ' '), stt_rec))
                {
                    if (this.txtMa_qs_pt.RowResult["chkso_ct"].ToString().Equals("1"))
                    {
                        if (ExMessageBox.Show(699, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                        {
                            this.txtso_ct_pt.SelectAll();
                            this.txtso_ct_pt.Focus();
                            return;
                        }
                    }
                    else if (this.txtMa_qs_pt.RowResult["chkso_ct"].ToString().Equals("2"))
                    {
                        int num = (int)ExMessageBox.Show(694, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text).Trim() + "]", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtso_ct_pt.SelectAll();
                        this.txtso_ct_pt.Focus();
                        return;
                    }
                }
                this.isOk = true;
                this.Close();
            }
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            FormTrans owner = this.Owner as FormTrans;
            this.isOk = false;
            this.Close();
        }
        private void txtMa_qs_pt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (e.NewFocus == this.GrdOkCancel.pnlButton.btnOk)
                return;
            //this.txtso_ct_pt.Text = (this.Owner as FormTrans).GetNewSoct(this.BindingSasObj, this.txtMa_qs_pt.Text);

            string strsoct = (this.Owner as FormTrans).GetNewSoct(this.BindingSasObj, this.txtMa_qs_pt.Text);
            if (string.IsNullOrEmpty(this.txtso_ct_pt.Text.Trim()))
            {
                this.txtso_ct_pt.Text = strsoct;
            }
        }

        public string GetNewSoct(string ma_qs)
        {
            string str1;
            DataTable table = this.BindingSasObj.ExcuteReader(new SqlCommand(Convert.ToInt16(this.BindingSasObj.GetOption("M_AUTO_SOCT").ToString()) == (short)1 ? "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'" : (str1 = "EXEC  [GetNewSoct] '" + ma_qs.Trim() + "'"))).Tables[0];
            if (table.Rows.Count > 0)
            {
                DataRow row = table.Rows[0];
                if (row[1] != null && row[1] != DBNull.Value)
                {
                    string str2 = row[1].ToString();
                    return string.Format(row[0].ToString(), Convert.ToDouble(str2));
                }
            }
            return "";
        }

        private void txtKind_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtKind.Text))
                this.txtKind.Value = this.kind;
        }

        private void txtMa_gd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_gd.RowResult == null)
                return;
            this.txtTen_gd.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_gd.RowResult["ten_gd"].ToString() : this.txtMa_gd.RowResult["ten_gd2"].ToString();
        }
        private void txtMa_kho_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_kho.RowResult == null)
                return;
            this.txtTen_kho.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kho.RowResult["ten_kho"].ToString() : this.txtMa_kho.RowResult["ten_kho2"].ToString();
        }

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_nx.RowResult != null)
            {
                this.tblTenNX.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_nx.RowResult["ten_nx"].ToString() : this.txtMa_nx.RowResult["ten_nx2"].ToString();
            }
            else
            {
                this.tblTenNX.Text = "";
            }
        }
    }
}
