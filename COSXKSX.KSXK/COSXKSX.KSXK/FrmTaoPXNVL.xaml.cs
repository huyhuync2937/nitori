using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace COSXKSX.KSXK
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
            this.txtso_ct.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            this.txtOng_ba.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("ong_ba");
            this.txtDien_giai.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("dien_giai");
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMa_qs.Filter = "ma_cts like '%KSX%' and status=1";
            this.txtMa_gd.SearchInit();
            this.txtMa_gd_PreviewLostFocus(this.txtMa_gd, (KeyboardFocusChangedEventArgs)null);
            this.txtMa_kho.SearchInit();
            this.txtMa_kho_PreviewLostFocus(this.txtMa_kho, (KeyboardFocusChangedEventArgs)null);
            this.txtMa_qs.SearchInit();
            this.txtMa_kh.SearchInit();
            this.txtMa_kh_PreviewLostFocus(this.txtMa_kh, (KeyboardFocusChangedEventArgs)null);
            this.txtMa_bpht.SearchInit();
            this.txtMa_bpht_PreviewLostFocus(this.txtMa_bpht, (KeyboardFocusChangedEventArgs)null);

            this.txtMa_nt.IsReadOnly = true;
            this.txtMa_nt.IsTabStop = false;
            this.txtMa_bpht.IsFocus =  true;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            FormTrans owner = this.Owner as FormTrans;
            if (string.IsNullOrEmpty(this.txtMa_bpht.Text.Trim()) || !this.txtMa_bpht.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(1697, StartupBase.SasObj, "Chưa vào mã bpht", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_bpht.IsFocus = true;
                return;
            }
            if (string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()) || !this.txtMa_kh.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(695, StartupBase.SasObj, "Chưa vào mã khách phiếu xuất", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_kh.IsFocus = true;
                return;
            }
            else if (string.IsNullOrEmpty(this.txtMa_kho.Text.Trim()) || !this.txtMa_kho.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(696, StartupBase.SasObj, "Chưa vào mã kho phiếu xuất", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_kho.IsFocus = true;
                return;
            }
            else if (string.IsNullOrEmpty(this.txtMa_qs.Text.Trim()) || !this.txtMa_qs.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(697, StartupBase.SasObj, "Chưa vào mã quyển sổ phiếu xuất", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_qs.IsFocus = true;
                return;
            }
            else if (string.IsNullOrEmpty(this.txtso_ct.Text.Trim()))
            {
                int num = (int)ExMessageBox.Show(698, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtso_ct.Focus();
                return;
            }
            else
            {
                string stt_rec = "";
                if (this.tbInfoPT != null && this.tbInfoPT.Rows.Count == 1)
                    stt_rec = this.tbInfoPT.Rows[0]["stt_rec"].ToString();
                if (owner.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtso_ct.Text.PadLeft(this.txtso_ct.MaxLength, ' '), stt_rec))
                {
                    if (this.txtMa_qs.RowResult["chkso_ct"].ToString().Equals("1"))
                    {
                        if (ExMessageBox.Show(699, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                        {
                            this.txtso_ct.SelectAll();
                            this.txtso_ct.Focus();
                            return;
                        }
                    }
                    else if (this.txtMa_qs.RowResult["chkso_ct"].ToString().Equals("2"))
                    {
                        int num = (int)ExMessageBox.Show(694, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtso_ct.SelectAll();
                        this.txtso_ct.Focus();
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
        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (e.NewFocus == this.GrdOkCancel.pnlButton.btnOk)
                return;

            string strsoct = (this.Owner as FormTrans).GetNewSoct(this.BindingSasObj, this.txtMa_qs.Text);
            if (string.IsNullOrEmpty(this.txtso_ct.Text.Trim()))
            {
                this.txtso_ct.Text = strsoct;
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

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_kh.RowResult == null)
                return;
            this.txtTen_kh.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kh.RowResult["ten_kh"].ToString() : this.txtMa_kh.RowResult["ten_kh2"].ToString();
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

        private void txtMa_bpht_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_bpht.RowResult != null)
            {
                this.txtTen_bpht.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_bpht.RowResult["ten_bpht"].ToString() : this.txtMa_bpht.RowResult["ten_bpht2"].ToString();
            }
            else
            {
                this.txtTen_bpht.Text = "";
            }
        }

        private void txtMa_tudo_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_tudo.RowResult != null)
            {
                this.txtTen_tudo.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_tudo.RowResult["ten_td"].ToString() : this.txtMa_tudo.RowResult["ten_td2"].ToString();
            }
            else
            {
                this.txtTen_tudo.Text = "";
            }
        }

        private void txtMa_tudo2_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_tudo2.RowResult != null)
            {
                this.txtTen_tudo2.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_tudo2.RowResult["ten_td"].ToString() : this.txtMa_tudo2.RowResult["ten_td2"].ToString();
            }
            else
            {
                this.txtTen_tudo2.Text = "";
            }
        }

        private void txtMa_tudo3_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_tudo3.RowResult != null)
            {
                this.txtTen_tudo3.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_tudo3.RowResult["ten_td"].ToString() : this.txtMa_tudo3.RowResult["ten_td2"].ToString();
            }
            else
            {
                this.txtTen_tudo3.Text = "";
            }
        }
    }
}
