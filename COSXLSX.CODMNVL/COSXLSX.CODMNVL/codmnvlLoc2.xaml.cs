using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace COSXLSX.CODMNVL
{
    public partial class codmnvlLoc2 : FormFilter
    {
        public bool isClose = false;
        DateTime dateTime1 = DateTime.MinValue;
        DateTime dateTime2 = DateTime.MinValue;
        public codmnvlLoc2()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmcodmnvlLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.BindingSasObj = StartupBase.SasObj;
            this.txtMa_ky.Text = "";
            this.txtMa_bpht.Text = StartUp.sMa_bpht_loc;
            this.txtSo_lsx.Text = StartUp.sso_lsx_loc;
            this.txtMa_ky.IsFocus = true;

        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            if (this.txtMa_ky.Text.Trim() == StartUp.sMa_ky_loc)
            {
                int num = (int)ExMessageBox.Show(221, StartupBase.SasObj, "Chọn kỳ giá thành không được trùng với kỳ giá thành điều kiện lọc vào!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_ky.IsFocus = true;
            }
            else if (this.txtMa_ky.Text.Trim() == string.Empty)
            {
                int num = (int)ExMessageBox.Show(223, StartupBase.SasObj, "Chưa vào kỳ giá thành!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_ky.IsFocus = true;
            }
            else if (!this.txtMa_ky.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(225, StartupBase.SasObj, "Kỳ giá thành không tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_ky.IsFocus = true;
            }
            else
            {
                if (this.txtMa_ky.RowResult != null)
                {
                    this.dateTime1 = this.txtMa_ky.RowResult["ngay1"] == null ? DateTime.MinValue : (DateTime)this.txtMa_ky.RowResult["ngay1"];
                    this.dateTime2 = this.txtMa_ky.RowResult["ngay2"] == null ? DateTime.MinValue : (DateTime)this.txtMa_ky.RowResult["ngay2"];
                }
                StartUp.ma_ky_truoc = this.txtMa_ky.Text.Trim();
                StartUp.so_lsx_truoc = this.txtSo_lsx.Text.Trim();
                StartUp.Ma_sp_truoc = this.txtSp.Text.Trim();
                StartUp.ma_hd_truoc = this.txtMa_hd.Text.Trim();
                StartUp.ma_bpht_truoc = this.txtMa_bpht.Text.Trim();
                this.Close();
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

        private void txtMa_ky_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_ky.RowResult != null)
                this.tblTen_ky.Text = StartupBase.M_LAN.Equals("V") ? this.txtMa_ky.RowResult["ten_ky"].ToString() : this.txtMa_ky.RowResult["ten_ky2"].ToString();
            else
                this.tblTen_ky.Text = "";
        }
        private void txtMa_hd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtMa_hd.SearchInit();
            if (this.txtMa_hd.RowResult == null)
            {
                this.lblNgayCt.Value = DBNull.Value;
            }
            else
            this.lblNgayCt.Value = this.txtMa_hd.RowResult["ngay_ct"];
        }
    }
}
