using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace COSXLSX.CODMNVL
{
    public partial class codmnvlLoc : FormFilter
    {
        public bool isClose = false;
        DateTime dateTime1 = DateTime.MinValue;
        DateTime dateTime2 = DateTime.MinValue;
        public codmnvlLoc()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmcodmnvlLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.BindingSasObj = StartupBase.SasObj;
            this.txtMa_ky.IsFocus = true;
            //this.grd.pnlButton.btnOk.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            //if (this.txtpass_word.Password != "Nitori@123")
            //{
            //    int num = (int)ExMessageBox.Show( StartupBase.SasObj, "Sai mật khẩu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtpass_word.Focus();
            //    return;
            //}
            if (this.txtMa_ky.Text.Trim() == string.Empty)
            {
                int num = (int)ExMessageBox.Show(223, StartupBase.SasObj, "Chưa vào kỳ giá thành!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_ky.IsFocus = true;
                return;
            }
            else if (!this.txtMa_ky.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(225, StartupBase.SasObj, "Kỳ giá thành không tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_ky.IsFocus = true;
                return;
            }
            else
            {
                if (this.txtMa_ky.RowResult != null)
                {
                    this.dateTime1 = this.txtMa_ky.RowResult["ngay1"] == null ? DateTime.MinValue : (DateTime)this.txtMa_ky.RowResult["ngay1"];
                    this.dateTime2 = this.txtMa_ky.RowResult["ngay2"] == null ? DateTime.MinValue : (DateTime)this.txtMa_ky.RowResult["ngay2"];
                }
                StartUp.sso_lsx_loc = this.txtSo_lsx.Text.Trim().ToString();
                StartUp.sMa_sp_loc = this.txtSp.Text.Trim().ToString();
                StartUp.sMa_bpht_loc = this.txtMa_bpht.Text.Trim().ToString();
                //StartUp.sMa_px_loc = this.txtMa_px.Text.Trim().ToString();

                StartUp.sMa_ky_loc = this.txtMa_ky.Text.Trim().ToString();
                StartUp.sMa_hd_loc = this.txtma_hd.Text.Trim().ToString();
                StartUp.sNgay1_loc = this.dateTime1;
                StartUp.sNgay2_loc = this.dateTime2;
                this.Hide();
                StartUp.CallGridVouchers(true, StartUp.sMa_ky_loc, StartUp.sMa_px_loc);
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
        //private void txtpx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //    this.txtMa_px.SearchInit();
        //    this.lblTen_px.Text = this.txtMa_px.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMa_px.RowResult["ten_px"].ToString() : this.txtMa_px.RowResult["ten_px2"].ToString());
        //}
        private void txtMa_ky_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_ky.RowResult != null)
                this.tblTen_ky.Text = StartupBase.M_LAN.Equals("V") ? this.txtMa_ky.RowResult["ten_ky"].ToString() : this.txtMa_ky.RowResult["ten_ky2"].ToString();
            else
                this.tblTen_ky.Text = "";
        }
        private void txtma_hd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtma_hd.RowResult != null)
                this.lblNgayCt.Value = this.txtma_hd.RowResult["ngay_ct"];
            else
                this.lblNgayCt.Value = null;
        }
    }
}
