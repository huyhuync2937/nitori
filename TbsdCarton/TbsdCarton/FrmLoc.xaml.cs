using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace TbsdCarton
{
    public partial class FrmLoc : FormFilter
    {
        public string dau_cuoi = "0";
        public FrmLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
        }

        private void frmLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.GridSearch.SasObj = StartupBase.SasObj;
            this.GridSearch.tableList = StartUp.tableList;
            //this.txtMaKho.SearchInit();
            //this.txtMaKho_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
            this.txtHDTuNgay.Focus();
        }

        public bool validateInput()
        {
            if (string.IsNullOrEmpty(this.txtHDTuNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtHDDenNgay.Text.Trim()) && ((DateTime)this.txtHDDenNgay.Value).Year > 2078)
            {
                int num = (int)ExMessageBox.Show(785, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtHDDenNgay.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(this.txtHDDenNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtHDTuNgay.Text.Trim()) && ((DateTime)this.txtHDTuNgay.Value).Year > 2078)
            {
                int num = (int)ExMessageBox.Show(800, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtHDTuNgay.Focus();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtHDDenNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtHDTuNgay.Text.Trim()))
            {
                if ((DateTime)this.txtHDTuNgay.Value > (DateTime)this.txtHDDenNgay.Value)
                {
                    int num = (int)ExMessageBox.Show(810, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtHDTuNgay.Focus();
                    this.txtHDTuNgay.SelectAll();
                    return false;
                }
                if (((DateTime)this.txtHDTuNgay.Value).Year < 1900 || ((DateTime)this.txtHDTuNgay.Value).Year > 2078)
                {
                    int num = (int)ExMessageBox.Show(815, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtHDTuNgay.Focus();
                    return false;
                }
            }
            //if (string.IsNullOrEmpty(this.txtCTTuNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtCTDenNgay.Text.Trim()) && ((DateTime)this.txtCTDenNgay.Value).Year > 2078)
            //{
            //    int num = (int)ExMessageBox.Show(820, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtCTDenNgay.Focus();
            //    return false;
            //}
            //if (string.IsNullOrEmpty(this.txtCTDenNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtCTTuNgay.Text.Trim()) && ((DateTime)this.txtCTTuNgay.Value).Year > 2078)
            //{
            //    int num = (int)ExMessageBox.Show(835, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtCTTuNgay.Focus();
            //    return false;
            //}
            //if (!string.IsNullOrEmpty(this.txtCTDenNgay.Text.Trim()) && !string.IsNullOrEmpty(this.txtCTTuNgay.Text.Trim()))
            //{
            //    if ((DateTime)this.txtCTTuNgay.Value > (DateTime)this.txtCTDenNgay.Value)
            //    {
            //        int num = (int)ExMessageBox.Show(845, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //        this.txtCTTuNgay.Focus();
            //        this.txtCTTuNgay.SelectAll();
            //        return false;
            //    }
            //    if (((DateTime)this.txtCTTuNgay.Value).Year < 1900 || ((DateTime)this.txtCTTuNgay.Value).Year > 2078)
            //    {
            //        int num = (int)ExMessageBox.Show(850, StartupBase.SasObj, "Năm không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //        this.txtCTTuNgay.Focus();
            //        return false;
            //    }
            //}
            return true;
        }

        //public string getCondition()
        //{
        //    int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("Ma_hd");
        //    string str = " 1=1 ";
        //    if (!string.IsNullOrEmpty(this.txtMaKhach.Text.Trim()))
        //        str = str + " and ma_kh like '" + this.txtMaKhach.Text.Trim() + "%'";
        //    if (!string.IsNullOrEmpty(this.txtSoCtBatDau.Text))
        //        str = str + " and dbo.PadL(RTRIM(ma_hd), " + databaseFieldLength.ToString() + ", ' ') >= '" + this.txtSoCtBatDau.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
        //    if (!string.IsNullOrEmpty(this.txtSoCtkt.Text))
        //        str = str + " and dbo.PadL(RTRIM(ma_hd), " + databaseFieldLength.ToString() + ", ' ') <= '" + this.txtSoCtkt.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
        //    if (!string.IsNullOrEmpty(this.txtMDVCS.Text.Trim()))
        //        str = str + " and ma_dvcs like '" + this.txtMDVCS.Text.Trim() + "%'";
        //    return str;
        //}

        //public string getFilter()
        //{
        //    string str = " 1=1 ";
        //    if (!string.IsNullOrEmpty(this.txtMaKho.Text))
        //        str = str + " and ma_kho Like '" + this.txtMaKho.Text.Trim() + "%'";
        //    if (!string.IsNullOrEmpty(this.txtMavt.Text))
        //        str = str + " and ma_vt Like '" + this.txtMavt.Text.Trim() + "%'";
        //    this.GridSearch._GenerateSQLString();
        //    if (this.GridSearch.arrStrFilter != null && !string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
        //        str = str + " and " + this.GridSearch.arrStrFilter[0];
        //    return str;
        //}

        public string ConvertDataToSql(object value, Type ValueType)
        {
            string str;
            switch (ValueType.ToString())
            {
                case "System.String":
                    str = string.Format("'{0}'", (object)(value as string).Replace("'", "'"));
                    break;
                case "System.DateTime":
                    str = string.Format("'{0}'", (object)((DateTime)value).ToString("yyyyMMdd"));
                    break;
                default:
                    str = string.Format("'{0}'", value);
                    break;
            }
            return str;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!this.validateInput())
                    return;
                //string filter = this.getFilter();
                //string condition = this.getCondition();
                this.Hide();
                if (!string.IsNullOrEmpty(this.txtHDTuNgay.Text))
                    StartUp.hdTuNgay_ = string.Format("{0:dd-MM-yyyy}", this.txtHDTuNgay.Value);
                if (!string.IsNullOrEmpty(this.txtHDDenNgay.Text))
                    StartUp.hdDenNgay_ = string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtHDDenNgay.Value);
                //if (!string.IsNullOrEmpty(this.txtCTTuNgay.Text))
                //    StartUp.ctTuNgay_ = string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtCTTuNgay.Value);
                //if (!string.IsNullOrEmpty(this.txtCTDenNgay.Text))
                //    StartUp.ctDenNgay_ = string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtCTDenNgay.Value);
                StartUp.CallGridReport(true, string.IsNullOrEmpty(this.txtHDTuNgay.Text) ? (object)"" : this.txtHDTuNgay.Value, string.IsNullOrEmpty(this.txtHDDenNgay.Text) ? (object)"" : this.txtHDDenNgay.Value, string.IsNullOrEmpty(this.txtMavt.Text) ? (object)"" : (object)this.txtMavt.Text.Trim());
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void gridMain_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void txtTinhChiPhi_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            int result = -1;
            textBox.SelectAll();
            if (!int.TryParse(textBox.Text, out result))
            {
                TextChange textChange = e.Changes.ElementAt<TextChange>(0);
                int addedLength = textChange.AddedLength;
                int offset = textChange.Offset;
                textBox.Text = textBox.Text.Remove(offset, addedLength);
            }
            else
            {
                int num = result == 0 ? 0 : (result != 1 ? 1 : 0);
                textBox.Text = num != 0 ? this.dau_cuoi : result.ToString();
                this.dau_cuoi = textBox.Text;
            }
        }

        //private void txtMaKhach_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //    if (this.txtMaKhach.RowResult == null)
        //        this.lblTenKhach.Text = "";
        //    else
        //        this.lblTenKhach.Text = !(StartupBase.M_LAN == "V") ? this.txtMaKhach.RowResult["ten_kh2"].ToString().Trim() : this.txtMaKhach.RowResult["ten_kh"].ToString().Trim();
        //}

        //private void txtMaKho_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //    if (this.txtMaKho.RowResult != null)
        //    {
        //        if (StartupBase.M_LAN == "V")
        //            this.lblTenKho.Text = this.txtMaKho.RowResult["ten_kho"].ToString();
        //        else
        //            this.lblTenKho.Text = this.txtMaKho.RowResult["ten_kho2"].ToString();
        //    }
        //    else
        //        this.lblTenKho.Text = string.Empty;
        //}

        //private void txtps_ck_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (!string.IsNullOrEmpty(this.txtps_ck.Text.Trim()))
        //        return;
        //    this.txtps_ck.Text = "1";
        //}

        //private void txtSoCtBatDau_LostFocus(object sender, RoutedEventArgs e)
        //{
        //}

        private void txtMavt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMavt.Text))
            {
                if (this.txtMavt.RowResult == null)
                    return;
                this.lblTenvt.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMavt.RowResult["ten_vt"].ToString() : this.txtMavt.RowResult["ten_vt"].ToString();
            }
            else
                this.lblTenvt.Text = string.Empty;
        }

    }
}
