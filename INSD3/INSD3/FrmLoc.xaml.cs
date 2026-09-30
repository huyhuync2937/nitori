using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace INSD3
{
    public partial class FrmLoc : FormFilter
    {
        public FrmLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
        }

        private void formFilter_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMaKho.SearchInit();
            this.txtMaVT.SearchInit();
            this.tblTenKho.Text = this.txtMaKho.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaKho.RowResult["ten_kho"].ToString() : this.txtMaKho.RowResult["ten_kho2"].ToString());
            this.tblTenVT.Text = this.txtMaVT.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaVT.RowResult["ten_vt"].ToString() : this.txtMaVT.RowResult["ten_vt2"].ToString());
            this.txtkieu_xem.Focus();
            this.GridSearch.SasObj = StartupBase.SasObj;
            this.GridSearch.tableList = StartUp.tableList;
            this.txtMDVCS.Text = StartUp.M_MA_DVCS;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.validateInput())
                return;
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            StartUp.ngay = this.txtNgay.Text;
            string filter = this.GetFilter();
            string condition = this.GetCondition();
            this.Hide();
            StartUp.CallGridVoucher(true, this.txtNgay.Value, condition + filter, this.txtkho_dl.Text, int.Parse(this.txtkieu_xem.Text));
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void txtMaKho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaKho.RowResult == null)
                this.tblTenKho.Text = "";
            else
                this.tblTenKho.Text = StartupBase.M_LAN.Equals("V") ? this.txtMaKho.RowResult["ten_kho"].ToString() : this.txtMaKho.RowResult["ten_kho2"].ToString();
        }

        private void txtMaVT_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaVT.RowResult == null)
                this.tblTenVT.Text = "";
            else
                this.tblTenVT.Text = StartupBase.M_LAN.Equals("V") ? this.txtMaVT.RowResult["ten_vt"].ToString() : this.txtMaVT.RowResult["ten_vt2"].ToString();
        }

        public bool validateInput()
        {
            if (this.txtNgay.Value == null || string.IsNullOrEmpty(this.txtNgay.Value.ToString()))
            {
                int num = (int)ExMessageBox.Show(1895, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay.Focus();
                return false;
            }
            if (!this.txtNgay.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(1900, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay.Focus();
                return false;
            }
            if (!this.txtMaKho.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(1910, StartupBase.SasObj, "Mã kho không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMaKho.IsFocus = true;
                return false;
            }
            if (this.txtMaVT.CheckLostFocus())
                return true;
            int num1 = (int)ExMessageBox.Show(1915, StartupBase.SasObj, "Mã vật tư không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtMaVT.IsFocus = true;
            return false;
        }

        public string GetFilter()
        {
            string str = "";
            this.GridSearch._GenerateSQLString();
            if (this.GridSearch.arrStrFilter != null && !string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
                str = str + " and " + this.GridSearch.arrStrFilter[0];
            return str;
        }

        public string GetCondition()
        {
            string str = " 1=1 ";
            if (!string.IsNullOrEmpty(this.txtMDVCS.Text.ToString()))
                str = str + " AND ma_kho IN (SELECT ma_kho FROM dmkho WHERE ma_dvcs LIKE '" + this.txtMDVCS.Text.ToString() + "%')";
            if (!string.IsNullOrEmpty(this.txtMaKho.Text.ToString()))
                str = str + " AND ma_kho LIKE '" + this.txtMaKho.Text.ToString() + "%'";
            if (!string.IsNullOrEmpty(this.txtMaVT.Text.ToString()))
                str = str + " AND ma_vt LIKE '" + this.txtMaVT.Text.ToString() + "%'";
            return str;
        }

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
    }
}
