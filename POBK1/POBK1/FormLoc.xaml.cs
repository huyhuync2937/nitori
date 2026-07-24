using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace POBK1
{
    public partial class FormLoc : FormFilter
    {
        public static readonly DependencyProperty MaDVCSProperty = DependencyProperty.Register("ReadOnlyMaDVCS", typeof(bool), typeof(Window), new PropertyMetadata((object)true));

        public FormLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
        }

        public bool MaDVCS
        {
            get
            {
                return (bool)this.GetValue(FormLoc.MaDVCSProperty);
            }
            set
            {
                this.SetValue(FormLoc.MaDVCSProperty, (object)value);
            }
        }

        private void TransactionFrm_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtDsManx.Text = StartUp.CommandInfo["parameter"].ToString();
            this.txtDsManx.Focus();
            this.GridSearch.SasObj = this.BindingSasObj;
            this.GridSearch.tableList = "v_POBK1";
            SysFunc.LoadIcon((Window)this);
            this.txtMaVT.SearchInit();
            this.txtMaKhach.SearchInit();
            this.txtMaKho.SearchInit();
            this.lblTenKhach.Text = this.txtMaKhach.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaKhach.RowResult["ten_kh"].ToString() : this.txtMaKhach.RowResult["ten_kh2"].ToString());
            this.lblTenKho.Text = this.txtMaKho.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaKho.RowResult["ten_kho"].ToString() : this.txtMaKho.RowResult["ten_kho2"].ToString());
            this.lblTenVT.Text = this.txtMaVT.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaVT.RowResult["ten_vt"].ToString() : this.txtMaVT.RowResult["ten_vt2"].ToString());
        }

        private void btnHuy_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnNhan_Click(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            if (string.IsNullOrEmpty(this.txtDsManx.Text.Trim()))
            {
                int num = (int)ExMessageBox.Show(9, StartupBase.SasObj, "Chưa vào danh sách mã nx (tài khoản có)!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtDsManx.Focus();
            }
            else if (!this.TxtStartDateTime.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(10, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtStartDateTime.Focus();
            }
            else if (!this.TxtEndDateTime.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(15, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else if (this.TxtStartDateTime.Value == null || this.TxtStartDateTime.Value == DBNull.Value)
            {
                int num = (int)ExMessageBox.Show(20, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtStartDateTime.Focus();
            }
            else if (this.TxtEndDateTime.Value == null || this.TxtEndDateTime.Value == DBNull.Value)
            {
                int num = (int)ExMessageBox.Show(25, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else if ((DateTime)this.TxtStartDateTime.Value > (DateTime)this.TxtEndDateTime.Value)
            {
                int num = (int)ExMessageBox.Show(35, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else
            {
                string filter = this.GetFilter();
                this.TxtStartDateTime.ValueToDisplayTextConverter = this.TxtStartDateTime.ValueToDisplayTextConverter;
                StartUp.dtInfo.Rows.Add((object)this.TxtStartDateTime.Text, (object)this.TxtEndDateTime.Text);
                int result;
                if (!int.TryParse(this.cbMauBaoCao.Value.ToString(), out result))
                    result = 1;
                this.Hide();
                string text = this.txtDsManx.Text;
                if (!string.IsNullOrEmpty(this.txtMaVT.Text))
                    StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value, filter, this.txtMaVT.Text.Trim(), int.Parse(this.cbMauBaoCao2.Value.ToString()), result, text);
                else
                    StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value, filter, this.txtMaVT.Text.Trim(), int.Parse(this.cbMauBaoCao2.Value.ToString()), result, text);
            }
        }

        public string GetFilter()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = " and nxt = 1";
            if (!string.IsNullOrEmpty(this.txtSoCtBatDau.Text))
                str = str + " and so_ct >= '" + this.txtSoCtBatDau.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSoCtKetThuc.Text))
                str = str + " and so_ct <= '" + this.txtSoCtKetThuc.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtMaKhach.Text))
                str = str + " and ma_kh LIKE '" + this.txtMaKhach.Text + "%'";
            if (!string.IsNullOrEmpty(this.txtMaVT.Text))
                str = str + " and ma_vt LIKE '" + this.txtMaVT.Text + "%'";
            if (!string.IsNullOrEmpty(this.txtMaKho.Text))
                str = str + " and ma_kho like '" + this.txtMaKho.Text + "%'";
            if (!string.IsNullOrEmpty(this.txtma_dvcs.Text))
                str = str + " and ma_dvcs LIKE '" + this.txtma_dvcs.Text + "%'";
            this.GridSearch._GenerateSQLString();
            if (this.GridSearch.arrStrFilter != null && !string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
                str = str + " and " + this.GridSearch.arrStrFilter[0];
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

        private void txtMaKhach_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaKhach.RowResult != null)
                this.lblTenKhach.Text = StartupBase.M_LAN == "V" ? this.txtMaKhach.RowResult["ten_kh"].ToString() : this.txtMaKhach.RowResult["ten_kh2"].ToString();
            else
                this.lblTenKhach.Text = "";
        }

        private void TransactionFrm_KeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }

        private void txtMaVT_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMaVT.Text))
            {
                if (this.txtMaVT.RowResult == null)
                    return;
                this.lblTenVT.Text = !(StartupBase.M_LAN == "V") ? this.txtMaVT.RowResult["ten_vt2"].ToString() : this.txtMaVT.RowResult["ten_vt"].ToString();
            }
            else
                this.lblTenVT.Text = string.Empty;
        }

        private void txtMaKho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMaKho.Text))
            {
                if (this.txtMaKho.RowResult == null)
                    return;
                this.lblTenKho.Text = !(StartupBase.M_LAN == "V") ? this.txtMaKho.RowResult["ten_kho2"].ToString() : this.txtMaKho.RowResult["ten_kho"].ToString();
            }
            else
                this.lblTenKho.Text = string.Empty;
        }

    }
}
