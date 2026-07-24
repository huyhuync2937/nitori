using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace QCSN
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
           
            this.GridSearch.SasObj = this.BindingSasObj;
            this.GridSearch.tableList = "v_QCSN";
            SysFunc.LoadIcon((Window)this);
            this.txtMaVT.SearchInit();
         
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
               
                if (!string.IsNullOrEmpty(this.txtMaVT.Text))
                    StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value,  this.txtMaVT.Text.Trim());
                else
                    StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value,  this.txtMaVT.Text.Trim());
            }
        }

        public string GetFilter()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = " and nxt = 1";
            
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

        

    }
}
