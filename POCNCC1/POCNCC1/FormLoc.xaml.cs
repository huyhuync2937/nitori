using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Windows;
using System.Windows.Input;

namespace POCNCC1
{
    public partial class FormLoc : FormFilter
    {
        public FormLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
        }

        private void TransactionFrm_Loaded(object sender, RoutedEventArgs e)
        {
            DateTime now1 = DateTime.Now;
            this.txtMaKhach.SasObj = this.BindingSasObj;
            this.txtMaKhach.SearchInit();
            this.txtMaKho.SearchInit();
               
            this.tblstatus.Text = "*";
            if (StartupBase.M_LAN.Trim().ToUpper().Equals("V"))
            {
                if (this.txtMaKhach.RowResult != null)
                    this.lblTenKhach.Text = this.txtMaKhach.RowResult["ten_kh"].ToString();
                if (this.txtMaKho.RowResult != null)
                    this.lblTenKho.Text = this.txtMaKho.RowResult["ten_kho"].ToString();
               
            }
            else
            {
                if (this.txtMaKhach.RowResult != null)
                    this.lblTenKhach.Text = this.txtMaKhach.RowResult["ten_kh2"].ToString();
                if (this.txtMaKho.RowResult != null)
                    this.lblTenKho.Text = this.txtMaKho.RowResult["ten_kho2"].ToString();
              
            }
            this.TxtStartDateTime.Focus();
            this.GridSearch.SasObj = this.BindingSasObj;
            this.GridSearch.tableList = "v_ct70hd";
            SysFunc.LoadIcon((Window)this);
            DateTime now2 = DateTime.Now;
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
            if (!this.TxtStartDateTime.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(2460, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtStartDateTime.Focus();
            }
            else if (!this.TxtEndDateTime.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(2465, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else if (this.TxtStartDateTime.Value == null || this.TxtStartDateTime.Value == DBNull.Value)
            {
                int num = (int)ExMessageBox.Show(2470, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtStartDateTime.Focus();
            }
            else if (this.TxtEndDateTime.Value == null || this.TxtEndDateTime.Value == DBNull.Value)
            {
                int num = (int)ExMessageBox.Show(2475, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else if ((DateTime)this.TxtStartDateTime.Value > (DateTime)this.TxtEndDateTime.Value)
            {
                int num = (int)ExMessageBox.Show(2485, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.TxtEndDateTime.Focus();
            }
            else
            {
                StartUp.Trangthai = tblstatus.Text;
                string filterpbg = this.GetFilterPHieubaogia();
                this.TxtStartDateTime.ValueToDisplayTextConverter = this.TxtStartDateTime.ValueToDisplayTextConverter;               
                this.Hide();
                StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value, int.Parse(this.cbMauBaoCao2.Value.ToString()), filterpbg);
            }
        }
       
        public string GetFilterPHieubaogia()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("Ma_hd");
            string str1 = " 1=1 ";
            if (!string.IsNullOrEmpty(this.txtSoCtBatDau.Text))
                str1 = str1 + " and dbo.PadL(RTRIM(" + StartUp.TableViewName + ".ma_hd), " + databaseFieldLength.ToString() + ", ' ') >= '" + this.txtSoCtBatDau.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSoCtkt.Text))
                str1 = str1 + " and dbo.PadL(RTRIM(" + StartUp.TableViewName + ".ma_hd), " + databaseFieldLength.ToString() + ", ' ') <= '" + this.txtSoCtkt.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtMaKhach.Text.Trim()))
                str1 = str1 + " and "+StartUp.TableViewName + ".ma_kh LIKE '" + this.txtMaKhach.Text.Trim() + "%'";
            
            if (!string.IsNullOrEmpty(this.txtMa_dvcs.Text.Trim()))
                str1 = str1 + " and " + StartUp.TableViewName + ".ma_dvcs LIKE '" + this.txtMa_dvcs.Text.Trim() + "%'";          
          
            if (this.tblstatus.Text == "1")
                str1 += " and " + StartUp.TableViewName + ".status = '1'";
            if (this.tblstatus.Text == "2")
                str1 += " and " + StartUp.TableViewName + ".status = '2'";
            return str1;
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
            {
                if (StartupBase.M_LAN.Equals("V"))
                    this.lblTenKhach.Text = this.txtMaKhach.RowResult["ten_kh"].ToString();
                else
                    this.lblTenKhach.Text = this.txtMaKhach.RowResult["ten_kh2"].ToString();
            }
            else
                this.txtMaKhach.Text = "";
        }

        private void TransactionFrm_KeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }
      

        private void txtMaKho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMaKho.Text))
            {
                if (this.txtMaKho.RowResult == null)
                    return;
                this.lblTenKho.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMaKho.RowResult["ten_kho2"].ToString() : this.txtMaKho.RowResult["ten_kho"].ToString();
            }
            else
                this.lblTenKho.Text = string.Empty;
        }

        private void txtSoCtBatDau_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtMaKhach_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMaKhach.Text))
            {
                if (this.txtMaKhach.RowResult == null)
                    return;
                this.lblTenKhach.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMaKhach.RowResult["ten_kh2"].ToString() : this.txtMaKhach.RowResult["ten_kh"].ToString();
            }
            else
                this.lblTenKhach.Text = "";
        }

        private void txtps_ck_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.tblstatus.Text.Trim()))
                return;
            this.tblstatus.Text = "1";
        }

       
    }
}
