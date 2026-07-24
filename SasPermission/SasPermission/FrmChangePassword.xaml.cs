using SasControls;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace SasPermission
{
    public partial class FrmChangePassword : FormFilter
    {
        public int rowIndex = -1;
        public int M_LEN_PASS = 0;
        private DataTable OldRow = (DataTable)null;
        private EditModeBindingObject FormInEditMode;
        private DataTable newDataTable;

        public FrmChangePassword()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.FormInEditMode = (EditModeBindingObject)this.FindResource((object)"IsInEditMode");
        }

        private void LoadForm()
        {
        }

        private void FrmChangePassword_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                int.TryParse(StartupBase.SasObj.GetOption("M_LEN_PASS").ToString().Trim(), out M_LEN_PASS);
                M_LEN_PASS = M_LEN_PASS < 0 ? 0 : M_LEN_PASS;

                this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                if (this.newDataTable.Rows.Count > 0)
                {
                    this.OldRow = this.newDataTable.Copy();
                    this.txtCurPass.IsEnabled = (int)Convert.ToInt16(this.OldRow.Rows[0]["user_id"]) == StartUp.user_id;
                    this.newDataTable.Rows[0]["user_name"] = (object)this.newDataTable.Rows[0]["user_name"].ToString().Trim();
                    if (this.txtCurPass.IsEnabled)
                        this.txtCurPass.Focus();
                    else
                        this.txtpassword.Focus();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            this.gridMain.DataContext = (object)this.newDataTable;
        }

        private bool CheckCurPass()
        {
            if (!this.txtCurPass.IsEnabled || BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(this.txtCurPass.Password.ToLower()))) == StartupBase.SasObj.UserInfo.Rows[0]["password"].ToString())
                return true;
            int num = (int)ExMessageBox.Show(925, StartupBase.SasObj, "Mật khẩu hiện tại không chính xác, vui lòng nhập lại mật khẩu hiện tại!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtCurPass.SelectAll();
            this.txtCurPass.Focus();
            return false;
        }

        public void IsEnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.checkValid())
                return;
            this.saveUser();
            StartUp.currActionTask = ActionTask.None;
            this.DialogResult = new bool?(true);
        }

        private void saveUser()
        {
            string str = this.Encrypt(this.txtpassword.Password);
            this.newDataTable.Rows[0]["password"] = (object)str;
            StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(string.Format("UPDATE UserInfo SET [password] = '{0}' WHERE [user_id] = {1}", (object)str, this.newDataTable.Rows[0]["user_id"])));
            GhiLog();
        }
        public void GhiLog()
        {
            string strsql = "INSERT INTO [dbo].[dmpaslog]([log_id],[machineid],[oldpass],[newpass],[oldname],[user_id],[user_id2],[date_log],[time_log],[status]) " +
                "VALUES(newid(),@machineid,@oldpass,@newpass,@oldname,@user_id,@user_id2,getdate(),convert(varchar(10), GETDATE(), 108),'1')";
            SqlCommand cmd = new SqlCommand(strsql);
            cmd.Parameters.Add("@machineid", SqlDbType.Char).Value = (object)System.Environment.MachineName.Trim();
            cmd.Parameters.Add("@oldpass", SqlDbType.NChar).Value = (object)this.txtCurPass.Password.Trim();
            cmd.Parameters.Add("@newpass", SqlDbType.NChar).Value = (object)this.txtpassword.Password.Trim();
            cmd.Parameters.Add("@oldname", SqlDbType.NChar).Value = (object)this.newDataTable.Rows[0]["user_name"].ToString().Trim();
            cmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)this.newDataTable.Rows[0]["user_id"];
            cmd.Parameters.Add("@user_id2", SqlDbType.Int).Value = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
            StartupBase.SasObj.ExcuteNonQuery(cmd);
        }
        private string Encrypt(string password)
        {
            password = password.ToLower();
            return BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(password)));
        }

        private bool checkValid()
        {
            if (!this.CheckCurPass())
                return false;
            if (this.txtpassword.Password == string.Empty)
            {
                int num = (int)ExMessageBox.Show(950, StartupBase.SasObj, "Chưa vào mật khẩu mới!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtpassword.Focus();
                return false;
            }

            // Kiểm tra pass
            if (M_LEN_PASS > 0)
            {
                if (this.txtpassword.Password.Trim().Length < M_LEN_PASS)
                {
                    int num = (int)ExMessageBox.Show(951, StartupBase.SasObj, "Mật khẩu chưa đủ độ dài!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtpassword.Focus();
                    return false;
                }

                if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsDigit)))
                {
                    int num = (int)ExMessageBox.Show(952, StartupBase.SasObj, "Mật khẩu phải chứa số!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtpassword.Focus();
                    return false;
                }
                if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsUpper)))
                {
                    int num = (int)ExMessageBox.Show(953, StartupBase.SasObj, "Mật khẩu phải chữ hoa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtpassword.Focus();
                    return false;
                }

                if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsLower)))
                {
                    int num = (int)ExMessageBox.Show(954, StartupBase.SasObj, "Mật khẩu phải chữ thường!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtpassword.Focus();
                    return false;
                }
            }
            if (!(this.txtpassword.Password != this.txtgokt.Password))
                return true;
            int num1 = (int)ExMessageBox.Show(955, StartupBase.SasObj, "Mật khẩu xác nhận không chính xác, vui lòng nhập lại!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtgokt.SelectAll();
            this.txtgokt.Focus();
            return false;
        }

        private void ShowCurrentPassword(bool isShow)
        {
            if (isShow)
                return;
            this.grdCurPass.Visibility = Visibility.Collapsed;
            this.Height -= 25.0;
        }
    }
}
