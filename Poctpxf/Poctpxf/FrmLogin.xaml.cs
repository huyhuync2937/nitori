using SasControls;
using SasFormBrowes;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Windows;

namespace Poctpxf
{
    public partial class FrmLogin : Form
    {

        public bool IsLogined { get; set; }

        public FrmLogin()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.IsLogined = false;
            this.Title = SysFunc.Cat_Dau("Đăng nhập");
            this.txtPassword.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(this.txtPassword.Password.ToLower()))) == StartupBase.SasObj.UserInfo.Rows[0]["password"].ToString())
            {
                this.IsLogined = true;
                this.Close();
            }
            else
            {
                int num = (int)ExMessageBox.Show(945, StartupBase.SasObj, "Mật khẩu không chính xác, vui lòng nhập lại mật khẩu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtPassword.SelectAll();
                this.txtPassword.Focus();
            }
        }
    }
}
