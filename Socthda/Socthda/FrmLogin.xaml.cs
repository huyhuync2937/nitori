using SasControls;
using SasFormBrowes;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmLogin : Form
  {

    public bool IsLogined { get; set; }

    public FrmLogin()
    {
      this.InitializeComponent();
      SysFunc.LoadIcon((Window) this);
      this.IsLogined = false;
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      this.txtpassword.Focus();
    }

    private string Encrypt(string password)
    {
      password = password.ToLower();
      return BitConverter.ToString(((HashAlgorithm) CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(password)));
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (this.Encrypt(this.txtpassword.Password) != StartupBase.SasObj.UserInfo.Rows[0]["password"].ToString())
      {
        int num = (int) ExMessageBox.Show(385, StartupBase.SasObj, "Mật khẩu không chính xác, vui lòng nhập lại mật khẩu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtpassword.SelectAll();
        this.txtpassword.Focus();
      }
      else
      {
        this.IsLogined = true;
        this.Close();
      }
    }
  }
}
