using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmKHInfo : Form
  {
    public bool isError = true;
   
    public FrmKHInfo()
    {
      this.InitializeComponent();
      this.DisplayLanguage = StartUpTrans.M_LAN;
      this.BindingSasObj = StartupBase.SasObj;
      SysFunc.LoadIcon((Window) this);
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      DataRow row = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow];
      if (row["ten_kh_thue"].ToString().Trim() == "")
        row["ten_kh_thue"] = (object) row["ten_kh"].ToString();
      this.txtTen_kh.Text = row["ten_kh_thue"].ToString();
      this.txtDia_chi.Text = row["dia_chi"].ToString();
      this.txtMa_so_thue.Text = row["ma_so_thue"].ToString();
      this.txtTen_kh.Focus();
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (!this.CheckValid())
        return;
      DataRow row = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow];
      row["ten_kh_thue"] = (object) this.txtTen_kh.Text;
      row["dia_chi"] = (object) this.txtDia_chi.Text;
      row["ma_so_thue"] = (object) this.txtMa_so_thue.Text;
      this.isError = false;
      this.Close();
    }

    private bool CheckValid()
    {
      bool flag = true;
      if (!StartUpTrans.M_MST_CHECK.Equals("0") && this.txtMa_so_thue.Text.Trim() != "" && !SysFunc.CheckSumMaSoThue(this.txtMa_so_thue.Text.Trim()))
      {
        if (StartUpTrans.M_MST_CHECK.Equals("1"))
        {
          int num1 = (int) ExMessageBox.Show(375, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }
        else
        {
          int num2 = (int) ExMessageBox.Show(380, StartupBase.SasObj, "Mã số thuế không hợp lệ, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          flag = false;
          this.txtMa_so_thue.Focus();
        }
      }
      return flag;
    }

    private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
    {
      this.isError = true;
      this.Close();
    }
  }
}
