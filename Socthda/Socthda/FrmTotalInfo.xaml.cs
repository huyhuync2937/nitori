using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmTotalInfo : Form
  {
    private CodeValueBindingObject Voucher_Ma_nt0;
    public FrmTotalInfo()
    {
      this.InitializeComponent();
      this.DisplayLanguage = StartUpTrans.M_LAN;
      this.BindingSasObj = StartupBase.SasObj;
      SysFunc.LoadIcon((Window) this);
      this.ConfirmGV.ButtonType = 1;
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      this.Voucher_Ma_nt0 = (CodeValueBindingObject) this.SOCTHDATotalInfo.FindResource((object) "Voucher_Ma_nt0");
      this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
      this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
      this.ConfirmGV.DataContext = (object) StartUpTrans.DsTrans.Tables[0].DefaultView;
    }

    private void ConfirmGV_OnOk(object sender, RoutedEventArgs e)
    {
      this.Close();
    }
  }
}
