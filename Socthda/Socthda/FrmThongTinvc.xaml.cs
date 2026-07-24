using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Threading;

namespace Socthda
{
  public partial class FrmThongTinvc : Form
  {
    private CodeValueBindingObject Voucher_Ma_nt0;
    public CodeValueBindingObject IsInEditMode;

    public FrmThongTinvc()
    {
      this.InitializeComponent();
      this.DisplayLanguage = StartUpTrans.M_LAN;
      this.BindingSasObj = StartupBase.SasObj;
      SysFunc.LoadIcon((Window) this);
      this.ConfirmGV.ButtonType = 1;
      this.IsInEditMode = (CodeValueBindingObject) this.FindResource((object) nameof (IsInEditMode));
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      this.Voucher_Ma_nt0 = (CodeValueBindingObject) this.SOCTHDAThongTinvc.FindResource((object) "Voucher_Ma_nt0");
      this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
      this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
      this.ConfirmGV.DataContext = (object) StartUpTrans.DsTrans.Tables[0].DefaultView;
      if (!this.IsInEditMode.Value)
        return;
      this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.txtso_hd.Focus()), DispatcherPriority.Background);
    }

    private void ConfirmGV_OnOk(object sender, RoutedEventArgs e)
    {
      this.Close();
    }
  }
}
