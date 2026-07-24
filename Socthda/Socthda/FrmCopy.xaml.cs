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
  public partial class FrmCopy : Form
  {
    public bool isCopy = false;
    public DateTime ngay_ct;

    public FrmCopy()
    {
      this.InitializeComponent();
    }

    private void FrmCopy_Loaded(object sender, RoutedEventArgs e)
    {
      this.txtNgay_ct_old.Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ngay_ct"];
      this.txtNgay_ct_new.Value = (object) DateTime.Now.Date;
      this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtNgay_ct_new.Focus()));
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (!this.CheckValid())
        return;
      this.ngay_ct = DateTime.Parse(this.txtNgay_ct_new.Value.ToString()).Date;
      this.isCopy = true;
      this.Close();
    }

    private bool CheckValid()
    {
      bool flag = true;
      if (flag && (this.txtNgay_ct_new.Value == null || this.txtNgay_ct_new.Value.ToString() == ""))
      {
        int num = (int) ExMessageBox.Show(350, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        flag = false;
        this.txtNgay_ct_new.Focus();
      }
      if (flag && this.txtNgay_ct_new.Value.ToString() != "")
      {
        if (!this.txtNgay_ct_new.IsValueValid)
        {
          int num = (int) ExMessageBox.Show(355, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          flag = false;
          this.txtNgay_ct_new.Focus();
          this.txtNgay_ct_new.SelectAll();
        }
        if (flag && !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(Convert.ToDateTime(this.txtNgay_ct_new.dValue))))
        {
          int num = (int) ExMessageBox.Show(360, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          flag = false;
          this.txtNgay_ct_new.Focus();
          this.txtNgay_ct_new.SelectAll();
        }
        if (flag && Convert.ToDateTime(this.txtNgay_ct_new.dValue) < NgayTC.GetStartDate(StartUp.M_ngay_ct0))
        {
          int num = (int) ExMessageBox.Show(365, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          flag = false;
          this.txtNgay_ct_new.Focus();
          this.txtNgay_ct_new.SelectAll();
        }
      }
      return flag;
    }

    private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
    {
      this.isCopy = false;
      this.Close();
    }
  }
}
