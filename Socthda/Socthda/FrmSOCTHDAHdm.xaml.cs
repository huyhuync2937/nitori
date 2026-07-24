using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmSOCTHDAHdm : Form
  {
    public bool isOk = false;
    public FrmView frm;
    public DataSet dsHdm;
   
    public FrmSOCTHDAHdm()
    {
      this.InitializeComponent();
      this.BindingSasObj = StartupBase.SasObj;
      SysFunc.LoadIcon((Window) this);
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      this.txtNgay_ct_old.Focus();
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (!string.IsNullOrEmpty(this.txtNgay_ct_old.Text.Trim()))
      {
        try
        {
          Convert.ToDateTime(this.txtNgay_ct_old.Text);
        }
        catch
        {
          int num = (int) ExMessageBox.Show(825, StartupBase.SasObj, "Ngày bắt đầu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          this.txtNgay_ct_old.Focus();
          return;
        }
      }
      if (!string.IsNullOrEmpty(this.txtNgay_ct_new.Text.Trim()))
      {
        try
        {
          Convert.ToDateTime(this.txtNgay_ct_new.Text);
        }
        catch
        {
          int num = (int) ExMessageBox.Show(830, StartupBase.SasObj, "Ngày kết thúc không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          this.txtNgay_ct_new.Focus();
          return;
        }
      }
      this.frm = new FrmView(this.GetFilter());
      this.frm.ShowDialog();
      this.isOk = this.frm.isOk;
      this.dsHdm = this.frm.dsHdm;
      this.Close();
    }

    private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
    {
      this.isOk = false;
      this.Close();
    }

    private void Form_KeyUp(object sender, KeyEventArgs e)
    {
      if (e.Key != Key.Escape)
        return;
      this.isOk = false;
      this.Close();
    }

    private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      this.tblTen_kh.Text = "";
      if (this.txtMa_kh.RowResult == null)
        return;
      this.tblTen_kh.Text = !StartUpTrans.M_LAN.ToUpper().Equals("V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
    }

    private string GetFilter()
    {
      string str = "1=1";
      if (this.txtNgay_ct_old.dValue != new DateTime())
        str = str + " AND ngay_ct >= '" + string.Format("{0:yyyyMMdd}", (object) this.txtNgay_ct_old.dValue) + "'";
      if (this.txtNgay_ct_new.dValue != new DateTime())
        str = str + " AND ngay_ct <= '" + string.Format("{0:yyyyMMdd}", (object) this.txtNgay_ct_new.dValue) + "'";
      if (!string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()))
        str = str + " AND ma_kh LIKE '%" + this.txtMa_kh.Text + "%'";
      if (!string.IsNullOrEmpty(this.txtma_hdm.Text.Trim()))
        str = str + " AND ma_hd LIKE '%" + this.txtma_hdm.Text + "%'";
      return str + " AND status = 2 and  isnull(status2,'') <> '2'";
    }
  }
}
