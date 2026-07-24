using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmSocthda_Pn : Form
  {
    private CodeValueBindingObject Voucher_Ma_nt0;
    public DataRowView drvFrmSOCTPNF_PN;
  
    public FrmSocthda_Pn(DataTable tbSource, string ten_vt)
    {
      this.InitializeComponent();
      SysFunc.LoadIcon((Window) this);
      this.EscToClose = true;
      this.Title = SysFunc.Cat_Dau(ten_vt.ToString());
      this.GrdSOCTHDA_PN.DataSource = (IEnumerable) tbSource.DefaultView;
    }

    private void FrmPoctpxf_PN_Loaded(object sender, RoutedEventArgs e)
    {
      this.Voucher_Ma_nt0 = (CodeValueBindingObject) this.FindResource((object) "Voucher_Ma_nt0");
      this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
      this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
      if (this.GrdSOCTHDA_PN.Records.Count > 0)
      {
        this.GrdSOCTHDA_PN.Focus();
        this.GrdSOCTHDA_PN.ActiveRecord = this.GrdSOCTHDA_PN.Records[0];
      }
      this.isVisibleField();
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (this.GrdSOCTHDA_PN.ActiveRecord == null)
        return;
      this.drvFrmSOCTPNF_PN = (this.GrdSOCTHDA_PN.ActiveRecord as DataRecord).DataItem as DataRowView;
      this.Close();
    }

    private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
    {
      this.Close();
    }

    public void isVisibleField()
    {
      if (this.Voucher_Ma_nt0.Text.Trim().Equals(StartUpTrans.M_ma_nt0))
      {
        this.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
        this.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
      }
      else
      {
        this.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Visible;
        this.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = this.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Width.Value.Value;
      }
    }
  }
}
