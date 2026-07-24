using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace Socthda
{
  public partial class FrmTaoPT : Form
  {
    public bool isOk = false;
    public int kind = 1;
    public string Ma_nt_ht = "";
    public string so_hd = "";
    public string ngay_hd = "";
    public string filterma_qs = "";
    public DataTable tbInfoPT = (DataTable) null;
    public FrmTaoPT()
    {
      this.InitializeComponent();
      SysFunc.LoadIcon((Window) this);
      this.BindingSasObj = StartupBase.SasObj;
      this.GrdOkCancel.pnlButton.btnCancel.Visibility = Visibility.Collapsed;
      this.txtso_ct_pt.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
      this.txtnguoi_nop.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("ong_ba");
      this.txtlydo_nop.MaxLength = this.BindingSasObj.GetDatabaseFieldLength("dien_giai");
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      if (this.tbInfoPT == null || this.tbInfoPT != null && this.tbInfoPT.Rows.Count == 0)
      {
        this.txtKind.Value = (object) this.kind;
        SqlCommand sqlcmd = new SqlCommand();
        sqlcmd.CommandText = "if(select count(1) from dmqs where ma_cts LIKE '%{0}%') = 1";
        sqlcmd.CommandText += " select ma_qs from dmqs where ma_cts LIKE '%{1}%'";
        if (this.txtKind.Value.ToString().Equals("1"))
        {
          this.txtMa_qs_pt.Filter = this.filterma_qs.Replace("HDA", "PT1");
          this.txtMa_gd.Filter = "ma_ct = 'PT1' and status = 1 and ma_gd IN ('2','9')";
          sqlcmd.CommandText = string.Format(sqlcmd.CommandText, (object) "PT1", (object) "PT1");
        }
        else
        {
          this.txtMa_qs_pt.Filter = this.filterma_qs.Replace("HDA", "BC1");
          this.txtMa_gd.Filter = "ma_ct = 'BC1' and status = 1 and ma_gd IN ('2','9')";
          sqlcmd.CommandText = string.Format(sqlcmd.CommandText, (object) "BC1", (object) "BC1");
        }
        this.txtMa_gd.Text = StartUpTrans.CommandInfo["parameter"].ToString().Split(';')[0];
        this.txtMa_gd.SearchInit();
        this.txtMa_gd_PreviewLostFocus((object) this.txtMa_gd, (KeyboardFocusChangedEventArgs) null);
        DataSet dataSet = this.BindingSasObj.ExcuteReader(sqlcmd);
        if (dataSet.Tables.Count == 1)
          this.txtMa_qs_pt.Text = dataSet.Tables[0].Rows[0]["ma_qs"].ToString();
        this.txtlydo_nop.Text = !StartUpTrans.M_LAN.Equals("V") ? string.Format("Invoice no. {0}, invoice date {1}", (object) this.so_hd, (object) this.ngay_hd) : string.Format("Thu tiền hóa đơn số {0}, ngày {1}", (object) this.so_hd, (object) this.ngay_hd);
      }
      else
      {
        this.txtKind.Value = (object) (this.tbInfoPT.Rows[0]["ma_ct"].ToString().Equals("PT1") ? 1 : 2);
        if (this.txtKind.Value.ToString().Equals("1"))
        {
          this.txtMa_qs_pt.Filter = this.filterma_qs.Replace("HDA", "PT1");
          this.txtMa_gd.Filter = "ma_ct = 'PT1' and status = 1 and ma_gd IN ('2','9')";
        }
        else
        {
          this.txtMa_qs_pt.Filter = this.filterma_qs.Replace("HDA", "BC1");
          this.txtMa_gd.Filter = "ma_ct = 'BC1' and status = 1 and ma_gd IN ('2','9')";
        }
        this.txtMa_gd.Text = this.tbInfoPT.Rows[0]["ma_gd"].ToString();
        this.txtMa_gd.SearchInit();
        this.txtMa_gd_PreviewLostFocus((object) this.txtMa_gd, (KeyboardFocusChangedEventArgs) null);
        this.txtMa_qs_pt.Text = this.tbInfoPT.Rows[0]["ma_qs"].ToString();
        this.txtso_ct_pt.Text = this.tbInfoPT.Rows[0]["so_ct"].ToString().Trim();
        this.txtnguoi_nop.Text = this.tbInfoPT.Rows[0]["ong_ba"].ToString();
        this.txtlydo_nop.Text = this.tbInfoPT.Rows[0]["dien_giai"].ToString();
      }
      this.txtMa_nt.Text = this.Ma_nt_ht;
      if (this.Ma_nt_ht != StartupBase.M_MA_NT0)
      {
        this.txtMa_nt.Filter = "ma_nt IN ('" + StartupBase.M_MA_NT0 + "','" + this.Ma_nt_ht + "')";
      }
      else
      {
        this.txtMa_nt.IsReadOnly = true;
        this.txtMa_nt.IsTabStop = false;
      }
      this.txtKind.Focus();
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      FormTrans owner = this.Owner as FormTrans;
      if (string.IsNullOrEmpty(this.txtMa_qs_pt.Text.Trim()) || !this.txtMa_qs_pt.CheckLostFocus())
      {
        int num = (int) ExMessageBox.Show(697, StartupBase.SasObj, "Chưa vào mã quyển sổ phiếu thu", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtMa_qs_pt.IsFocus = true;
      }
      else if (string.IsNullOrEmpty(this.txtso_ct_pt.Text.Trim()))
      {
        int num = (int) ExMessageBox.Show(698, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtso_ct_pt.Focus();
      }
      else
      {
        string stt_rec = "";
        if (this.tbInfoPT != null && this.tbInfoPT.Rows.Count == 1)
          stt_rec = this.tbInfoPT.Rows[0]["stt_rec"].ToString();
        if (owner.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text, this.txtso_ct_pt.Text.PadLeft(this.txtso_ct_pt.MaxLength, ' '), stt_rec))
        {
          if (this.txtMa_qs_pt.RowResult["chkso_ct"].ToString().Equals("1"))
          {
            if (ExMessageBox.Show(699, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
            {
              this.txtso_ct_pt.SelectAll();
              this.txtso_ct_pt.Focus();
              return;
            }
          }
          else if (this.txtMa_qs_pt.RowResult["chkso_ct"].ToString().Equals("2"))
          {
            int num = (int) ExMessageBox.Show(694, StartupBase.SasObj, "Số chứng từ đã tồn tại. Số cuối cùng là: [" + owner.GetLastSoct(StartupBase.SasObj, this.txtMa_qs_pt.Text).Trim() + "]", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtso_ct_pt.SelectAll();
            this.txtso_ct_pt.Focus();
            return;
          }
        }
        this.isOk = true;
        this.Close();
      }
    }

    private void txtMa_qs_pt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      if (e.NewFocus == this.GrdOkCancel.pnlButton.btnOk)
        return;
      this.txtso_ct_pt.Text = (this.Owner as FormTrans).GetNewSoct(this.BindingSasObj, this.txtMa_qs_pt.Text);
    }

    public string GetNewSoct(string ma_qs)
    {
      string str1;
      DataTable table = this.BindingSasObj.ExcuteReader(new SqlCommand(Convert.ToInt16(this.BindingSasObj.GetOption("M_AUTO_SOCT").ToString()) == (short) 1 ? "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'" : (str1 = "EXEC  [GetNewSoct] '" + ma_qs.Trim() + "'"))).Tables[0];
      if (table.Rows.Count > 0)
      {
        DataRow row = table.Rows[0];
        if (row[1] != null && row[1] != DBNull.Value)
        {
          string str2 = row[1].ToString();
          return string.Format(row[0].ToString(), (object) Convert.ToDouble(str2));
        }
      }
      return "";
    }

    private void txtKind_LostFocus(object sender, RoutedEventArgs e)
    {
      if (string.IsNullOrEmpty(this.txtKind.Text))
        this.txtKind.Value = (object) this.kind;
      if (this.txtKind.Value.ToString().Equals("1"))
        this.txtMa_qs_pt.Filter = "ma_cts like '%PT1%' and status=1 \r\n                                        AND ((NOT EXISTS(SELECT 1 FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs) \r\n                                        OR EXISTS (SELECT ma_qs FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs \r\n                                        AND [user_id] = " + (object) StartUpTrans.M_User_Id + ")))";
      else
        this.txtMa_qs_pt.Filter = "ma_cts like '%BC1%' and status=1 \r\n                                        AND ((NOT EXISTS(SELECT 1 FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs) \r\n                                        OR EXISTS (SELECT ma_qs FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs \r\n                                        AND [user_id] = " + (object) StartUpTrans.M_User_Id + ")))";
    }

    private void txtMa_gd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      if (this.txtMa_gd.RowResult == null)
        return;
      this.txtTen_gd.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_gd.RowResult["ten_gd"].ToString() : this.txtMa_gd.RowResult["ten_gd2"].ToString();
    }
  }
}
