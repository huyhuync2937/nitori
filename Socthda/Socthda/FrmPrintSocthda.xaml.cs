using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Threading;

namespace Socthda
{
  public partial class FrmPrintSocthda : Form
  {
    public DataSet DsPrint = new DataSet();
    private DataSet DsTmpPrint = (DataSet) null;
    private int so_dong_in = 1;
    private bool isFirstLoad = true;
    private bool viewing = false;
    private bool IsND51;

    public FrmPrintSocthda(bool isND51)
    {
      this.InitializeComponent();
      SysFunc.LoadIcon((Window) this);
      this.GridSearch.LocalSasObj = StartupBase.SasObj;
      this.IsND51 = isND51;
      this.GridSearch.ReportGroupName = !isND51 ? StartUpTrans.CommandInfo["rep_file"].ToString() : StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_file"].ToString();
      if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
        this.btnExport.Content = (object) this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
      else
        this.btnExport.Content = (object) this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
    }

    private void Form_Loaded(object sender, RoutedEventArgs e)
    {
      this.tblHt_tt.Visibility = Visibility.Collapsed;
      this.txtHt_tt.Visibility = Visibility.Collapsed;
      this.GridSearch.DSource = this.DsPrint;
      if (this.isFirstLoad)
      {
        this.isFirstLoad = false;
        this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
        this.Dispatcher.BeginInvoke((Delegate) new Action(() =>
        {
          if (this.GridSearch.XGReport == null)
            return;
          this.GridSearch.XGReport.RecordActivated += new EventHandler<RecordActivatedEventArgs>(this.XGReport_RecordActivated);
        }), DispatcherPriority.Background);
      }
      DataTable phIn = StartUpTrans.GetPhIn();
      if (phIn.Rows.Count == 0)
      {
        DataRow row = phIn.NewRow();
        row["ma_ct"] = (object) StartUpTrans.Ma_ct;
        row["stt_rec"] = (object) StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
        row["so01"] = (object) 0;
        row["so02"] = StartUpTrans.DmctInfo["so_lien"];
        phIn.Rows.Add(row);
      }
      if (phIn.Rows.Count == 1)
      {
        DataRow row = phIn.Rows[0];
        if (this.IsND51)
          row["so02"] = (object) StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
      }
      this.DsTmpPrint = this.DsPrint.Copy();
      this.so_dong_in = (int) Convert.ToInt16(StartUpTrans.DmctInfo["so_dong_in"]);
      this.DataContext = (object) phIn;
    }

    private bool IsPhieuht()
    {
      return (((DataRecord) this.GridSearch.XGReport.ActiveRecord).DataItem as DataRowView)["nhom_bc"].ToString().Trim() == "VcIn";
    }

    private void XGReport_RecordActivated(object sender, RecordActivatedEventArgs e)
    {
      if (this.GridSearch.XGReport.ActiveRecord == null || this.GridSearch.XGReport.ActiveRecord.RecordType != RecordType.DataRecord)
        return;
      if (this.IsPhieuht())
      {
        this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
      }
      else
      {
        this.Form_Loaded((object) null, (RoutedEventArgs) null);
        if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1")
        {
          this.txtlien.IsReadOnly = true;
          this.txtlien.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
        }
        else
        {
          this.txtlien.IsReadOnly = false;
          this.txtlien.Text = StartUpTrans.DmctInfo["so_lien"].ToString().Trim();
        }
      }
    }

    private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
      this.View();
    }

    private void InsertRowCT(string nd51)
    {
      if (((IEnumerable<DataRow>) this.DsPrint.Tables["TableCT"].DefaultView.ToTable().Select("tag = 1")).Count<DataRow>() > 0)
        return;
      string str = this.DsPrint.Tables["TablePH"].DefaultView[0]["stt_rec"].ToString();
      int num1 = 1;
      foreach (DataRowView dataRowView in this.DsPrint.Tables["TableCT"].DefaultView)
      {
        dataRowView["stt"] = dataRowView["stt_rec0"] != (object) "999" ? (object) num1 : (object) DBNull.Value;
        ++num1;
      }
      Decimal num2 = Convert.ToDecimal(this.DsPrint.Tables["TablePH"].DefaultView[0]["t_ck"]);
      Decimal num3 = Convert.ToDecimal(this.DsPrint.Tables["TablePH"].DefaultView[0]["t_ck_nt"]);
      if (num2 != new Decimal(0) || num3 != new Decimal(0))
      {
        DataRow row = this.DsPrint.Tables["TableCT"].NewRow();
        row["stt_rec"] = (object) str;
        row["stt_rec0"] = (object) -1;
        row["ten_vt"] = (object) "Chiết khấu";
        row["ten_vt2"] = (object) "Discount";
        row["tien2"] = (object) num2;
        row["tien_nt2"] = (object) num3;
        row["tag"] = (object) 1;
        this.DsPrint.Tables["TableCT"].Rows.Add(row);
      }
      int count = this.DsPrint.Tables["TableCT"].DefaultView.Count;
      this.GridSearch.InsertSubRow("HDA", "TableCT");
      this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
      this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
      this.GridSearch.DSource = this.DsPrint;
    }

    private void ResetTableCt()
    {
      string str = this.DsPrint.Tables["TablePH"].DefaultView[0]["stt_rec"].ToString();
      if (this.DsTmpPrint == null)
        return;
      this.DsPrint = this.DsTmpPrint.Copy();
      this.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
      this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
      this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
      this.GridSearch.DSource = this.DsPrint;
    }

    private void btnin_Click(object sender, RoutedEventArgs e)
    {
      if (this.GridSearch.XGReport.ActiveRecord == null)
        return;
      if (this.IsPhieuht())
      {
        this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
        this.GridSearch.V_In((short) 1);
      }
      else
      {
        this.hddt();
        if (this.txtlien.Value != null)
        {
          int int16_1 = (int) Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]);
          DataRowView dataItem = (this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView;
          string nd51 = dataItem["nd51"].ToString();
          if (nd51 == "1" && int16_1 > 0)
          {
            if (ExMessageBox.Show(390, StartupBase.SasObj, "Hóa đơn đã được in, có muốn in lại hay không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
              return;
            FrmLogin frmLogin = new FrmLogin();
            frmLogin.ShowDialog();
            if (!frmLogin.IsLogined)
              return;
          }
          int num = 1;
          int result1 = 0;
          int result2 = 0;
          int int32 = Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord) this.GridSearch.XGReport.ActiveRecord, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()));
          int.TryParse(this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien_hd"].ToString(), out result1);
          int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result2);
          if (int32 > result1)
            this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object) "BẢN SAO";
          if (int16_1 >= 1)
            this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object) "BẢN SAO";
          for (int int16_2 = (int) Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value))); num <= int16_2; ++num)
          {
            this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien"] = (object) num;
            if (int32 <= result1)
              this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = num > int32 && num <= result1 && int16_1 < 1 ? (object) "" : (object) "BẢN SAO";
            this.InsertRowCT(nd51);
            this.GridSearch.V_In((short) 1, result2 >= num && string.IsNullOrEmpty(this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"].ToString()));
          }
          if (nd51 == "1" && this.GridSearch.PrintSuccess)
          {
            string stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartUpTrans.UpdateSl_in(stt_rec, dataItem["id"].ToString(), this.txtlien.Text);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = StartUpTrans.GetSl_in(stt_rec);
          }
          this.ResetTableCt();
          this.DsPrint.Tables["TablePH"].DefaultView[0]["ht_tt"] = (object) this.txtHt_tt.Text;
          DataTable dataContext = this.DataContext as DataTable;
          StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }
        this.Close();
      }
    }

    private void btnin_lt_Click(object sender, RoutedEventArgs e)
    {
      if (this.GridSearch.XGReport.ActiveRecord == null)
        return;
      DataRowView dataItem = (this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView;
      string nd51 = dataItem["mau_tu_in"].ToString();
      bool flag1 = false;
      for (int index = 1; index < this.DsPrint.Tables[0].Rows.Count; ++index)
      {
        if (!this.DsPrint.Tables[0].Rows[index]["tinh_trang_hddt"].ToString().Equals("0"))
        {
          flag1 = true;
          break;
        }
      }
      if (flag1)
      {
        int num1 = (int) ExMessageBox.Show(396, StartupBase.SasObj, "Có chứng từ thuộc hóa đơn điện tử, không in liên tục được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
      }
      else if (StartUp.IsQLHD && nd51 == "1")
      {
        int num2 = (int) ExMessageBox.Show(395, StartupBase.SasObj, "Có chứng từ thuộc mẫu hóa đơn tự in, không in liên tục được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
      }
      else
      {
        if (!(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString() == "1") || ExMessageBox.Show(400, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
          return;
        if (this.IsPhieuht())
        {
          for (int index = 1; index < this.DsPrint.Tables[0].Rows.Count; ++index)
          {
            this.GridSearch.DSource = StartUp.GetPhieuht(this.DsPrint.Tables[0].Rows[index]["stt_rec"].ToString());
            this.GridSearch.V_In((short) 1);
          }
        }
        else
        {
          List<int> intList = new List<int>();
          if (this.txtlien.Value != null)
          {
            int iRow = FrmSocthda.iRow;
            bool flag2 = false;
            int result1 = 0;
            int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result1);
            int int16_1 = (int) Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
            for (int index1 = 1; index1 < this.DsPrint.Tables[0].Rows.Count; ++index1)
            {
              for (int index2 = 1; index2 <= int16_1; ++index2)
              {
                string stt_rec = this.DsPrint.Tables[0].Rows[index1]["stt_rec"].ToString();
                this.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
                this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
                this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
                this.DsPrint.Tables["TableMST_NM"].Rows.Clear();
                this.DsPrint.Tables["TableMST_NM"].Rows.Add(FrmSocthda.CreateTableMST(this.DsPrint.Tables["TablePH"].DefaultView[0]["ma_so_thue"].ToString().TrimEnd(), "TableMST_NM").Rows[0].ItemArray);
                if (index2 == 1)
                  intList.Add(Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord) this.GridSearch.XGReport.ActiveRecord, stt_rec)));
                if (this.DsPrint.Tables[0].Rows[index1]["status"].ToString() != "3")
                {
                  int num3 = intList[index1 - 1];
                  int result2 = 0;
                  int.TryParse(this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien_hd"].ToString(), out result2);
                  int int16_2 = (int) Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].Rows[index1]["sl_in"]);
                  if (int16_2 >= 1)
                    this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object) "BẢN SAO";
                  this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = num3 <= result2 ? (index2 > num3 && index2 <= result2 && int16_2 < 1 ? (object) "" : (object) "BẢN SAO") : (object) "BẢN SAO";
                  this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien"] = (object) index2;
                  if (nd51 == "1" && int16_2 > 0 && !flag2)
                  {
                    if (ExMessageBox.Show(405, StartupBase.SasObj, "Hóa đơn đã được in, có muốn in lại hay không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                      return;
                    FrmLogin frmLogin = new FrmLogin();
                    frmLogin.ShowDialog();
                    if (!frmLogin.IsLogined)
                      return;
                    flag2 = true;
                  }
                  this.InsertRowCT(nd51);
                  this.GridSearch.V_In((short) 1, result1 >= index2 && string.IsNullOrEmpty(this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"].ToString()));
                  if (nd51 == "1" && index2 == 1 && this.GridSearch.PrintSuccess)
                  {
                    StartUpTrans.UpdateSl_in(stt_rec, dataItem["id"].ToString(), this.txtlien.Text);
                    StartUpTrans.DsTrans.Tables[0].Rows[index1]["sl_in"] = StartUpTrans.GetSl_in(stt_rec);
                  }
                }
              }
            }
            this.ResetTableCt();
            this.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables["TablePH"].Rows[iRow]["stt_rec"].ToString() + "'";
            this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables["TablePH"].Rows[iRow]["stt_rec"].ToString() + "'";
            this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
            this.DsPrint.Tables["TableMST_NM"].Rows.Clear();
            this.DsPrint.Tables["TableMST_NM"].Rows.Add(FrmSocthda.CreateTableMST(this.DsPrint.Tables["TablePH"].DefaultView[0]["ma_so_thue"].ToString().TrimEnd(), "TableMST_NM").Rows[0].ItemArray);
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
          }
          this.Close();
        }
      }
    }

    private void btnxem_Click(object sender, RoutedEventArgs e)
    {
      this.View();
      this.ht_tt();
    }

    private void View()
    {
      if (this.viewing)
        return;
      this.viewing = true;
      this.txtHt_tt.Text = FrmSocthda.hinhthuc_tt;
      if (this.GridSearch.XGReport.ActiveRecord == null)
        return;
      if (this.IsPhieuht())
      {
        this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
        this.GridSearch.V_Xem();
      }
      else
      {
        this.hddt();
        string nd51 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString();
        int int32 = Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord) this.GridSearch.XGReport.ActiveRecord, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()));
        if (Convert.ToInt16(this.DsPrint.Tables["TablePH"].DefaultView[0]["sl_in"]) >= (short) 1)
          this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object) "BẢN SAO";
        if (int32 > 0)
          this.DsPrint.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object) "BẢN SAO";
        this.DsPrint.Tables[0].DefaultView[0]["so_ct_goc"] = (object) this.txtctu0.Text;
        this.DsPrint.Tables[0].DefaultView[0]["ht_tt"] = (object) this.txtHt_tt.Text;
        this.InsertRowCT(nd51);
        StartUpTrans.SetPhIn(this.DataContext as DataTable);
        this.GridSearch.V_Xem();
        this.ResetTableCt();
        new Thread((ParameterizedThreadStart) (x =>
        {
          Thread.Sleep(2000);
          this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.viewing = false), DispatcherPriority.ApplicationIdle);
        })).Start();
      }
    }

    private void btnthoat_Click(object sender, RoutedEventArgs e)
    {
      this.Close();
    }

    private void txtlien_LostFocus(object sender, RoutedEventArgs e)
    {
      if (this.txtlien.IsFocusWithin || !(this.txtlien.Value.ToString() == ""))
        return;
      this.txtlien.Value = (object) 0;
    }

    private void txtctu0_LostFocus(object sender, RoutedEventArgs e)
    {
      if (this.txtctu0.IsFocusWithin)
        return;
      if (this.txtctu0.Value.ToString() == "")
        this.txtctu0.Value = (object) 0;
      this.DsPrint.Tables["TablePH"].DefaultView[0]["so_ct_goc"] = this.txtctu0.Value;
    }

    private void btnExport_Click(object sender, RoutedEventArgs e)
    {
      if (this.DsPrint.Tables[0].DefaultView.Count != 1)
        return;
      this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.DsPrint.Tables[0].DefaultView[0]), new WindowInteropHelper((Window) this).Handle);
    }

    private void txtHt_tt_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      this.txtHt_tt.SelectionStart = 0;
    }

    private void txtHt_tt_LostFocus(object sender, RoutedEventArgs e)
    {
      this.DsPrint.Tables[0].DefaultView[0]["ht_tt"] = (object) this.txtHt_tt.Text;
    }

    private void hddt()
    {
      if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["is_hddt"].ToString().Equals("1") || !StartUp.M_SD_HDDT.Equals("1") || this.DsPrint.Tables["TablePH"].DefaultView[0]["sd_hddt_yn"].ToString().Trim().Equals("0"))
        return;
      if (this.DsPrint.Tables["TablePH"].DefaultView[0]["mau_hddt"].ToString().Trim() != string.Empty)
      {
        this.DsPrint.Tables["TablePH"].DefaultView[0]["kh_mau_hd"] = (object) this.DsPrint.Tables["TablePH"].DefaultView[0]["mau_hddt"].ToString();
        this.DsPrint.Tables["TablePH"].DefaultView[0]["mau_hd"] = (object) this.DsPrint.Tables["TablePH"].DefaultView[0]["mau_hddt"].ToString();
      }
      if (this.DsPrint.Tables["TablePH"].DefaultView[0]["so_seri_hddt"].ToString().Trim() != string.Empty)
        this.DsPrint.Tables["TablePH"].DefaultView[0]["so_seri"] = (object) this.DsPrint.Tables["TablePH"].DefaultView[0]["so_seri_hddt"].ToString();
      if (this.DsPrint.Tables["TablePH"].DefaultView[0]["so_ct_hddt"].ToString().Trim() != string.Empty)
        this.DsPrint.Tables["TablePH"].DefaultView[0]["so_ct"] = (object) this.DsPrint.Tables["TablePH"].DefaultView[0]["so_ct_hddt"].ToString();
    }

    private void ht_tt()
    {
      SqlCommand sqlcmd = new SqlCommand("update phin set ht_tt= @ht_tt where stt_rec=@stt_rec");
      sqlcmd.Parameters.Add("@ht_tt", SqlDbType.NVarChar).Value = (object) this.txtHt_tt.Text;
      sqlcmd.Parameters.Add("@stt_rec", SqlDbType.NVarChar).Value = (object) StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
      StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
    }

  }
}
