using SasControls;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using SasLib;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace Socthda
{
  public partial class FrmSearchSocthda : FormFilter
  {
    public new static readonly DependencyProperty SasObjProperty = DependencyProperty.Register(nameof (SasObj), typeof (SasObject), typeof (FrmSearchSocthda), (PropertyMetadata) new UIPropertyMetadata((PropertyChangedCallback) null));
   
    public FrmSearchSocthda(SasObject _SasObj, string _filterID, string _tableList)
    {
      this.InitializeComponent();
      this.SasObj = _SasObj;
      this.BindingSasObj = _SasObj;
      this.GridSearch.filterID = _filterID;
      this.GridSearch.tableList = _tableList;
      this.GridSearch.SasObj = _SasObj;
      this.M_LAN = StartUpTrans.M_LAN;
    }

    public string M_LAN { get; set; }

    public SasObject SasObj
    {
      get
      {
        return (SasObject) this.GetValue(FrmSearchSocthda.SasObjProperty);
      }
      set
      {
        this.SetValue(FrmSearchSocthda.SasObjProperty, (object) value);
      }
    }

    private void FrmSearchSocthda_Loaded(object sender, RoutedEventArgs e)
    {
      this.txtMa_kh.SearchInit();
      this.txtMaDVCS.SearchInit();
      this.txtMa_nx.SearchInit();
      this.txtMa_bp.SearchInit();
      DataView defaultView = this.SasObj.GetPostInfo(StartUpTrans.Ma_ct).DefaultView.ToTable().DefaultView;
      DataRow row = defaultView.Table.NewRow();
      row["ten_post"] = row["ten_act"] = (object) "Tất cả";
      row["ten_post2"] = row["ten_act2"] = (object) "All";
      defaultView.Table.Rows.Add(row);
      this.txtStatus.ItemsSource = (IEnumerable) defaultView.Table.AsEnumerable().OrderBy<DataRow, string>((Func<DataRow, string>) (x => x.Field<string>("ma_post"))).AsDataView<DataRow>();
      this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.txtStatus.SelectedIndex = 0), DispatcherPriority.Background);
      if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
      {
        if (this.txtMa_kh.RowResult != null)
          this.tblTen_kh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
        if (this.txtMaDVCS.RowResult != null)
          this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
        if (this.txtMa_nx.RowResult != null)
          this.tblTen_tk.Text = this.txtMa_nx.RowResult["ten_tk"].ToString();
        if (this.txtMa_bp.RowResult != null)
          this.lbltenbp.Text = this.txtMa_bp.RowResult["ten_bp"].ToString();
      }
      else
      {
        if (this.txtMa_kh.RowResult != null)
          this.tblTen_kh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
        if (this.txtMaDVCS.RowResult != null)
          this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
        if (this.txtMa_nx.RowResult != null)
          this.tblTen_tk.Text = this.txtMa_nx.RowResult["ten_tk2"].ToString();
        if (this.txtMa_bp.RowResult != null)
          this.lbltenbp.Text = this.txtMa_bp.RowResult["ten_bp2"].ToString();
      }
      this.txtNgay_ct1.Value = (object) (DateTime) this.SasObj.GetSysvar("M_ngay_ct1");
      this.txtNgay_ct2.Value = (object) (DateTime) this.SasObj.GetSysvar("M_ngay_ct2");
      this.txtloc_nsd.Value = StartUpTrans.DmctInfo["m_loc_nsd"] == DBNull.Value ? (object) 0 : StartUpTrans.DmctInfo["m_loc_nsd"];
      this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.txtMa_qs.IsFocus = true), DispatcherPriority.Background);
      this.txtMa_qs.IsFocus = true;
      this.txtMa_qs.SelectAllOnFocus = true;
    }

    private void FormSearch_Unloaded(object sender, RoutedEventArgs e)
    {
      this.GridSearch.SasObj = (SasObject) null;
      this.SasObj = (SasObject) null;
      this.BindingSasObj = (SasObject) null;
      foreach (DependencyObject child in this.GrdPhLoc.Children)
        BindingOperations.ClearAllBindings(child);
    }

    private string GetPhFilterExpr()
    {
      int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
      string str = "1=1 ";
      if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
        str = str + " and ngay_ct >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof (DateTime));
      if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
        str = str + " and ngay_ct <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof (DateTime));
      if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
        str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
      if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
        str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
      if (!string.IsNullOrEmpty(this.txtMa_kh.Text))
        str = str + " and ma_kh = " + this.ConvertDataToSql((object) this.txtMa_kh.Text.Trim(), typeof (string));
      if (!string.IsNullOrEmpty(this.txtMa_nx.Text))
        str = str + " and ma_nx like " + this.ConvertDataToSql((object) (this.txtMa_nx.Text.Trim() + "%"), typeof (string));
      if (Convert.ToInt16(this.txtloc_nsd.Value) == (short) 1)
        str = str + " and [user_id] = " + (object) StartUpTrans.M_User_Id;
      if (!string.IsNullOrEmpty(this.txtMa_qs.Text))
        str = str + " and ma_qs LIKE '" + this.txtMa_qs.Text.Trim() + "%'";
      if (!string.IsNullOrEmpty(this.txtMa_bp.Text))
        str = str + " and ma_bp LIKE '" + this.txtMa_bp.Text.Trim() + "%'";
      if (this.txtStatus.Value != null && !string.IsNullOrEmpty(this.txtStatus.Value.ToString().Trim()))
        str = str + " and status = '" + this.txtStatus.Value.ToString().Trim() + "'";
      if (!string.IsNullOrEmpty(this.txtMaDVCS.Text))
        str = str + " and ma_dvcs LIKE '" + this.txtMaDVCS.Text.Trim() + "%'";
      if (!SysFunc.CheckPermission(this.SasObj, ActionTask.View, StartupBase.Menu_Id))
        str = str + " AND user_id0 = " + this.SasObj.UserInfo.Rows[0]["user_id"].ToString();
      if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
        str = str + " and " + this.GridSearch.arrStrFilter[0];
      return str;
    }

    private string GetCtFilterExpr()
    {
      string str = "1=1 ";
      if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[1]))
        str = str + " and " + this.GridSearch.arrStrFilter[1];
      return str;
    }

    public string ConvertDataToSql(object value, Type ValueType)
    {
      string str;
      switch (ValueType.ToString())
      {
        case "System.String":
          str = string.Format("'{0}'", (object) (value as string).Replace("'", "'"));
          break;
        case "System.DateTime":
          str = string.Format("'{0}'", (object) ((DateTime) value).ToString("yyyyMMdd"));
          break;
        default:
          str = string.Format("'{0}'", value);
          break;
      }
      return str;
    }

    private void grdConfirm_OnOk(object sender, RoutedEventArgs e)
    {
      try
      {
        if (Keyboard.FocusedElement.GetType().Equals(typeof (TextBoxAutoComplete)) && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus() || !this.CheckValid())
          return;
        this.SasObj.SetSysvar("M_ngay_ct1", (object) this.txtNgay_ct1.dValue);
        this.SasObj.SetSysvar("M_ngay_ct2", (object) this.txtNgay_ct2.dValue);
        bool flag = false;
        this.GridSearch._GenerateSQLString();
        this.GridSearch.GrdSearch.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
        StartUp.TransFilterCmd.Parameters["@PhFilter"].Value = (object) this.GetPhFilterExpr();
        StartUp.TransFilterCmd.Parameters["@CtFilter"].Value = (object) this.GetCtFilterExpr();
        StartUp.TransFilterCmd.Parameters["@Sl_ct"].Value = (object) -1;
        DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
        string str1 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>) (p => p.Field<Decimal?>("t_tt"))).Sum().Value.ToString(this.SasObj.GetOption("M_IP_TIEN").ToString());
        string str2 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>) (p => p.Field<Decimal?>("t_tt_nt"))).Sum().Value.ToString(this.SasObj.GetOption("M_IP_TIEN_NT").ToString());
        int count1 = dataSet.Tables[0].Rows.Count;
        if (count1 > 0)
        {
          flag = true;
          int num = (int) ExMessageBox.Show(410, StartupBase.SasObj, "Có [" + (object) count1 + "] chứng từ. Tổng phát sinh [" + str2 + "] / [" + str1 + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }
        else
        {
          int num1 = (int) ExMessageBox.Show(415, StartupBase.SasObj, "Không có chứng từ nào như vậy!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }
        if (flag)
        {
          if (StartUp.M_AR_CK == 0 && !StartUp.HiddenFieldIsSetted)
          {
            StartUp.stringBrowse1 = StartUp.EditCkFields(StartUp.stringBrowse1);
            StartUp.stringBrowse2 = StartUp.EditCkFields(StartUp.stringBrowse2);
            StartUp.stringBrowse3 = StartUp.EditCkFields(StartUp.stringBrowse3);
            StartUp.stringBrowse4 = StartUp.EditCkFields(StartUp.stringBrowse4);
          }
          if (StartUp.M_KM_CK == 0 && !StartUp.HiddenFieldIsSetted)
          {
            StartUp.stringBrowse1 = StartUp.EditKmFields(StartUp.stringBrowse1);
            StartUp.stringBrowse2 = StartUp.EditKmFields(StartUp.stringBrowse2);
            StartUp.stringBrowse3 = StartUp.EditKmFields(StartUp.stringBrowse3);
            StartUp.stringBrowse4 = StartUp.EditKmFields(StartUp.stringBrowse4);
          }
          StartUp.HiddenFieldIsSetted = true;
          FormView formView = new FormView(this.SasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
          formView.ListFieldSum = "t_tt_nt;t_tt";
          formView.TongCongLabel = "Tổng thanh toán";
          if (StartUpTrans.M_LAN.Equals("V"))
            formView.frmBrw.Title = StartUp.M_Tilte + ". Ky: " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
          else
            formView.frmBrw.Title = StartUp.M_Tilte + ". Period: " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
          FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
          formView.frmBrw.LanguageID = "Socthda_8";
          formView.ShowDialog();
          StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString());
          int count2 = StartUpTrans.DsTrans.Tables[0].Rows.Count;
          int count3 = StartUpTrans.DsTrans.Tables[1].Rows.Count;
          for (int index = count2 - 1; index >= 1; --index)
            StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(index);
          for (int index = 0; index < count3; ++index)
            StartUpTrans.DsTrans.Tables[1].Rows.RemoveAt(0);
          int count4 = dataSet.Tables[0].Rows.Count;
          for (int index = 0; index < count4; ++index)
            StartUpTrans.DsTrans.Tables[0].Rows.Add(dataSet.Tables[0].Rows[index].ItemArray);
          int count5 = dataSet.Tables[1].Rows.Count;
          for (int index = 0; index < count5; ++index)
            StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);
          if (dataSet.Tables[0].Rows.Count > 0)
          {
            if (FrmSocthda.iRow > dataSet.Tables[0].Rows.Count - 1)
              FrmSocthda.iRow = dataSet.Tables[0].Rows.Count - 1;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
          }
          if (formView.DataGrid.ActiveRecord != null)
          {
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
              string stt_rec = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
              FrmSocthda.iRow = index + 1;
              StartUp.DataFilter(stt_rec);
            }
          }
          this.Close();
        }
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    private bool CheckValid()
    {
      bool flag = true;
      if (flag && (this.txtNgay_ct1.Value == null || this.txtNgay_ct1.Value.ToString() == ""))
      {
        int num = (int) ExMessageBox.Show(420, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtNgay_ct1.Focus();
        flag = false;
      }
      if (flag && !this.txtNgay_ct1.IsValueValid)
      {
        int num = (int) ExMessageBox.Show(425, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        flag = false;
        this.txtNgay_ct1.Focus();
        this.txtNgay_ct1.SelectAll();
      }
      if (flag && (this.txtNgay_ct2.Value == null || this.txtNgay_ct2.Value.ToString() == ""))
      {
        int num = (int) ExMessageBox.Show(430, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtNgay_ct2.Focus();
        flag = false;
      }
      if (flag && !this.txtNgay_ct2.IsValueValid)
      {
        int num = (int) ExMessageBox.Show(435, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        flag = false;
        this.txtNgay_ct2.Focus();
        this.txtNgay_ct2.SelectAll();
      }
      if (flag && Convert.ToDateTime(this.txtNgay_ct1.Value) > Convert.ToDateTime(this.txtNgay_ct2.Value))
      {
        int num = (int) ExMessageBox.Show(440, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        this.txtNgay_ct1.Focus();
        flag = false;
      }
      return flag;
    }

    private void txtloc_nsd_LostFocus(object sender, RoutedEventArgs e)
    {
      if (!(this.txtloc_nsd.Text.Trim() == ""))
        return;
      this.txtloc_nsd.Value = (object) 1;
    }

    private void txtMa_kh_LostFocus(object sender, RoutedEventArgs e)
    {
      if (this.txtMa_kh.RowResult != null)
      {
        if (StartUpTrans.M_LAN.Equals("V"))
          this.tblTen_kh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
        else
          this.tblTen_kh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
      }
      else
        this.tblTen_kh.Text = "";
    }

    private void txtMa_nx_LostFocus(object sender, RoutedEventArgs e)
    {
      if (this.txtMa_nx.RowResult != null)
      {
        if (StartUpTrans.M_LAN.Equals("V"))
          this.tblTen_tk.Text = this.txtMa_nx.RowResult["ten_tk"].ToString();
        else
          this.tblTen_tk.Text = this.txtMa_nx.RowResult["ten_tk2"].ToString();
      }
      else
        this.tblTen_tk.Text = "";
    }

    private void txtMaDVCS_LostFocus(object sender, RoutedEventArgs e)
    {
      if (this.txtMaDVCS.RowResult != null)
      {
        if (StartUpTrans.M_LAN.Equals("V"))
          this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
        else
          this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
      }
      else
        this.lblTenDVCS.Text = "";
    }

    private void txtMa_bp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      if (string.IsNullOrEmpty(this.txtMa_bp.Text.Trim()))
      {
        this.lbltenbp.Text = "";
      }
      else
      {
        if (this.txtMa_bp.RowResult == null)
          return;
        this.lbltenbp.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_bp.RowResult["ten_bp2"].ToString() : this.txtMa_bp.RowResult["ten_bp"].ToString();
      }
    }

    private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
      if (string.IsNullOrEmpty(this.txtMa_qs.Text.Trim()))
      {
        this.lbltenqs.Text = "";
      }
      else
      {
        if (this.txtMa_qs.RowResult == null)
          return;
        this.lbltenqs.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_qs.RowResult["ten_qs2"].ToString() : this.txtMa_qs.RowResult["ten_qs"].ToString();
      }
    }
  }
}
