using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace POBK1_2
{
    public class StartUp : StartupBase
  {
    public static DataSet DataSourceReport = new DataSet();
    public static string sqlTableView = "v_pobk1";
    public static string SqlTableKey = "stt_rec";
    public static string SqlTableObjectName = "ten_kh";
    private static SqlCommand cmd1 = new SqlCommand();
    private static SqlCommand cmd = new SqlCommand();
    public static string TableName = "ct70";
    public static DataTable dtInfo;
    private static SasFormBrowes.FormBrowseDynamic.FormBrowseDynamic oBrowse;
    public static DataRow CommandInfo;
    public static DateTime M_ngay_ct0;
    public static string M_ma_nt0;
    private static FormLoc _frmLoc;

    public override void Run()
    {
      StartupBase.Namespace = "POBK1_2";
      try
      {
        StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
        DateTime now1 = DateTime.Now;
        StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
        StartUp.M_ngay_ct0 = (DateTime) StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
        StartUp.dtInfo = new DataTable();
        StartUp.dtInfo.TableName = "TbInfo";
        StartUp.dtInfo.Columns.Add("StartDate");
        StartUp.dtInfo.Columns.Add("EndDate");
        if (StartUp.CommandInfo == null)
          return;
        StartUp._frmLoc = new FormLoc();
        StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString());
        if (StartupBase.M_LAN != "V")
          StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
        DateTime now2 = DateTime.Now;
        StartUp._frmLoc.ShowDialog();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.Message);
      }
    }

    public static void CallGridVouchers(
      object StartDate,
      object EndDate,
      string filter,
      string MaVT)
    {
      try
      {
        StartUp.oBrowse = new SasFormBrowes.FormBrowseDynamic.FormBrowseDynamic(StartupBase.SasObj, "POBK1_2.exe", StartUp.sqlTableView);
        string[] strArray = StartUp.CommandInfo["store_proc"].ToString().Split('|');
        StartUp.cmd = new SqlCommand();
        StartUp.cmd.CommandText = "Exec " + strArray[0] + " @StartDate, @EndDate, @Condition, @Select, @Join";
        StartUp.cmd.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartDate.ToString()) ? (object) "" : (object) string.Format("{0:yyyyMMdd}", (object) (DateTime) StartDate);
        StartUp.cmd.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = string.IsNullOrEmpty(EndDate.ToString()) ? (object) "" : (object) string.Format("{0:yyyyMMdd}", (object) (DateTime) EndDate);
        StartUp.cmd.Parameters.Add("@Condition", SqlDbType.NVarChar).Value = (object) filter;
        StartUp.cmd.Parameters.Add("@Select", SqlDbType.NVarChar).Value = (object) StartUp.oBrowse.SelectedFields;
        StartUp.cmd.Parameters.Add("@Join", SqlDbType.NVarChar).Value = (object) StartUp.oBrowse.JoinClause;
        DataTable dataTable = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
        if (string.IsNullOrEmpty(StartUp.oBrowse.BrowseFields))
          StartUp.oBrowse.BrowseFields = !StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["EBrowse1"].ToString().Trim() : StartUp.CommandInfo["VBrowse1"].ToString().Trim();
        StartUp.oBrowse.ObrowseView = dataTable.DefaultView;
        StartUp.oBrowse.frmBrw.Loaded += (RoutedEventHandler) ((s, e) => StartUp.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.LabelClickAction = LabelClickAction.SortByMultipleFields)));
        StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowseDynamic.FormBrowseDynamic.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
        StartUp.oBrowse.LoadTemplate += new SasFormBrowes.FormBrowseDynamic.FormBrowseDynamic.GridLoadTemplate(StartUp.oBrowse_LoadTemplate);
        StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
        StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
        if (StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport") is ToolBar name)
        {
          ToolBarButton toolBarButton = new ToolBarButton();
          toolBarButton.Text = StartupBase.M_LAN.Equals("V") ? "Màn hình lọc" : "Filter Windows";
          toolBarButton.Name = "btnSelect";
          toolBarButton.ImagePath = "Images\\Search.png";
          toolBarButton.BorderBrush = (Brush) null;
          toolBarButton.Click += (RoutedEventHandler) ((s, e) =>
          {
            StartUp._frmLoc.isFirstLoad = false;
            StartUp._frmLoc.Visibility = Visibility.Visible;
            StartUp._frmLoc.Closing += (CancelEventHandler) ((ss, ee) =>
            {
              ee.Cancel = true;
              StartUp._frmLoc.Visibility = Visibility.Hidden;
            });
          });
          name.Items.Insert(1, (object) toolBarButton);
        }
        StartUp.oBrowse.frmBrw.LanguageID = "POBK1_2_1";
        StartUp.oBrowse.ShowDialog();
        Application.Current.Shutdown();
      }
      catch (Exception ex)
      {
        int num = (int) MessageBox.Show(ex.InnerException.Message);
      }
    }

    private static void oBrowse_LoadTemplate(string vbrowse, string select, string join)
    {
      StartUp.cmd.Parameters["@Select"].Value = (object) select;
      StartUp.cmd.Parameters["@Join"].Value = (object) join;
      DataTable dt = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
      StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts.Clear();
      FieldLayout fieldLayout = SysFunc.CreateFieldLayout(StartupBase.SasObj, StartUp.oBrowse.frmBrw.oBrowse as BasicGridView, vbrowse, dt);
      StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts.Add(fieldLayout);
      StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable) dt.DefaultView;
      StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
      StartUp.oBrowse.UpdateSumaryFields();
    }

    public static void QueryData(
      bool isFirstLoad,
      object StartDate,
      object EndDate,
      string filter,
      string MaVT)
    {
      try
      {
        if (isFirstLoad)
        {
          StartUp.CallGridVouchers(StartDate, EndDate, filter, MaVT);
        }
        else
        {
          StartUp.cmd.Parameters["@StartDate"].Value = string.IsNullOrEmpty(StartDate.ToString()) ? (object) "" : (object) string.Format("{0:yyyyMMdd}", (object) (DateTime) StartDate);
          StartUp.cmd.Parameters["@EndDate"].Value = string.IsNullOrEmpty(EndDate.ToString()) ? (object) "" : (object) string.Format("{0:yyyyMMdd}", (object) (DateTime) EndDate);
          StartUp.cmd.Parameters["@Condition"].Value = (object) filter;
          StartUp.cmd.Parameters["@Select"].Value = (object) StartUp.oBrowse.SelectedFields;
          StartUp.cmd.Parameters["@Join"].Value = (object) StartUp.oBrowse.JoinClause;
          DataTable dataTable = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
          StartUp.oBrowse.ObrowseView = dataTable.DefaultView;
          StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
          StartUp.oBrowse.UpdateSumaryFields();
        }
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    private static void oBrowse_CTRL_R(object sender, EventArgs e)
    {
      string filter = StartUp._frmLoc.GetFilter();
      StartUp.QueryData(false, StartUp._frmLoc.TxtStartDateTime.Value, StartUp._frmLoc.TxtEndDateTime.Value, filter, StartUp._frmLoc.txtMaVT.Text.Trim());
    }

    private static void oBrowse_Esc(object sender, EventArgs e)
    {
      StartUp.oBrowse.frmBrw.Close();
      if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
        return;
      Application.Current.Shutdown();
    }
  }
}
