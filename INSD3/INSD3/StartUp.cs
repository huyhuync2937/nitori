using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;

namespace INSD3
{
    public class StartUp : StartupBase
  {
    public static DataSet dsReport = new DataSet();
    public static DataTable tbInfo = new DataTable();
    public static int kindStyleReport = 0;
    public static string tableList = "v_ct70";
    public static string M_MA_DVCS = "";
    private static SqlCommand cmd = new SqlCommand();
    public static DataTable tbDetail = (DataTable) null;
    public static DataTable DtGroupInfo = (DataTable) null;
    public static DataTable GroupSelected = (DataTable) null;
    private static string SumFields = "";
    public static SasFormBrowes.FormBrowse oBrowse;
    public static FrmLoc frmLoc;
    public static DataRow commandInfo;
    public static DateTime M_ngay_ct0;
    public static string ngay;
    public static string M_ma_nt0;

    public override void Run()
    {
      StartupBase.Namespace = "INSD3";
      try
      {
        StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
        StartUp.M_ngay_ct0 = (DateTime) StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
        if (StartupBase.SasObj.DmdvcsInfo.Rows.Count > 0)
          StartUp.M_MA_DVCS = StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString();
        StartUp.commandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
        if (StartUp.commandInfo == null)
          return;
        StartUp.frmLoc = new FrmLoc();
        StartUp.frmLoc.Title = StartupBase.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUp.commandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUp.commandInfo["bar2"].ToString());
        StartUp.frmLoc.ShowDialog();
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    public static void CallGridVoucher(
      bool isFirstLoad,
      object EndDate,
      string strFilter,
      string strGopKhoDL,
      int kieu_xem)
    {
      DataTable dataTable;
      if (isFirstLoad)
      {
        StartUp.cmd.CommandText = " exec " + StartUp.commandInfo["store_proc"].ToString() + " @advfilter, @ngay_ct2, @ngay_ct0, @gop_kho_dl, @Type";
        StartUp.cmd.Parameters.Add("@advfilter", SqlDbType.NVarChar).Value = (object) strFilter;
        StartUp.cmd.Parameters.Add("@ngay_ct2", SqlDbType.Char, 16).Value = (object) string.Format("{0:yyyyMMdd}", (object) (DateTime) EndDate);
        StartUp.cmd.Parameters.Add("@ngay_ct0", SqlDbType.Char, 16).Value = (object) string.Format("{0:yyyyMMdd}", (object) StartUp.M_ngay_ct0.Date);
        StartUp.cmd.Parameters.Add("@gop_kho_dl", SqlDbType.Char).Value = (object) strGopKhoDL;
        StartUp.cmd.Parameters.Add("@Type", SqlDbType.Char).Value = (object) kieu_xem;
        StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
        StartUp.SumFields = "";
        dataTable = StartUp.dsReport.Tables[0].Copy();
        StartUp.dsReport.Tables[0].TableName = "tbDetail";
        StartUp.tbDetail = StartUp.dsReport.Tables[0].Copy();
        StartUp.tbInfo.Columns.Add("DateTime");
        StartUp.tbInfo.Rows.Add((object) StartUp.ngay);
        StartUp.tbInfo.TableName = "tbInfo";
        StartUp.dsReport.Tables.Add(StartUp.tbInfo);
        string stringBrowse = StartUp.GetStringBrowse(kieu_xem);
        StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, StartUp.tbDetail.DefaultView, stringBrowse);
        StartUp.oBrowse.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(StartUp.oBrowse_Esc);
        StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
        StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
        StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
        StartUp.oBrowse.frmBrw.Title = StartupBase.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUp.commandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUp.commandInfo["bar2"].ToString());
        if (kieu_xem == 2)
          StartUp.oBrowse.SetRowColorByTag("tag", "0");
      }
      else
      {
        StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
        dataTable = StartUp.dsReport.Tables[0].Copy();
        StartUp.tbDetail = StartUp.dsReport.Tables[0].Copy();
        StartUp.dsReport.Tables[0].TableName = "tbDetail";
        StartUp.tbDetail.TableName = "tbDetail";
        StartUp.dsReport.Tables.Add(StartUp.tbInfo.Copy());
        StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable) StartUp.tbDetail.DefaultView;
        StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
        StartUp.oBrowse.UpdateSumaryFields();
      }
      if (!isFirstLoad)
        return;
      StartUp.oBrowse.frmBrw.LanguageID = "INSD3_1";
      StartUp.oBrowse.ShowDialog();
      StartUp.frmLoc.Close();
    }

    private static string GetStringBrowse(int kieu_xem)
    {
      string[] strArray = StartUp.commandInfo[StartupBase.M_LAN == "V" ? "vbrowse1" : "ebrowse1"].ToString().Split('|');
      if (kieu_xem == 2)
        return strArray[1];
      string fieldFormat_sl = "{0}:H=SL mã kho {1}";
      string fieldFormat_du = "{0}:H=Tiền mã kho {1}";
      if (StartupBase.M_LAN == "V")
      {
        fieldFormat_sl = "{0}:H=SL mã kho {1}";
        fieldFormat_du = "{0}:H=Tiền mã kho {1}";
      }
      else
      {
        fieldFormat_sl = "{0}:H=Site stock {1}";
        fieldFormat_du = "{0}:H=Site amount {1}";
      }
      if (StartUp.dsReport.Tables[1].Rows.Count == 0)
        return strArray[0].Replace(";{1}", "").Replace(";{0}", "");
      return string.Join(";", ((IEnumerable<string>) strArray[0].Split(';')).Select<string, string>((Func<string, string>) (x =>
      {
        if (x.StartsWith("{0}"))
          return string.Join(";", StartUp.dsReport.Tables[1].Rows.Cast<DataRow>().Select<DataRow, string>((Func<DataRow, string>) (a => string.Format(x, (object) string.Format(fieldFormat_sl, (object) a["sl"].ToString().Trim(), (object) a["ma_kho"].ToString().Trim())))).ToArray<string>());
        if (!x.StartsWith("{1}"))
          return x;
        x = x.Replace("{1}", "{0}");
        return string.Join(";", StartUp.dsReport.Tables[1].Rows.Cast<DataRow>().Select<DataRow, string>((Func<DataRow, string>) (a => string.Format(x, (object) string.Format(fieldFormat_du, (object) a["du"].ToString().Trim(), (object) a["ma_kho"].ToString().Trim())))).ToArray<string>());
      })).ToArray<string>());
    }

    private static string GetTableShow(int mau_bc, bool isDetail)
    {
      string empty = string.Empty;
      string[] strArray = (string[]) null;
      switch (StartupBase.M_LAN)
      {
        case "V":
          switch (mau_bc)
          {
            case 0:
              strArray = StartUp.commandInfo["Vbrowse1"].ToString().Split('|');
              break;
            case 1:
              strArray = StartUp.commandInfo["Vbrowse2"].ToString().Split('|');
              break;
          }
          break;
        default:
          switch (mau_bc)
          {
            case 0:
              strArray = StartUp.commandInfo["Ebrowse1"].ToString().Split('|');
              break;
            case 1:
              strArray = StartUp.commandInfo["Ebrowse2"].ToString().Split('|');
              break;
          }
          break;
      }
      return isDetail ? strArray[1] : strArray[0];
    }

    public static string GetFieldList(string str1, string str2)
    {
      return "F" + str1 + ":H=" + str2;
    }

    private static void oBrowse_CTRL_R(object sender, EventArgs e)
    {
      StartUp.CallGridVoucher(false, StartUp.frmLoc.txtNgay.Value, StartUp.frmLoc.GetFilter() + StartUp.frmLoc.GetCondition(), StartUp.frmLoc.txtkho_dl.Text, int.Parse(StartUp.frmLoc.txtkieu_xem.Text));
    }

    public static void oBrowse_Esc(object sender, EventArgs e)
    {
    }

    public static void oBrowse_F7(object sender, EventArgs e)
    {
      DataSet LocalDataSet = StartUp.dsReport.Copy();
      new ReportManager(StartupBase.SasObj, StartUp.commandInfo["rep_file"].ToString(), 0).Preview(LocalDataSet);
    }

    private static void oBrowse_F11(object sender, EventArgs e)
    {
      INSD3F10 insD3F10 = new INSD3F10(StartUp.GroupSelected);
      insD3F10.DisplayLanguage = StartupBase.M_LAN;
      SysFunc.LoadIcon((Window) insD3F10);
      insD3F10.Title = SysFunc.Cat_Dau(insD3F10.Title);
      try
      {
        if (!insD3F10.ShowDialog())
          return;
        StartUp.GroupSelected = insD3F10.DataOption.Copy();
        if (StartUp.DtGroupInfo == null)
        {
          SqlCommand sqlcmd = new SqlCommand("Select loai_nh, ma_nh, ten_nh, ten_nh2 from dmnhvt");
          StartUp.DtGroupInfo = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
        }
        string GroupFields = "";
        string SortFields = "";
        int result1 = 0;
        int result2;
        int.TryParse(insD3F10.DataOption.Rows[0]["group1"].ToString(), out result2);
        int result3;
        int.TryParse(insD3F10.DataOption.Rows[0]["group2"].ToString(), out result3);
        int result4;
        int.TryParse(insD3F10.DataOption.Rows[0]["group3"].ToString(), out result4);
        int.TryParse(insD3F10.DataOption.Rows[0]["sortby"].ToString(), out result1);
        if (result2 != 0)
          GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + insD3F10.DataOption.Rows[0]["group1"].ToString();
        if (result3 != 0)
          GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + insD3F10.DataOption.Rows[0]["group2"].ToString();
        if (result4 != 0)
          GroupFields = GroupFields + (GroupFields == "" ? "" : ";") + insD3F10.DataOption.Rows[0]["group3"].ToString();
        switch (result1)
        {
          case 0:
            SortFields = StartupBase.M_LAN == "V" ? "ten_vt" : "ten_vt2";
            break;
          case 1:
            SortFields = "ma_vt";
            break;
        }
        SysFunc.InsertListGroup(StartUp.tbDetail, StartUp.DtGroupInfo, "nh_vt", GroupFields, "ten_vt;ten_vt2", SortFields, StartUp.SumFields);
        StartUp.dsReport.Tables[1].Rows.Clear();
        StartUp.dsReport.Tables[1].Merge(StartUp.tbDetail.Copy());
        StartUp.dsReport.Tables[1].DefaultView.Sort = StartUp.tbDetail.DefaultView.Sort;
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    public static string fieldShow(int kindReport)
    {
      string empty = string.Empty;
      switch (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper())
      {
        case "V":
          switch (kindReport)
          {
            case 1:
              empty = StartUp.commandInfo["Vbrowse1"].ToString();
              break;
            case 2:
              empty = StartUp.commandInfo["Vbrowse2"].ToString();
              break;
          }
          break;
        default:
          switch (kindReport)
          {
            case 1:
              empty = StartUp.commandInfo["Ebrowse1"].ToString();
              break;
            case 2:
              empty = StartUp.commandInfo["Ebrowse2"].ToString();
              break;
          }
          break;
      }
      return empty;
    }
  }
}
