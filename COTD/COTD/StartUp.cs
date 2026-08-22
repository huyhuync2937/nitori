using Infragistics.Windows.DataPresenter;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace COTD
{
  public class StartUp : StartupBase
  {
    public static string sqlTableName = "dmtd";
    public static string sqlTableView = "v_dmtd";
    public static string SqlTableKey = "ma_td";
    public static string SqlTableObjectName = "ten_td";
    public static ActionTask currActionTask = ActionTask.None;
    public static string currSqlTableKey = string.Empty;
    public static string titleWindow = string.Empty;
    public static bool IsLoadFromLookUp = false;
    public static string _parameter = "";
    public static DataTable LastEditTable = (DataTable) null;
    public static string _menu_id = "";
    public static DataRow CommandInfo;
    public static string TableName;
    public static SasFormBrowes.FrameBrowse oBrowse;
    public static string M_ma_nt0;

    public override void Run()
    {
      StartupBase.Namespace = "COTD";
      try
      {
        StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
        StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
        CotdLoc cotdLoc = new CotdLoc();
        StartUp._parameter = StartUp.CommandInfo["parameter"].ToString().Trim();
        cotdLoc.Title = SysFunc.Cat_Dau(StartUp.CommandInfo[StartupBase.M_LAN.Equals("V") ? "bar" : "bar2"].ToString());
        cotdLoc.ShowDialog();
        if (cotdLoc.isClose)
          return;
        StartUp._parameter = cotdLoc.txtTd.Value.ToString();
        if (!StartUp._parameter.Equals("1"))
        {
          StartUp.sqlTableName = StartUp.sqlTableName + StartUp._parameter + "0";
          StartUp.sqlTableView += StartUp._parameter;
        }
        else
          StartUp.sqlTableName += "0";
        StartUp.oBrowse = new SasFormBrowes.FrameBrowse(StartupBase.SasObj, StartUp.sqlTableName);
        StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartUp.CommandInfo[StartupBase.M_LAN.Equals("V") ? "bar" : "bar2"].ToString().Trim()) + " " + StartUp._parameter;
        StartUp.TableName = StartupBase.M_LAN.Equals("V") ? "trường tự do" : "free field";
        SysFunc.LoadIcon((Window) StartUp.oBrowse.frmBrw);
        StartUp.oBrowse.F2 += new SasFormBrowes.FrameBrowse.GridKeyUp_F2(this.oBrowse_F2);
        StartUp.oBrowse.F3 += new SasFormBrowes.FrameBrowse.GridKeyUp_F3(this.oBrowse_F3);
        StartUp.oBrowse.F4 += new SasFormBrowes.FrameBrowse.GridKeyUp_F4(this.oBrowse_F4);
        StartUp.oBrowse.Ctrl_F4 += new SasFormBrowes.FrameBrowse.GridKeyUp_Ctrl_F4(this.oBrowse_Ctrl_F4);
        StartUp.oBrowse.F6 += new SasFormBrowes.FrameBrowse.GridKeyUp_F6(this.oBrowse_F6);
        StartUp.oBrowse.F7 += new SasFormBrowes.FrameBrowse.GridKeyUp_F7(this.oBrowse_F7);
        StartUp.oBrowse.F8 += new SasFormBrowes.FrameBrowse.GridKeyUp_F8(this.oBrowse_F8);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_Moi += new ListToolBar.Execute(this.ToolBar_Cm_Moi);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_Sua += new ListToolBar.Execute(this.ToolBar_Cm_Sua);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_Xoa += new ListToolBar.Execute(this.ToolBar_Cm_Xoa);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_In += new ListToolBar.Execute(this.ToolBar_Cm_In);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_Copy += new ListToolBar.Execute(this.ToolBar_Cm_Copy);
        StartUp.oBrowse.frmBrw.ToolBar.Cm_DoiMa += new ListToolBar.Execute(this.ToolBar_Cm_DoiMa);
        StartUp.oBrowse.frmBrw.ShowInTaskbar = true;
        StartUp.oBrowse.frmBrw.LanguageID = "COTDBrowse";
        StartUp.oBrowse.ShowDialog();
        cotdLoc.Close();
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    private void oBrowse_Ctrl_F4(object sender, EventArgs e)
    {
      this.Copy();
    }

    private void ToolBar_Cm_DoiMa()
    {
      this.DoiMa();
    }

    private void ToolBar_Cm_Copy()
    {
      this.Copy();
    }

    private void ToolBar_Cm_In()
    {
      this.In();
    }

    private void ToolBar_Cm_Xoa()
    {
      this.Xoa();
    }

    private void ToolBar_Cm_Sua()
    {
      this.Sua();
    }

    private void ToolBar_Cm_Moi()
    {
      this.Them();
    }

    private void oBrowse_F2(object sender, EventArgs e)
    {
      if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
        return;
      StartUp.currActionTask = ActionTask.View;
      StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
      StartUp.titleWindow = !(StartupBase.M_LAN == "V") ? string.Format("View free field information {0}", (object) StartUp._parameter) : string.Format("Xem thong tin truong tu do {0}", (object) StartUp._parameter);
      this.showWindow();
    }

    private void oBrowse_F3(object sender, EventArgs e)
    {
      this.Sua();
    }

    private void Sua()
    {
      if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
        return;
      StartUp.currActionTask = ActionTask.Edit;
      StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
      StartUp.titleWindow = !(StartupBase.M_LAN == "V") ? string.Format("Edit free field {0}", (object) StartUp._parameter) : string.Format("Sua truong tu do {0}", (object) StartUp._parameter);
      this.showWindow();
    }

    private void oBrowse_F4(object sender, EventArgs e)
    {
      this.Them();
    }

    private void Them()
    {
      if (StartUp.currActionTask != ActionTask.None)
        return;
      StartUp.currActionTask = ActionTask.Add;
      StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord != null && StartUp.oBrowse.ActiveRecord.RecordType == RecordType.DataRecord ? StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim() : "";
      StartUp.titleWindow = !(StartupBase.M_LAN == "V") ? string.Format("Add free field {0}", (object) StartUp._parameter) : string.Format("Them truong tu do {0}", (object) StartUp._parameter);
      this.showWindow();
    }

    private void Copy()
    {
      if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
        return;
      StartUp.currActionTask = ActionTask.Copy;
      StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
      StartUp.titleWindow = !(StartupBase.M_LAN == "V") ? string.Format("Add free field {0}", (object) StartUp._parameter) : string.Format("Them truong tu do {0}", (object) StartUp._parameter);
      this.showWindow();
    }

    private void oBrowse_F6(object sender, EventArgs e)
    {
      this.DoiMa();
    }

    private void DoiMa()
    {
      if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
        return;
      DataRecord activeRecord = StartUp.oBrowse.ActiveRecord;
      activeRecord.Cells[StartUp.SqlTableKey].Value = (object) activeRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
      string[] TitleAndName = new string[2]
      {
        StartupBase.M_LAN.Equals("V") ? "Doi ma " + StartUp.TableName : "Change " + StartUp.TableName + " code",
        StartupBase.M_LAN.Equals("V") ? "ten_td" : "ten_td2"
      };
      SysFunc.ChangeListID(StartupBase.SasObj, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1), StartUp.oBrowse, StartUp.SqlTableKey, TitleAndName);
    }

    private void oBrowse_F7(object sender, EventArgs e)
    {
      this.In();
    }

    private void In()
    {
      SqlCommand sqlcmd = new SqlCommand("select " + SysFunc.GetFieldFromStrBrowse(StartUp.oBrowse.Listinfo["full_field"].ToString()) + " from " + StartUp.sqlTableView);
      DataSet datasetreport = StartupBase.SasObj.ExcuteReader(sqlcmd);
      ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.CommandInfo["rep_file"].ToString());
      SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.GetAllData(), ref datasetreport, 0);
      reportManager.Preview(datasetreport);
    }

    private void oBrowse_F8(object sender, EventArgs e)
    {
      this.Xoa();
    }

    private void Xoa()
    {
      if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
        return;
      StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString();
      StartUp.currActionTask = ActionTask.Delete;
      if (SysFunc.CheckPermission(StartupBase.SasObj, StartUp.currActionTask, StartupBase.Menu_Id))
      {
        try
        {
          SqlCommand sqlcmd = new SqlCommand("exec dbo.CheckDeleteListId @ma_dm, @" + StartUp.SqlTableKey + ", @notvalue");
          sqlcmd.Parameters.Add("@ma_dm", SqlDbType.VarChar).Value = (object) StartUp.sqlTableName;
          sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.VarChar).Value = (object) StartUp.currSqlTableKey;
          if (StartUp.sqlTableName.Equals("dmtd0"))
            sqlcmd.Parameters.Add("@notvalue", SqlDbType.VarChar).Value = (object) "dmtd2;dmtd3";
          else if (StartUp.sqlTableName.Equals("dmtd20"))
            sqlcmd.Parameters.Add("@notvalue", SqlDbType.VarChar).Value = (object) "dmtd;dmtd3";
          else if (StartUp.sqlTableName.Equals("dmtd30"))
            sqlcmd.Parameters.Add("@notvalue", SqlDbType.VarChar).Value = (object) "dmtd2;dmtd";
          if ((int) StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0)
          {
            if (ExMessageBox.Show(3675, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
              this.deleteRowByKey(StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1));
          }
          else
          {
            int num = (int) ExMessageBox.Show(3680, StartupBase.SasObj, "Đã có phát sinh không được xóa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
          }
        }
        catch (SqlException ex)
        {
          ErrorLog.CatchMessage(ex);
        }
      }
      else
      {
        int num1 = (int) ExMessageBox.Show(3685, StartupBase.SasObj, "Không có quyền xóa [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
      }
      StartUp.currActionTask = ActionTask.None;
      StartUp.oBrowse.frmBrw.ReloadData();
    }

    private void deleteRowByKey(string tableName)
    {
      DataTable row = StartUp.GetRow(tableName);
      if (row.Rows.Count <= 0)
        return;
      ListFunc.deleteRowInDatabaseByKey(tableName, StartUp.SqlTableKey, row.Rows[0], StartupBase.SasObj);
    }

    public DataTable Extend_oBrowse_Command(
      SasObject SasObj,
      ActionTask Mode,
      string Current_Key)
    {
      StartUp.currActionTask = Mode;
      if (SysFunc.CheckPermission(SasObj, StartUp.currActionTask, StartupBase.Menu_Id))
      {
        StartUp.TableName = StartupBase.M_LAN.Equals("V") ? "tu do" : "Free field";
        StartUp.IsLoadFromLookUp = true;
        try
        {
          StartupBase.SasObj = SasObj;
          StartUp.currSqlTableKey = Current_Key.Trim();
          switch (Mode)
          {
            case ActionTask.View:
              StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thong tin truong " + StartUp.TableName : "View free field information";
              break;
            case ActionTask.Add:
            case ActionTask.Copy:
              StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them truong " + StartUp.TableName : "Add free field";
              break;
            case ActionTask.Edit:
              StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sua thong tin truong " + StartUp.TableName : "Edit free field information";
              break;
          }
          this.showWindow();
        }
        catch (Exception ex)
        {
          ErrorLog.CatchMessage(ex);
        }
      }
      else
      {
        int num = (int) ExMessageBox.Show(3690, SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        StartUp.currActionTask = ActionTask.None;
      }
      return StartUp.LastEditTable;
    }

    public DataTable Extend_oBrowse_Command_tudo1(
      SasObject SasObj,
      ActionTask Mode,
      string Current_Ma_kh)
    {
      StartUp.sqlTableName = "dmtd0";
      StartUp.sqlTableView = "v_dmtd";
      StartUp._parameter = "1";
      return this.Extend_oBrowse_Command(SasObj, Mode, Current_Ma_kh);
    }

    public DataTable Extend_oBrowse_Command_tudo2(
      SasObject SasObj,
      ActionTask Mode,
      string Current_Ma_kh)
    {
      StartUp.sqlTableName = "dmtd20";
      StartUp.sqlTableView = "v_dmtd2";
      StartUp._parameter = "2";
      return this.Extend_oBrowse_Command(SasObj, Mode, Current_Ma_kh);
    }

        public DataTable Extend_oBrowse_Command_tudo4(
     SasObject SasObj,
     ActionTask Mode,
     string Current_Ma_kh)
        {
            StartUp.sqlTableName = "dmtd40";
            StartUp.sqlTableView = "v_dmtd4";
            StartUp._parameter = "4";
            return this.Extend_oBrowse_Command(SasObj, Mode, Current_Ma_kh);
        }

        public DataTable Extend_oBrowse_Command_tudo3(
      SasObject SasObj,
      ActionTask Mode,
      string Current_Ma_kh)
    {
      StartUp.sqlTableName = "dmtd30";
      StartUp.sqlTableView = "v_dmtd3";
      StartUp._parameter = "3";
      return this.Extend_oBrowse_Command(SasObj, Mode, Current_Ma_kh);
    }

        public DataTable Extend_oBrowse_Command_tudo5(
     SasObject SasObj,
     ActionTask Mode,
     string Current_Ma_kh)
        {
            StartUp.sqlTableName = "dmtd50";
            StartUp.sqlTableView = "v_dmtd5";
            StartUp._parameter = "5";
            return this.Extend_oBrowse_Command(SasObj, Mode, Current_Ma_kh);
        }

        private void showWindow()
    {
      if (SysFunc.CheckPermission(StartupBase.SasObj, StartUp.currActionTask, StartupBase.Menu_Id))
      {
        newForm newForm = new newForm();
        switch (StartUp.currActionTask)
        {
          case ActionTask.View:
            newForm.EnableEditMode(false);
            break;
          case ActionTask.Add:
            newForm.EnableEditMode(true);
            break;
          case ActionTask.Edit:
            newForm.EnableEditMode(true);
            break;
          case ActionTask.Copy:
            newForm.EnableEditMode(true);
            break;
        }
        newForm.ShowDialog();
        if (StartUp.currActionTask == ActionTask.View || (StartUp.LastEditTable == null || StartUp.oBrowse == null))
          return;
        StartUp.oBrowse.frmBrw.ReloadData(StartUp.LastEditTable.Rows[0]);
      }
      else
      {
        int num = (int) ExMessageBox.Show(3695, StartupBase.SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        StartUp.currActionTask = ActionTask.None;
      }
    }

    public static DataTable GetRow(string sqlTableName)
    {
      DataTable dataTable = (DataTable) null;
      SqlCommand sqlcmd = new SqlCommand();
      try
      {
        string str = "select * from " + sqlTableName + " where ";
        if (StartUp.currActionTask == ActionTask.Add)
        {
          str += "1=0";
        }
        else
        {
          DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, sqlTableName);
          string[] strArray1 = StartUp.currSqlTableKey.Split(';');
          string[] strArray2 = StartUp.SqlTableKey.Split(';');
          for (int index = 0; index < ((IEnumerable<string>) strArray2).Count<string>(); ++index)
          {
            if (index > 0)
              str += " and";
            DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + strArray2[index] + "'");
            str = str + " " + strArray2[index] + "=@" + strArray2[index];
            sqlcmd.Parameters.Add("@" + strArray2[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString())).Value = (object) strArray1[index];
          }
        }
        sqlcmd.CommandText = str;
        dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
      }
      catch (SqlException ex)
      {
        ErrorLog.CatchMessage(ex);
      }
      return dataTable;
    }
  }
}
