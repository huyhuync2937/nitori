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

namespace DMBPNS
{
    public class StartUp : StartupBase
    {
        public static string sqlTableListName = "v_DMBPNS";
        public static string sqlTableName = "DMBPNS";
        public static string sqlTableView = "v_DMBPNS";
        public static string SqlTableKey = "ma_bpns";
        public static string SqlTableObjectName = "ma_bpns";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;     
        public static string titleWindow = string.Empty;
        public static bool IsLoadFromLookUp = false;
        public static DataRow LastEditRow = (DataRow)null;   
        public static string TableName;
        public static DataRow CommandInfo;
        public static SasFormBrowes.FrameBrowse oBrowse;


        public override void Run()
        {
            StartupBase.Namespace = "DMBPNS";
            try
            {
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
                StartUp.TableName = StartupBase.M_LAN.Equals("V") ? "bộ phận nhân sự" : "Department";
                StartUp.oBrowse = new SasFormBrowes.FrameBrowse(StartupBase.SasObj, StartUp.sqlTableName);
                StartUp.oBrowse.frmBrw.Title = StartUp.titleWindow;
                SysFunc.LoadIcon((Window)StartUp.oBrowse.frmBrw);
                StartUp.oBrowse.F2 += new SasFormBrowes.FrameBrowse.GridKeyUp_F2(this.oBrowse_F2);
                StartUp.oBrowse.F3 += new SasFormBrowes.FrameBrowse.GridKeyUp_F3(this.oBrowse_F3);
                StartUp.oBrowse.F4 += new SasFormBrowes.FrameBrowse.GridKeyUp_F4(this.oBrowse_F4);
              
                StartUp.oBrowse.Ctrl_F4 += new SasFormBrowes.FrameBrowse.GridKeyUp_Ctrl_F4(this.oBrowse_Ctrl_F4);
                StartUp.oBrowse.F8 += new SasFormBrowes.FrameBrowse.GridKeyUp_F8(this.oBrowse_F8);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Moi += new ListToolBar.Execute(this.ToolBar_Cm_Moi);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Sua += new ListToolBar.Execute(this.ToolBar_Cm_Sua);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Xoa += new ListToolBar.Execute(this.ToolBar_Cm_Xoa);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Copy += new ListToolBar.Execute(this.ToolBar_Cm_Copy);
                StartUp.oBrowse.F6 += new SasFormBrowes.FrameBrowse.GridKeyUp_F6(this.oBrowse_F6);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_DoiMa += new ListToolBar.Execute(this.ToolBar_Cm_DoiMa);

                StartUp.oBrowse.frmBrw.LanguageID = "DMBPNS_1";
                StartUp.oBrowse.ShowDialog();
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
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName + " information";
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
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName + " information";
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
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord == null ? "" : StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add new " + StartUp.TableName;
            this.showWindow();
        }

        private void Copy()
        {
            if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.currActionTask = ActionTask.Copy;
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord == null ? "" : StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add new " + StartUp.TableName;
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
            activeRecord.Cells[StartUp.SqlTableKey].Value = (object)activeRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
            string[] TitleAndName = new string[2]
            {
        StartupBase.M_LAN.Equals("V") ? "Doi ma " + StartUp.TableName : "Change " + StartUp.TableName + " code",
        StartupBase.M_LAN.Equals("V") ? "ten_bpns" : "ten_bpns2"
            };
            SysFunc.ChangeListID(StartupBase.SasObj, StartUp.sqlTableName, StartUp.oBrowse, StartUp.SqlTableKey, TitleAndName);
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
                        SqlCommand sqlcmd = new SqlCommand("exec dbo.CheckDeleteListId @ma_dm, @" + StartUp.SqlTableKey);
                        sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)StartUp.sqlTableName;
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)StartUp.currSqlTableKey;
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0)
                        {
                            if (ExMessageBox.Show(2335, StartupBase.SasObj, "Có chắc chắn xóa không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                                this.deleteRowByKey(StartUp.sqlTableName);
                        }
                        else
                        {
                            int num = (int)ExMessageBox.Show(2340, StartupBase.SasObj, "Đã có phát sinh không được xóa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                    }
                    catch (SqlException ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(2345, StartupBase.SasObj, "Không có quyền xóa [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Question);
                }
                StartUp.currActionTask = ActionTask.None;
                StartUp.oBrowse.frmBrw.ReloadData();    
        }
        private void deleteRowByKey(string tableName)
        {
            DataTable row = StartUp.GetRow(tableName);
            if (row.Rows.Count <= 0 || ListFunc.deleteRowInDatabaseByKey(tableName, StartUp.SqlTableKey, row.Rows[0], StartupBase.SasObj) <= 0)
                return;
            string sql = "Delete FROM " + StartUp.sqlTableName + " WHERE " + StartUp.SqlTableKey + " LIKE @curma;"  ;
            SqlCommand sqlcmd = new SqlCommand(sql);
            sqlcmd.Parameters.Add("@curma", SqlDbType.VarChar).Value = (object)row.Rows[0][StartUp.SqlTableKey].ToString().Trim();           
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        public DataTable Extend_oBrowse_Command(
          SasObject SasObj,
          ActionTask Mode,
          string Current_Ma)
        {
            StartUp.TableName = StartupBase.M_LAN.Equals("V") ? "bộ phận nhân sự" : "item";
            StartUp.currActionTask = Mode;
            if (SysFunc.CheckPermission(SasObj, StartUp.currActionTask, StartupBase.Menu_Id))
            {
                StartUp.IsLoadFromLookUp = true;
                switch (Mode)
                {
                    case ActionTask.View:
                        try
                        {
                            StartupBase.SasObj = SasObj;
                            StartUp.currSqlTableKey = Current_Ma;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName + " information";
                            this.showWindow();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                    case ActionTask.Add:
                        try
                        {
                            StartupBase.SasObj = SasObj;
                            StartUp.currSqlTableKey = Current_Ma;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add new " + StartUp.TableName;
                            this.showWindow();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                    case ActionTask.Edit:
                        try
                        {
                            StartupBase.SasObj = SasObj;
                            StartUp.currSqlTableKey = Current_Ma;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName + " information";
                            this.showWindow();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                    case ActionTask.Copy:
                        try
                        {
                            StartupBase.SasObj = SasObj;
                            StartUp.currSqlTableKey = Current_Ma;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add new " + StartUp.TableName;
                            this.showWindow();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                }
            }
            else
            {
                int num = (int)ExMessageBox.Show(2350, SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                StartUp.currActionTask = ActionTask.None;
            }
            return StartUp.LastEditRow == null ? (DataTable)null : StartUp.LastEditRow.Table;
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
                        newForm.EnableEditMode(false);
                        break;
                    case ActionTask.Copy:
                        newForm.EnableEditMode(true);
                        break;
                }
                newForm.ShowDialog();
                if (StartUp.currActionTask == ActionTask.View || (StartUp.LastEditRow == null || StartUp.oBrowse == null))
                    return;
                StartUp.oBrowse.frmBrw.ReloadData(StartUp.LastEditRow);
            }
            else
            {
                int num = (int)ExMessageBox.Show(2355, StartupBase.SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                StartUp.currActionTask = ActionTask.None;
            }
        }
       
        public static string CreateSqlfilter(string _listSqlKey, string _listSQlKeyValue)
        {
            string[] strArray1 = _listSqlKey.Split(';');
            string[] strArray2 = _listSQlKeyValue.Split(';');
            string str = "";
            for (int index = 0; index < strArray1.Length; ++index)
                str = str + " " + strArray1[index] + " = '" + strArray2[index] + "' and";
            return str.Length > 3 ? str.Substring(0, str.Length - 3) : string.Empty;
        }
        public static DataTable GetRow(string sqlTableName)
        {
            DataTable dataTable = (DataTable)null;
            SqlCommand sqlcmd = new SqlCommand();
            try
            {
                string str = "select * from " + sqlTableName + " where ";
                DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, sqlTableName);
                string[] strArray1 = StartUp.currSqlTableKey.Split(';');
                string[] strArray2 = StartUp.SqlTableKey.Split(';');
                for (int index = 0; index < ((IEnumerable<string>)strArray2).Count<string>(); ++index)
                {
                    if (index > 0)
                        str += " and";
                    DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + strArray2[index] + "'");
                    str = str + " " + strArray2[index] + "=@" + strArray2[index];
                    sqlcmd.Parameters.Add("@" + strArray2[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString())).Value = (object)strArray1[index];
                }
                sqlcmd.CommandText = str;
                dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                Console.WriteLine("sql: {0}", str);
            }
            catch (SqlException ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return dataTable;
        }
    }
}
