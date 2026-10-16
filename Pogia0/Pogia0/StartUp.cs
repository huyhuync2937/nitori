using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;

namespace Pogia0
{
    public class StartUp : StartupBase
    {
        public static string sqlTableName = "dmgia0";
        public static string sqlTableView = "v_dmgia0";
        public static string SqlTableKey = "ma_vt;ngay_nhap";
        public static string SqlTableObjectName = "ten_vt";
        public static string TableName = "giá nhập";
        public static string TableName_CatDau = "gia nhap";
        public static string TableName_CatDau2 = "purchase price";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;
        public static string titleWindow = string.Empty;
        public static bool IsLoadFromLookUp = false;
        public static DataTable LastEditTable = (DataTable)null;
        public static DataRow CommandInfo;
        public static FrameBrowse oBrowse;

        public override void Run()
        {
            StartupBase.Namespace = "Pogia0";
            try
            {
                StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.oBrowse = new FrameBrowse(StartupBase.SasObj, StartUp.sqlTableName);
                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
                SysFunc.LoadIcon((Window)StartUp.oBrowse.frmBrw);
                StartUp.oBrowse.F2 += new FrameBrowse.GridKeyUp_F2(this.oBrowse_F2);
                StartUp.oBrowse.F3 += new FrameBrowse.GridKeyUp_F3(this.oBrowse_F3);
                StartUp.oBrowse.F4 += new FrameBrowse.GridKeyUp_F4(this.oBrowse_F4);
                StartUp.oBrowse.Ctrl_F4 += new FrameBrowse.GridKeyUp_Ctrl_F4(this.oBrowse_Ctrl_F4);
                StartUp.oBrowse.F7 += new FrameBrowse.GridKeyUp_F7(this.oBrowse_F7);
                StartUp.oBrowse.F8 += new FrameBrowse.GridKeyUp_F8(this.oBrowse_F8);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Moi += new ListToolBar.Execute(this.ToolBar_Cm_Moi);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Sua += new ListToolBar.Execute(this.ToolBar_Cm_Sua);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Xoa += new ListToolBar.Execute(this.ToolBar_Cm_Xoa);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_In += new ListToolBar.Execute(this.ToolBar_Cm_In);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Copy += new ListToolBar.Execute(this.ToolBar_Cm_Copy);
                StartUp.oBrowse.frmBrw.ShowInTaskbar = true;
                this.UpdateFieldSearch();
                StartUp.oBrowse.frmBrw.LanguageID = "Pogia0_1";
                StartUp.oBrowse.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public void UpdateFieldSearch()
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("Exec [dbo].[SMDM_UPDATE_SEARCH_dmgia0] @tableName, @tableView, @listKey, @strKey, @value");
                sqlcmd.Parameters.Add("@tableName", SqlDbType.VarChar).Value = (object)"";
                sqlcmd.Parameters.Add("@tableView", SqlDbType.VarChar).Value = (object)"";
                sqlcmd.Parameters.Add("@listKey", SqlDbType.VarChar).Value = (object)"";
                sqlcmd.Parameters.Add("@strKey", SqlDbType.VarChar).Value = (object)"";
                sqlcmd.Parameters.Add("@value", SqlDbType.NVarChar).Value = (object)"";
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
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
            StartUp.currSqlTableKey = this.strKeyValue();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName_CatDau2 + " information";
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
            StartUp.currSqlTableKey = this.strKeyValue();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName_CatDau2 + " information";
            this.showWindow();
        }

        private string strKeyValue()
        {
            string[] strArray = StartUp.SqlTableKey.Split(';');
            string str = string.Empty;
            string empty = string.Empty;
            for (int index1 = 0; index1 < ((IEnumerable<string>)strArray).Count<string>(); ++index1)
            {
                string index2 = strArray[index1];
                DataRowView dataItem = StartUp.oBrowse.ActiveRecord.DataItem as DataRowView;
                str = index1 != ((IEnumerable<string>)strArray).Count<string>() - 1 ? str + dataItem[index2].ToString() + ";" : str + string.Format("{0:yyyyMMdd}", (object)(DateTime)dataItem[index2]);
            }
            return str;
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
            if (StartUp.oBrowse.ActiveRecord != null)
                StartUp.currSqlTableKey = this.strKeyValue();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
            this.showWindow();
        }

        private void Copy()
        {
            if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.currActionTask = ActionTask.Copy;
            StartUp.currSqlTableKey = this.strKeyValue();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
            SysFunc.ChangeListID(StartupBase.SasObj, StartUp.sqlTableName, StartUp.oBrowse, StartUp.SqlTableKey);
        }

        private void oBrowse_F7(object sender, EventArgs e)
        {
            this.In();
        }

        private void In()
        {
            SqlCommand sqlcmd = new SqlCommand("select * from " + StartUp.sqlTableView + " ORDER BY ma_vt,ngay_nhap");
            DataSet LocalDataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            new ReportManager(StartupBase.SasObj, StartUp.CommandInfo["rep_file"].ToString()).Preview(LocalDataSet);
        }

        private void oBrowse_F8(object sender, EventArgs e)
        {
            this.Xoa();
        }

        private void Xoa()
        {
            if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.currSqlTableKey = this.strKeyValue();
            StartUp.currActionTask = ActionTask.Delete;
            if (SysFunc.CheckPermission(StartupBase.SasObj, StartUp.currActionTask, StartupBase.Menu_Id))
            {
                try
                {
                    if (ExMessageBox.Show(2435, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                        this.deleteRowByKey(StartUp.sqlTableName);
                }
                catch (SqlException ex)
                {
                    ErrorLog.CatchMessage(ex);
                }
            }
            else
            {
                int num = (int)ExMessageBox.Show(2445, StartupBase.SasObj, "Không có quyền xóa [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
          string Current_Ma_kh)
        {
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
                            StartUp.currSqlTableKey = Current_Ma_kh;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName_CatDau2 + " information";
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
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
                            StartUp.currSqlTableKey = Current_Ma_kh;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName_CatDau2 + " information";
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
                int num = (int)ExMessageBox.Show(2450, SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                StartUp.currActionTask = ActionTask.None;
            }
            return StartUp.LastEditTable;
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
                int num = (int)ExMessageBox.Show(2455, StartupBase.SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                StartUp.currActionTask = ActionTask.None;
            }
        }

        public static DataTable GetRow(string sqlTableName)
        {
            DataTable dataTable = (DataTable)null;
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
                    for (int index = 0; index < ((IEnumerable<string>)strArray2).Count<string>(); ++index)
                    {
                        if (index > 0)
                            str += " and";
                        DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + strArray2[index] + "'");
                        str = str + " " + strArray2[index] + "=@" + strArray2[index];
                        if (index == 0)
                            sqlcmd.Parameters.Add("@" + strArray2[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString())).Value = (object)strArray1[index];
                        else
                            sqlcmd.Parameters.Add("@" + strArray2[index], SqlDbType.VarChar).Value = (object)strArray1[index];
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
