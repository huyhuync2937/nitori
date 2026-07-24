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
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace SasIeCt0
{
    public class StartUp : StartupBase
    {
        public static string sqlTableName = "dmbp";
        public static string sqlTableView = "v_dmbp";
        public static string SqlTableKey = "ma_bp";
        public static string SqlTableObjectName = "ten_bp";
        public static string TableName = "nhân viên";
        public static string TableName_CatDau = "nhan vien";
        public static string TableName_CatDau2 = "staff";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;
        public static string titleWindow = string.Empty;
        public static string titleWindow2 = string.Empty;
        public static bool IsLoadFromLookUp = false;
        public static DataTable LastEditTable = (DataTable)null;
        public static DataRow CommandInfo;
        public static FrameBrowse oBrowse;
        public static string M_IP_TIEN;
        public static string M_IP_TIEN_NT;

        public override void Run()
        {
            StartupBase.Namespace = "SasIeCt0";
            try
            {
                StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                StartupBase.SasObj.SynchroFile(".", "SasIeCt.exe");
                Process.Start("SasIeCt.exe", string.Format("\"{0}\"", StartupBase.Menu_Id));
                if (!(Process.GetCurrentProcess().ProcessName != "SasProcess.exe"))
                    return;
                Application.Current.Shutdown();
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
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thong tin " + StartUp.TableName_CatDau : "View " + StartUp.TableName_CatDau2 + " information";
            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName_CatDau2 + " information";
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
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sua thong tin " + StartUp.TableName_CatDau : "Edit " + StartUp.TableName_CatDau2 + " information";
            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName_CatDau2 + " information";
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
            if (StartUp.oBrowse.ActiveRecord != null)
                StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them " + StartUp.TableName_CatDau : "Add " + StartUp.TableName_CatDau2;
            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
            this.showWindow();
        }

        private void Copy()
        {
            if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.currActionTask = ActionTask.Copy;
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them " + StartUp.TableName_CatDau : "Add " + StartUp.TableName_CatDau2;
            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
            string[] TitleAndName = new string[2]
            {
        StartupBase.M_LAN.Equals("V") ? "Doi ma " + StartUp.TableName : "Change " + StartUp.TableName_CatDau2 + " code",
        StartupBase.M_LAN.Equals("V") ? "ten_bp" : "ten_bp2"
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
                        if (ExMessageBox.Show(40, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                            this.deleteRowByKey(StartUp.sqlTableName);
                    }
                    else
                    {
                        int num = (int)ExMessageBox.Show(45, StartupBase.SasObj, "Đã có phát sinh không được xóa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                }
                catch (SqlException ex)
                {
                    ErrorLog.CatchMessage(ex);
                }
            }
            else
            {
                int num1 = (int)ExMessageBox.Show(50, StartupBase.SasObj, "Không có quyền xóa [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thong tin " + StartUp.TableName_CatDau : "View " + StartUp.TableName_CatDau2 + " information";
                            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName : "View " + StartUp.TableName_CatDau2 + " information";
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
                            StartUp.currSqlTableKey = Current_Ma_kh;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them " + StartUp.TableName_CatDau : "Add " + StartUp.TableName_CatDau2;
                            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sua thong tin " + StartUp.TableName_CatDau : "Edit " + StartUp.TableName_CatDau2 + " information";
                            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName : "Edit " + StartUp.TableName_CatDau2 + " information";
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
                            StartUp.currSqlTableKey = Current_Ma_kh;
                            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them " + StartUp.TableName_CatDau : "Add " + StartUp.TableName_CatDau2;
                            StartUp.titleWindow2 = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
                int num = (int)ExMessageBox.Show(55, SasObj, "Không có quyền [" + StartUp.titleWindow2.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                int num = (int)ExMessageBox.Show(60, StartupBase.SasObj, "Không có quyền [" + StartUp.titleWindow2.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                        sqlcmd.Parameters.Add("@" + strArray2[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString())).Value = (object)strArray1[index];
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
