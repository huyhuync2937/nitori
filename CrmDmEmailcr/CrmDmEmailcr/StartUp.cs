using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using SasUtilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CrmDmEmailcr
{
    public class StartUp : StartupBase
    {
        public static string sqlTableName = "dmemailcr";
        public static string sqlTableView = "v_dmemailcr";
        public static string SqlTableKey = "ma_emailcr";
        public static string SqlTableObjectName = "ten_emailcr";
        public static string TableName = "email";
        public static string TableName_CatDau = "emailcr";
        public static string TableName_CatDau2 = "emailcr";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;
        public static string titleWindow = string.Empty;
        public static string titleWindow2 = string.Empty;
        public static bool IsLoadFromLookUp = false;
        public static DataTable LastEditTable = (DataTable)null;
        public static DataRow CommandInfo;
        public static FrameBrowse oBrowse;
        public static DataRow drEmail;
        public override void Run()
        {
            StartupBase.Namespace = "CrmDmEmailcr";
            try
            {
                StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                if (StartupBase.M_LAN != "V")
                    StartUp.TableName = "Route";
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.oBrowse = new FrameBrowse(StartupBase.SasObj, StartUp.sqlTableName);
                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
                SysFunc.LoadIcon((Window)StartUp.oBrowse.frmBrw);
                StartUp.oBrowse.F2 += new FrameBrowse.GridKeyUp_F2(this.oBrowse_F2);
                StartUp.oBrowse.F3 += new FrameBrowse.GridKeyUp_F3(this.oBrowse_F3);
                StartUp.oBrowse.F4 += new FrameBrowse.GridKeyUp_F4(this.oBrowse_F4);
                StartUp.oBrowse.Ctrl_F4 += new FrameBrowse.GridKeyUp_Ctrl_F4(this.oBrowse_Ctrl_F4);
                StartUp.oBrowse.F6 += new FrameBrowse.GridKeyUp_F6(this.oBrowse_F6);
                StartUp.oBrowse.F7 += new FrameBrowse.GridKeyUp_F7(this.oBrowse_F7);
                StartUp.oBrowse.F8 += new FrameBrowse.GridKeyUp_F8(this.oBrowse_F8);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Moi += new ListToolBar.Execute(this.ToolBar_Cm_Moi);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Sua += new ListToolBar.Execute(this.ToolBar_Cm_Sua);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Xoa += new ListToolBar.Execute(this.ToolBar_Cm_Xoa);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_Copy += new ListToolBar.Execute(this.ToolBar_Cm_Copy);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_DoiMa += new ListToolBar.Execute(this.ToolBar_Cm_DoiMa);
                StartUp.oBrowse.frmBrw.ToolBar.Cm_In += new ListToolBar.Execute(this.ToolBar_Cm_In);
                StartUp.oBrowse.frmBrw.ShowInTaskbar = true;

                object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbList");
                if (name != null)
                {
                    ToolBar toolBar = name as ToolBar;
                    ToolBarButton toolBarButton = new ToolBarButton();
                    toolBarButton.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton.Name = "btnSenEmail";
                    toolBarButton.Text = "Gửi Email";
                    toolBarButton.ToolTip = "F12";
                    toolBarButton.ImagePath = "Images\\Feedback.png";
                    toolBarButton.Click += new RoutedEventHandler(this.btnSenEmail_Click);
                    toolBar.Items.Insert(11, toolBarButton);
                }
                DataTable dt = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select top 1 * from dmemail where user_id = {0}", Convert.ToInt32(StartupBase.SasObj.UserInfo.Rows[0]["user_id"])))).Tables[0];
                StartUp.drEmail = (dt == null || dt.Rows.Count == 0) ? null: dt.Rows[0];
                StartUp.oBrowse.frmBrw.LanguageID = "CrmDmEmailcr_brow";
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
        StartupBase.M_LAN.Equals("V") ? "ten_emailcr" : "ten_emailcr2"
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

        private void btnSenEmail_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
                return;

           if(StartUp.drEmail == null)
            {
                int num = (int)ExMessageBox.Show(6801, StartupBase.SasObj, "User chưa được cấu hình email!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString();
            string ma_kh_list = StartUp.oBrowse.ActiveRecord.Cells["ma_kh_list"].Value.ToString().Trim();
            string[] arr = ma_kh_list.Split(';');

            Info.fromAddress = drEmail["e_mail"].ToString().Trim();
            Info.fromPass = ClsEmail.DeCrypt(drEmail["pass"].ToString().Trim(), "0123456789").ToString().Trim();
            Info.fromName = StartupBase.SasObj.UserInfo.Rows[0]["comment"].ToString().Trim();
            Info.subject = StartUp.oBrowse.ActiveRecord.Cells["ghi_chu"].Value.ToString().Trim();
            string strquery = string.Format("SELECT STUFF((SELECT ',' + e_mail FROM dmkhcr where ISNULL(e_mail,'') <> '' AND ma_kh in (select items from dbo.SplitString('{0}', ';')) FOR XML PATH('')),1,1,'') AS e_mail",ma_kh_list.ToString().Trim());
            string email_list = StartupBase.SasObj.ExcuteScalar(new SqlCommand(strquery)).ToString();
            FrmSendEmaiCr frm = new FrmSendEmaiCr();
            frm.txtma_kh_list.Text = ma_kh_list;
            frm.txtemail_to_list.Text = email_list;
            frm.ShowDialog();
        }
    }
}
