using Infragistics.Windows.DataPresenter;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Intc
{
     public class StartUp : StartupBase
      {
        public static string sqlTableUpdateName = "dmtieuchi";
        public static string sqlTableName = "dmtieuchi";
        public static string sqlTableView = "v_dmtieuchi";
        public static string SqlTableKey = "ma_tieu_chi";
        public static string SqlTableObjectName = "ten_tieu_chi";
        public static string TableName = "tiêu chí";
        public static string TableName_CatDau = "tiêu chí";
        public static string TableName_CatDau2 = "criteria";
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKey = string.Empty;
        public static string titleWindow = string.Empty;
        public static bool IsLoadFromLookUp = false;
        public static DataTable LastEditTable = (DataTable) null;
        public static string Parameter = "";
        public static DataRow CommandInfo;
        public static FrameBrowse oBrowse;

        public override void Run()
        {
          StartupBase.Namespace = "Intc";
          try
          {
            StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
            StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
            StartUp.Parameter = StartUp.CommandInfo["parameter"].ToString().Trim();
            StartUp.oBrowse = new FrameBrowse(StartupBase.SasObj, StartUp.sqlTableName);
            StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
            StartUp.TableName = StartupBase.M_LAN.Equals("V") ? "khách hàng" : "customer";
            SysFunc.LoadIcon((Window) StartUp.oBrowse.frmBrw);
            StartUp.oBrowse.F2 += new FrameBrowse.GridKeyUp_F2(this.oBrowse_F2);
            StartUp.oBrowse.F3 += new FrameBrowse.GridKeyUp_F3(this.oBrowse_F3);
            StartUp.oBrowse.F4 += new FrameBrowse.GridKeyUp_F4(this.oBrowse_F4);
            StartUp.oBrowse.Ctrl_F4 += new FrameBrowse.GridKeyUp_Ctrl_F4(this.oBrowse_Ctrl_F4);
            StartUp.oBrowse.F6 += new FrameBrowse.GridKeyUp_F6(this.oBrowse_F6);
            StartUp.oBrowse.F8 += new FrameBrowse.GridKeyUp_F8(this.oBrowse_F8);
            StartUp.oBrowse.frmBrw.ToolBar.Cm_Moi += new ListToolBar.Execute(this.ToolBar_Cm_Moi);
            StartUp.oBrowse.frmBrw.ToolBar.Cm_Sua += new ListToolBar.Execute(this.ToolBar_Cm_Sua);
            StartUp.oBrowse.frmBrw.ToolBar.Cm_Xoa += new ListToolBar.Execute(this.ToolBar_Cm_Xoa);
            StartUp.oBrowse.frmBrw.ToolBar.Cm_Copy += new ListToolBar.Execute(this.ToolBar_Cm_Copy);
            StartUp.oBrowse.frmBrw.ToolBar.Cm_DoiMa += new ListToolBar.Execute(this.ToolBar_Cm_DoiMa);
            StartUp.oBrowse.frmBrw.ShowInTaskbar = true;
            object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbList");
            if (name != null)
            {
              ToolBar toolBar = name as ToolBar;
              new ToolBarButton()
              {
                ImagePath = "Images\\Preview.png",
                Text = (StartupBase.M_LAN.Equals("V") ? "Sắp xếp" : "Sort")
              }.Click += new RoutedEventHandler(this.btSapXep_Click);
            }
            StartUp.oBrowse.frmBrw.LanguageID = "Arkh_1";
            StartUp.oBrowse.ShowDialog();
          }
          catch (Exception ex)
          {
            ErrorLog.CatchMessage(ex);
          }
        }

        private void btSapXep_Click(object sender, RoutedEventArgs e)
        {
          this.SapXep((object) null, (KeyEventArgs) null);
        }

        private void oBrowse_F11(object sender, EventArgs e)
        {
          this.SapXep((object) null, (KeyEventArgs) null);
        }

        private void SapXep(object sender, KeyEventArgs e)
        {
          ArkhF10 arkhF10 = new ArkhF10();
          SysFunc.LoadIcon((Window) arkhF10);
          arkhF10.Title = SysFunc.Cat_Dau(arkhF10.Title);
          if (!arkhF10.ShowDialog())
            return;
          XamDataGrid dataGrid = this.FindDataGrid((DependencyObject) StartUp.oBrowse.frmBrw);
          if (dataGrid == null)
            return;
          dataGrid.FieldLayouts[0].SortedFields.Clear();
          switch (arkhF10.txtSapXep.Text.Trim())
          {
            case "1":
              dataGrid.FieldLayouts[0].SortedFields.Add(new FieldSortDescription("ma_kh", ListSortDirection.Ascending, false));
              break;
            case "2":
              dataGrid.FieldLayouts[0].SortedFields.Add(new FieldSortDescription("ten_kh", ListSortDirection.Ascending, false));
              break;
            case "3":
              dataGrid.FieldLayouts[0].SortedFields.Add(new FieldSortDescription("nh_kh1", ListSortDirection.Ascending, false));
              break;
            case "4":
              dataGrid.FieldLayouts[0].SortedFields.Add(new FieldSortDescription("nh_kh2", ListSortDirection.Ascending, false));
              break;
            case "5":
              dataGrid.FieldLayouts[0].SortedFields.Add(new FieldSortDescription("nh_kh3", ListSortDirection.Ascending, false));
              break;
          }
        }

        private XamDataGrid FindDataGrid(DependencyObject parent)
        {
          if (parent == null)
            return (XamDataGrid) null;
          if (parent is XamDataGrid)
            return parent as XamDataGrid;
          for (int childIndex = 0; childIndex < VisualTreeHelper.GetChildrenCount(parent); ++childIndex)
          {
            DependencyObject dataGrid = (DependencyObject) this.FindDataGrid(VisualTreeHelper.GetChild(parent, childIndex));
            if (dataGrid != null && dataGrid is XamDataGrid)
              return dataGrid as XamDataGrid;
          }
          return (XamDataGrid) null;
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
          StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Xem thông tin " + StartUp.TableName_CatDau : "View " + StartUp.TableName_CatDau2 + " information";
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
          StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin " + StartUp.TableName_CatDau : "Edit " + StartUp.TableName_CatDau2 + " information";
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
          StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add new customer";
          this.showWindow();
        }

        private void Copy()
        {
          if (StartUp.currActionTask != ActionTask.None || StartUp.oBrowse.ActiveRecord == null)
            return;
          StartUp.currActionTask = ActionTask.Copy;
          StartUp.currSqlTableKey = StartUp.oBrowse.ActiveRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
          StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add accumulative amount";
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
            StartupBase.M_LAN.Equals("V") ? "Đổi mã tiêu chí" : "Change criteria code",
            StartupBase.M_LAN.Equals("V") ? "ten_tieu_chi" : "ten_tieu_chi2"
          };
          DataRecord activeRecord = StartUp.oBrowse.ActiveRecord;
          activeRecord.Cells[StartUp.SqlTableKey].Value = (object) activeRecord.Cells[StartUp.SqlTableKey].Value.ToString().Trim();
          SysFunc.ChangeListID(StartupBase.SasObj, StartUp.sqlTableUpdateName, StartUp.oBrowse, StartUp.SqlTableKey, TitleAndName);
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
              sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object) StartUp.sqlTableUpdateName;
              sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object) StartUp.currSqlTableKey;
              if ((int) StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0)
              {
                if (ExMessageBox.Show(1810, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                  this.deleteRowByKey(StartUp.sqlTableUpdateName);
              }
              else
              {
                int num = (int) ExMessageBox.Show(1815, StartupBase.SasObj, "Đã có phát sinh không được xóa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
              }
            }
            catch (SqlException ex)
            {
              ErrorLog.CatchMessage(ex);
            }
          }
          else
          {
            int num1 = (int) ExMessageBox.Show(1820, StartupBase.SasObj, "Không có quyền xóa [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                  StartUp.currSqlTableKey = Current_Ma_kh;
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
              case ActionTask.Copy:
                try
                {
                  StartupBase.SasObj = SasObj;
                  StartUp.currSqlTableKey = Current_Ma_kh;
                  StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Thêm " + StartUp.TableName : "Add " + StartUp.TableName_CatDau2;
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
            int num = (int) ExMessageBox.Show(1825, SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
            int num = (int) ExMessageBox.Show(1830, StartupBase.SasObj, "Không có quyền [" + StartUp.titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
              for (int index = 0; index < strArray2.Length; ++index)
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

        public static string CatChuoi(string lstString, string M_LAN)
        {
          string str = string.Empty;
          string[] separator = new string[1]{ "*|" };
          if (string.IsNullOrEmpty(lstString.Trim()))
            str = !M_LAN.ToUpper().Trim().Equals("V") ? lstString.Split(separator, StringSplitOptions.None)[1] : lstString.Split(separator, StringSplitOptions.None)[0];
          return str;
        }
      }
}
