using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CrmDmEmailcr
{
    public partial class newForm : FormList
    {
        private DataTable newDataTable = new DataTable();
        private DataTable OldRow = (DataTable)null;
        private EditModeBindingObject FormInEditMode;

        public newForm()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.FormInEditMode = (EditModeBindingObject)this.FindResource("IsInEditMode");
            if (StartUp.currActionTask != ActionTask.View)
                return;
            this.ConfirmGV.ButtonType = 1;
        }

        private void LoadForm()
        {
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName);
            this.txtma_emailcr.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_emailcr");
            this.txtten_emailcr.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_emailcr");
            this.txtten_emailcr2.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_emailcr2");
            this.txtghi_chu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ghi_chu");
            this.txtma_tra_cuu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_tra_cuu");
            this.txtma_kh_list.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_kh_list");
            this.Title = StartUp.titleWindow;
        }

        private void newForm_Loaded(object sender, RoutedEventArgs e)
        {
            this.LoadForm();
            TextBox child1 = SysFunc.FindChild<TextBox>((DependencyObject)this, "txt" + StartUp.SqlTableKey);
            if (child1 != null)
            {
                child1.SelectAll();
                child1.Focus();
            }
            else
                Debug.Write("Findchild not found");
            switch (StartUp.currActionTask)
            {
                case ActionTask.View:
                    try
                    {
                        StartUp.LastEditTable = (DataTable)null;
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        this.ConfirmGV.ButtonType = 1;
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        DataRow row = this.newDataTable.NewRow();
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_emailcr.MaxLength)
                                    row[StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(row[StartUp.SqlTableKey].ToString().Trim()))
                                row[StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                        }
                        row["status"] = "1";
                        this.newDataTable.Rows.Add(row);
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        if (this.newDataTable.Rows.Count > 0)
                            this.OldRow = this.newDataTable.Copy();
                        SqlCommand sqlcmd = new SqlCommand("exec dbo.CheckDeleteListId @ma_dm, @" + StartUp.SqlTableKey);
                        sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)StartUp.sqlTableName;
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)StartUp.currSqlTableKey;
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) <= 0)
                        {
                            if (child1 != null)
                                child1.IsReadOnly = true;
                            TextBox child2 = SysFunc.FindChild<TextBox>((DependencyObject)this, "txt" + StartUp.SqlTableObjectName);
                            if (child2 != null)
                            {
                                child2.SelectAll();
                                child2.Focus();
                            }
                            break;
                        }
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        this.newDataTable.Rows[0]["status"] = "1";
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_emailcr.MaxLength)
                                    this.newDataTable.Rows[0][StartUp.SqlTableKey] = str;
                            }
                            if (string.IsNullOrEmpty(this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim()))
                                this.newDataTable.Rows[0][StartUp.SqlTableKey] = SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                            break;
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                        break;
                    }
            }
            this.gridMain.DataContext = (object)this.newDataTable;
        }

        private void saveCustomer()
        {
            StartUp.LastEditTable = (DataTable)null;
            if (StartUp.currActionTask == ActionTask.View)
                return;
            this.newDataTable.AcceptChanges();
            int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            string str = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
            if (this.newDataTable.Columns.Contains("search"))
                SysFunc.SetStrSearch(StartupBase.SasObj, StartUp.sqlTableName, ref this.newDataTable);
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                this.newDataTable.Rows[0]["date"] = DateTime.Now;
                this.newDataTable.Rows[0]["time"] = DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = num;
                this.newDataTable.Rows[0]["user_name"] = str;
                if (this.OldRow != null)
                    ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            else
            {
                this.newDataTable.Rows[0]["date"] = DateTime.Now;
                this.newDataTable.Rows[0]["time"] = DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = num;
                this.newDataTable.Rows[0]["user_name"] = str;
                this.newDataTable.Rows[0]["date0"] = DateTime.Now;
                this.newDataTable.Rows[0]["time0"] = DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id0"] = num;
                this.newDataTable.Rows[0]["user_name0"] = str;
                ListFunc.inserRowInDataBase(StartUp.sqlTableName, this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            StartUp.LastEditTable = this.newDataTable;
        }

        private bool CheckValid()
        {
            bool flag1 = true;
            if (this.txtma_emailcr.Text.Trim() == string.Empty && flag1)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(15, StartupBase.SasObj, "Chưa vào mã [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_emailcr.Focus();
                flag1 = false;
            }
            if (this.txtma_emailcr.Text.Trim() != "" && flag1)
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtma_emailcr.Text.Trim());
                if (str != "" && flag1)
                {
                    this.TabInfor.SelectedIndex = 0;
                    int num = (int)ExMessageBox.Show(20, StartupBase.SasObj, "Mã không được chứa các ký tự [" + str + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtma_emailcr.SelectAll();
                    this.txtma_emailcr.Focus();
                    flag1 = false;
                }
                if (flag1)
                {
                    try
                    {
                        bool flag2 = true;
                        if (StartUp.currActionTask == ActionTask.Edit && flag1)
                        {
                            flag2 = false;
                            if (this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim() != this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim())
                                flag2 = true;
                        }
                        SqlCommand sqlcmd = new SqlCommand("exec dbo.CheckExistListId @ma_dm, @" + StartUp.SqlTableKey);
                        sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)StartUp.sqlTableName;
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)this.txtma_emailcr.Text.Trim();
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num = (int)ExMessageBox.Show(25, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_emailcr.SelectAll();
                            this.txtma_emailcr.Focus();
                            flag1 = false;
                        }
                        if (flag1 && flag2)
                        {
                            string Value_old = "";
                            if (StartUp.currActionTask == ActionTask.Edit)
                                Value_old = this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            if (SysFunc.CheckStringContain(StartupBase.SasObj, StartUp.sqlTableName, StartUp.SqlTableKey, this.txtma_emailcr.Text.Trim(), Value_old))
                            {
                                this.TabInfor.SelectedIndex = 0;
                                int num = (int)ExMessageBox.Show(30, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_emailcr.SelectAll();
                                this.txtma_emailcr.Focus();
                                flag1 = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                }
            }
            if (this.txtten_emailcr.Text.Trim() == string.Empty && flag1)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(35, StartupBase.SasObj, "Chưa vào tên [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtten_emailcr.Focus();
                flag1 = false;
            }
            return flag1;
        }

        public void EnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask == ActionTask.View)
            {
                this.Close();
            }
            else
            {
                if (!this.CheckValid())
                    return;
                this.saveCustomer();
                StartUp.currActionTask = ActionTask.None;
                this.Close();
            }
        }

        private void txtma_emailcr_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_emailcr.Text = this.txtma_emailcr.Text.Trim();
            this.txtma_emailcr.SelectAll();
        }

        private void txtma_tra_cuu_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_tra_cuu.Text = this.txtma_tra_cuu.Text.Trim();
            this.txtma_tra_cuu.SelectAll();
        }

        private void FormList_Closed(object sender, EventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None)
                SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.currActionTask = ActionTask.None;
        }

        private void ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
        {
            SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.LastEditTable = (DataTable)null;
            this.Close();
        }

        private void FormList_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Escape)
                return;
            StartUp.LastEditTable = (DataTable)null;
            this.Close();
        }
        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtstatus.Text == ""))
                return;
            this.newDataTable.Rows[0]["status"] = 0;
        }

        private void btnma_kh_list_Click(object sender, RoutedEventArgs e)
        {
            DanhSachNhanVien danhSachNhanvien = new DanhSachNhanVien();
            danhSachNhanvien.strma_ns = this.txtma_kh_list.Text.Trim();

            bool? nullable = danhSachNhanvien.ShowDialog();
            if ((!nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                return;

            this.newDataTable.Rows[0].SetField<string>("ma_kh_list", danhSachNhanvien.strma_ns);
        }
    }
}
