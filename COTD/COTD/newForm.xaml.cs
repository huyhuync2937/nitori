using Infragistics.Windows.Controls;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace COTD
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
            this.FormInEditMode = (EditModeBindingObject)this.FindResource((object)"IsInEditMode");
            if (StartUp.currActionTask != ActionTask.View)
                return;
            this.ConfirmGV.ButtonType = 1;
        }

        private void LoadForm()
        {
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1));
            StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
            StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
            this.txtma_td.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_td");
            this.txtten_td.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_td");
            this.txtten_td2.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_td2");
            this.txtghi_chu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ghi_chu");
            this.Title = SysFunc.Cat_Dau(StartUp.titleWindow) + " " + (StartUp._parameter.Trim().Equals("1") ? "" : StartUp._parameter.Trim());
            if (StartUp._parameter.Equals("5"))
            {
                this.txtnh_td1.ListID = "dmnhtd"  + "11";
                this.txtnh_td2.ListID = "dmnhtd" + "12";
                this.txtnh_td3.ListID = "dmnhtd" + "13";
                //this.txtnh_td3.ListID = "dmnhtd" + StartUp._parameter + "5";
            }
            else
            {
                this.txtnh_td1.ListID = "dmnhtd" + StartUp._parameter + "1";
                this.txtnh_td2.ListID = "dmnhtd" + StartUp._parameter + "2";
                this.txtnh_td3.ListID = "dmnhtd" + StartUp._parameter + "3";
            }

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
                        row["tien_nt"] = (object)0;
                        row["tien"] = (object)0;
                        row["du_kh_yn"] = (object)0;
                        row["status"] = (object)1;
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1));
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_td.MaxLength)
                                    row[StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(row[StartUp.SqlTableKey].ToString().Trim()))
                                row[StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                        }
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
                        {
                            this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            this.OldRow = this.newDataTable.Copy();
                        }
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
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1));
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_td.MaxLength)
                                    this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim()))
                                this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
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
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               this.txtnh_td1.SearchInit();
               this.txtnh_td1_PreviewLostFocus((object)this.txtnh_td1, (KeyboardFocusChangedEventArgs)null);
               this.txtnh_td2.SearchInit();
               this.txtnh_td2_PreviewLostFocus((object)this.txtnh_td2, (KeyboardFocusChangedEventArgs)null);
               this.txtnh_td3.SearchInit();
               this.txtnh_td3_PreviewLostFocus((object)this.txtnh_td3, (KeyboardFocusChangedEventArgs)null);
           }), DispatcherPriority.Background);
        }

        private void saveCustomer()
        {
            StartUp.LastEditTable = (DataTable)null;
            if (StartUp.currActionTask == ActionTask.View)
                return;
            this.newDataTable.AcceptChanges();
            int int16 = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"]);
            string str = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
            if (this.newDataTable.Columns.Contains("search"))
                SysFunc.SetStrSearch(StartupBase.SasObj, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1), ref this.newDataTable);
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)int16;
                this.newDataTable.Rows[0]["user_name"] = (object)str;
                if (this.OldRow != null)
                    ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1), StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            else
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)int16;
                this.newDataTable.Rows[0]["user_name"] = (object)str;
                this.newDataTable.Rows[0]["date0"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time0"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id0"] = (object)int16;
                this.newDataTable.Rows[0]["user_name0"] = (object)str;
                ListFunc.inserRowInDataBase(StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1), this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            StartUp.LastEditTable = this.newDataTable;
        }

        private bool CheckValid()
        {
            bool flag1 = true;
            if (flag1 && this.txtma_td.Text.Trim() == "")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(3595, StartupBase.SasObj, "Chưa vào mã [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_td.Focus();
                flag1 = false;
            }
            if (flag1 && this.txtma_td.Text.Trim() != "")
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtma_td.Text.Trim());
                if (str != "" && flag1)
                {
                    this.TabInfor.SelectedIndex = 0;
                    int num = (int)ExMessageBox.Show(3600, StartupBase.SasObj, "Mã không được chứa các ký tự [" + str + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtma_td.SelectAll();
                    this.txtma_td.Focus();
                    flag1 = false;
                }
                if (flag1)
                {
                    try
                    {
                        bool flag2 = true;
                        if (flag1 && StartUp.currActionTask == ActionTask.Edit)
                        {
                            flag2 = false;
                            if (this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim() != this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim())
                                flag2 = true;
                        }
                        SqlCommand sqlcmd = new SqlCommand("exec dbo.CheckExistListId @ma_dm, @" + StartUp.SqlTableKey);
                        sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)StartUp.sqlTableName;
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)this.txtma_td.Text.Trim();
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num = (int)ExMessageBox.Show(3605, StartupBase.SasObj, "Mã đã có hoặc lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_td.SelectAll();
                            this.txtma_td.Focus();
                            flag1 = false;
                        }
                        if (flag1 && flag2)
                        {
                            string Value_old = "";
                            if (StartUp.currActionTask == ActionTask.Edit)
                                Value_old = this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            if (SysFunc.CheckStringContain(StartupBase.SasObj, StartUp.sqlTableName.Substring(0, StartUp.sqlTableName.Length - 1), StartUp.SqlTableKey, this.txtma_td.Text.Trim(), Value_old))
                            {
                                this.TabInfor.SelectedIndex = 0;
                                int num = (int)ExMessageBox.Show(3610, StartupBase.SasObj, "Mã đã có hoặc lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_td.SelectAll();
                                this.txtma_td.Focus();
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
            if (flag1 && this.txtten_td.Text.Trim() == "")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(3615, StartupBase.SasObj, "Chưa vào tên [" + StartUp.TableName + "]!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtten_td.Focus();
                flag1 = false;
            }
            if (flag1 && !this.txtnh_td1.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(3660, StartupBase.SasObj, "Phân nhóm 1 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag1 = false;
                this.txtnh_td1.IsFocus = true;
            }
            if (flag1 && !this.txtnh_td2.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(3665, StartupBase.SasObj, "Phân nhóm 2 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag1 = false;
                this.txtnh_td2.IsFocus = true;
            }
            if (flag1 && !this.txtnh_td3.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(3670, StartupBase.SasObj, "Phân nhóm 3 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag1 = false;
                this.txtnh_td3.IsFocus = true;
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

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.Return) || Keyboard.Modifiers != ModifierKeys.Alt)
                return;
            TextBox textBox = sender as TextBox;
            textBox.SelectedText = Environment.NewLine;
            ++textBox.SelectionStart;
            textBox.SelectionLength = 0;
        }

        protected override bool IsEnterToPassObject(object sender)
        {
            return (!(sender is TextBox) || !((sender as TextBox).Name == "txtghi_chu") || Keyboard.Modifiers != ModifierKeys.Alt) && base.IsEnterToPassObject(sender);
        }

        private void txtma_td_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_td.Text = this.txtma_td.Text.Trim();
            this.txtma_td.SelectAll();
        }

        private void txtnh_td1_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtnh_td1.RowResult == null)
                this.txtten_nh_td1.Text = "";
            else
                this.txtten_nh_td1.Text = StartupBase.M_LAN == "V" ? this.txtnh_td1.RowResult["ten_nh"].ToString() : this.txtnh_td1.RowResult["ten_nh2"].ToString();
        }

        private void txtnh_td3_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtnh_td3.RowResult == null)
                this.txtten_nh_td3.Text = "";
            else
                this.txtten_nh_td3.Text = StartupBase.M_LAN == "V" ? this.txtnh_td3.RowResult["ten_nh"].ToString() : this.txtnh_td3.RowResult["ten_nh2"].ToString();
        }

        private void txtnh_td2_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtnh_td2.RowResult == null)
                this.txtten_nh_td2.Text = "";
            else
                this.txtten_nh_td2.Text = StartupBase.M_LAN == "V" ? this.txtnh_td2.RowResult["ten_nh"].ToString() : this.txtnh_td2.RowResult["ten_nh2"].ToString();
        }

        private Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtstatus.Text == ""))
                return;
            this.newDataTable.Rows[0]["status"] = (object)0;
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

        private void txtma_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
        }
    }
}
