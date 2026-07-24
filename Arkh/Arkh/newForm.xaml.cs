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
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Automation;
using mshtml;
using System.Windows.Interop;
using System.Runtime.InteropServices;

namespace Arkh
{
    public partial class newForm : FormList
    {
        private DataTable newDataTable = new DataTable();
        private DataTable OldRow = (DataTable)null;
        private EditModeBindingObject FormInEditMode;
        System.Windows.Threading.DispatcherTimer Starttimer;
        string url = "http://help.sis.vn:7785/AP/getmasothue";
        //string mst = "0101218690";
        void khoitaostarttimer()
        {
            Starttimer = new System.Windows.Threading.DispatcherTimer();
            Starttimer.Interval = new TimeSpan(0, 0, 1);
            Starttimer.Tick += starttimer_Tick;
        }

        public newForm()
        {
            this.InitializeComponent();
            khoitaostarttimer();
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
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableView);
            this.txtma_kh.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_kh");
            this.txtten_kh.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_kh");
            this.txtten_kh2.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_kh2");
            this.txtten_kh3.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_kh3");
            this.txtma_tra_cuu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_tra_cuu");
            this.txtdia_chi.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "dia_chi");
            this.txtdoi_tac.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "doi_tac");
            this.txtma_so_thue.MaxLength = 14;
            this.txtdien_thoai.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "dien_thoai");
            this.txtfax.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "fax");
            this.txtemail.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "e_mail");
            this.txttk_nh.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "tk_nh");
            this.txtten_nh.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_nh");
            this.txttinh_thanh.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "tinh_thanh");
            this.txtghi_chu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ghi_chu");
            this.txtso_the.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "so_the");
            this.Title = SysFunc.Cat_Dau(StartUp.titleWindow);

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
                        int length = StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length;
                        DataTable dmdmInfo = StartupBase.SasObj.DmdmInfo;
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        DataRow row = this.newDataTable.NewRow();
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableUpdateName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_kh.MaxLength)
                                    row[StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(row[StartUp.SqlTableKey].ToString().Trim()))
                                row[StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                        }
                        row["status"] = (object)"1";
                        if (StartUp.Parameter == "1")
                            row["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs.ToString();
                        this.newDataTable.Rows.Add(row);
                        this.txtma_kh.Focus();
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
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableUpdateName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_kh.MaxLength)
                                    this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim()))
                                this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                        }
                        this.newDataTable.Rows[0]["status"] = (object)"1";
                        this.newDataTable.Rows[0]["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs.ToString();
                        break;
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                        break;
                    }
            }
            this.gridMain.DataContext = this.gridMainB.DataContext = (object)this.newDataTable;
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
                   {
                       this.txtMa_dvcs.SearchInit();
                       this.txtMa_dvcs_LostFocus((object)this.txtMa_dvcs, (RoutedEventArgs)null);
                   }), DispatcherPriority.Background);
        }

        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtstatus.Text == ""))
                return;
            this.newDataTable.Rows[0]["status"] = (object)0;
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
                SysFunc.SetStrSearch(StartupBase.SasObj, "dmkh", ref this.newDataTable);
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
                this.newDataTable.Rows[0]["user_name"] = (object)str;
                if (this.OldRow != null)
                    ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableUpdateName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            else
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
                this.newDataTable.Rows[0]["user_name"] = (object)str;
                this.newDataTable.Rows[0]["date0"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time0"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id0"] = (object)num;
                this.newDataTable.Rows[0]["user_name0"] = (object)str;
                ListFunc.inserRowInDataBase(StartUp.sqlTableUpdateName, this.newDataTable.Rows[0], StartupBase.SasObj);
            }
           
            StartUp.LastEditTable = this.newDataTable;
        }
        public static string GetConnectionString(int _time)
        {
            string text = "";
            new DataSet();
            try
            {
                text = StartupBase.SasObj.M_ConnectString;
                if (!text.ToUpper().Contains("TIMEOUT"))
                {
                    return text + ";Connect Timeout = " + _time;
                }
                return text;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
                return "";
            }
        }

        private bool CheckValid()
        {
            bool flag1 = true;
            if (this.txtma_kh.Text.Trim() == string.Empty && flag1)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1745, StartupBase.SasObj, "Chưa vào mã [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_kh.Focus();
                flag1 = false;
            }
            if (this.txtma_kh.Text.Trim() != "" && flag1)
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtma_kh.Text.Trim());
                if (str != "" && flag1)
                {
                    this.TabInfor.SelectedIndex = 0;
                    int num = (int)ExMessageBox.Show(1750, StartupBase.SasObj, "Mã không được chứa các ký tự [" + str + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtma_kh.SelectAll();
                    this.txtma_kh.Focus();
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
                        sqlcmd.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)StartUp.sqlTableUpdateName;
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)this.txtma_kh.Text.Trim();
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num = (int)ExMessageBox.Show(1755, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_kh.SelectAll();
                            this.txtma_kh.Focus();
                            flag1 = false;
                        }
                        if (flag1 && flag2)
                        {
                            string Value_old = "";
                            if (StartUp.currActionTask == ActionTask.Edit)
                                Value_old = this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            if (SysFunc.CheckStringContain(StartupBase.SasObj, StartUp.sqlTableUpdateName, StartUp.SqlTableKey, this.txtma_kh.Text.Trim(), Value_old))
                            {
                                this.TabInfor.SelectedIndex = 0;
                                int num = (int)ExMessageBox.Show(1760, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_kh.SelectAll();
                                this.txtma_kh.Focus();
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
            if (flag1 && this.txtten_kh.Text.Trim() == "")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1765, StartupBase.SasObj, "Chưa vào tên [" + StartUp.TableName + "]!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtten_kh.Focus();
                flag1 = false;
            }
            if (this.txtma_so_thue.Text.Trim() != "" && flag1)
            {
                string str = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                if (!str.Equals("0"))
                {
                    if (!SysFunc.CheckSumMaSoThue(this.txtma_so_thue.Text.Trim()))
                    {
                        if (str.Equals("1"))
                        {
                            int num1 = (int)ExMessageBox.Show(1770, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                        else
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num2 = (int)ExMessageBox.Show(1775, StartupBase.SasObj, "Mã số thuế không hợp lệ, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_so_thue.SelectAll();
                            this.txtma_so_thue.Focus();
                            flag1 = false;
                        }
                    }
                    if (flag1)
                    {
                        SqlCommand sqlcmd = new SqlCommand("select * from " + StartUp.sqlTableUpdateName + " where ma_so_thue = @ma_so_thue");
                        sqlcmd.Parameters.Add("@ma_so_thue", SqlDbType.Char).Value = (object)this.txtma_so_thue.Text.Trim();
                        DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                        if (StartUp.currActionTask != ActionTask.Edit && table.Rows.Count > 0 || StartUp.currActionTask == ActionTask.Edit && table.Rows.Count > 1)
                        {
                            if (str.Equals("1"))
                            {
                                int num2 = (int)ExMessageBox.Show(1780, StartupBase.SasObj, "Mã số thuế này đã có!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else
                            {
                                this.TabInfor.SelectedIndex = 0;
                                int num3 = (int)ExMessageBox.Show(1785, StartupBase.SasObj, "Mã số thuế này đã có, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_so_thue.SelectAll();
                                this.txtma_so_thue.Focus();
                                flag1 = false;
                            }
                        }
                    }
                }
            }
            if (flag1 && !this.txttk.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1790, StartupBase.SasObj, "Tài khoản ngầm định không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txtnh_kh1.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1795, StartupBase.SasObj, "Nhóm khách 1 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtnh_kh1.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txtnh_kh2.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1800, StartupBase.SasObj, "Nhóm khách 2 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtnh_kh2.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txtnh_kh3.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(1805, StartupBase.SasObj, "Nhóm khách 3 không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtnh_kh3.IsFocus = true;
                flag1 = false;
            }
            SqlConnection sqlConnection = new SqlConnection(GetConnectionString(500000));
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandTimeout = 600000;
            sqlCommand.Connection = sqlConnection;
            sqlConnection.Open();
            string tk_portal = this.newDataTable.Rows[0]["tk_portal"].ToString();
            string ma_kh = this.newDataTable.Rows[0]["ma_kh"].ToString();

            if (!string.IsNullOrEmpty(tk_portal))
            {
                sqlCommand.CommandText = string.Format("EXEC {0} '{1}', '{2}'", "CreateUser", tk_portal, ma_kh);
                int result = Convert.ToInt32(sqlCommand.ExecuteScalar());

                if (result == 1)
                {
                    // thành công
                }
                else
                {
                    int num = (int)ExMessageBox.Show(2000, StartupBase.SasObj, "Đã tồn tại tk portal!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txttk_portal.Focus();
                    flag1 = false;
                }
            }
            return flag1;
        }

        private void txttk_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txttk.RowResult == null)
                this.txtten_tk.Text = "";
            else
                this.txtten_tk.Text = StartupBase.M_LAN.Equals("V") ? this.txttk.RowResult["ten_tk"].ToString() : this.txttk.RowResult["ten_tk2"].ToString();
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

        private void txtma_kh_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_kh.Text = this.txtma_kh.Text.Trim();
            this.txtma_kh.SelectAll();
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

        private void txtHan_ck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtHan_ck.RowResult == null)
            {
                this.newDataTable.Rows[0]["han_tt"] = (object)0;
                //this.txtthck.Text = "";
            }
            else
            {
                this.newDataTable.Rows[0]["han_tt"] = string.IsNullOrEmpty(this.txtHan_ck.RowResult["han_tt"].ToString()) ? (object)0 : this.txtHan_ck.RowResult["han_tt"];
                //this.txtthck.Text = StartupBase.M_LAN.Equals("V") ? this.txtHan_ck.RowResult["ten_thck"].ToString() : this.txtHan_ck.RowResult["ten_thck2"].ToString();
            }
        }

        private void txtMa_dvcs_LostFocus(object sender, RoutedEventArgs e)
        {
            AutoCompleteTextBox autoCompleteTextBox = sender as AutoCompleteTextBox;
            if (autoCompleteTextBox.RowResult == null)
            {
                this.txtTen_dvcs.Text = "";
            }
            else
            {
                try
                {
                    this.txtTen_dvcs.Text = !(StartupBase.M_LAN == "V") ? autoCompleteTextBox.RowResult["ten_dvcs2"].ToString() : autoCompleteTextBox.RowResult["ten_dvcs"].ToString();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            }
        }
        private void txtso_the_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtso_the.Text = this.txtso_the.Text.Trim();
            this.txtso_the.SelectAll();
        }
        HTMLDocument document;
        void starttimer_Tick(object sender, EventArgs e)
        {
            Starttimer.Stop();
            var result = document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList").innerHTML;
            if (document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd2") != null)
            {
                this.newDataTable.Rows[0]["ma_so_thue"] = document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd2").innerText;
            }
            else
            {
                this.newDataTable.Rows[0]["ma_so_thue"] = "";
            }
            if (document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd3") != null)
            {
                this.newDataTable.Rows[0]["ten_kh"] = document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd3").innerText;
                this.newDataTable.Rows[0]["ten_kh2"] = document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd3").innerText;
            }
            else
            {
                this.newDataTable.Rows[0]["ten_kh"] = "";
                this.newDataTable.Rows[0]["ten_kh2"] = "";
            }
            if (document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd4") != null)
            {
                this.newDataTable.Rows[0]["dia_chi"] = document.getElementById("ctl00_C_UC_ENT_LIST1_CtlList_ctl02_Cmd4").innerText;
            }
            else
            {
                this.newDataTable.Rows[0]["dia_chi"] = "";
            }
            string url = @"https://google.com";
            Webctl.Source = new Uri(url);
        }
        void LoadDocument()
        {
            string url = @"https://dichvuthongtin.dkkd.gov.vn/inf/default.aspx";
            Webctl.Source = new Uri(url);
            Webctl.LoadCompleted += Webctl_LoadCompleted;
        }

        bool Isclick = false;
        public void geturl()
        {
            document = Webctl.Document as HTMLDocument;
            //  mshtml.IHTMLElement dom =document.getElementById("ctl00_RptProdGroups_ctl02_LiActiveProdGroup");
            if (!Isclick)
            {
                Isclick = true;
                IHTMLElement element = (IHTMLElement)document.createElement("a");
                element.setAttribute("href", "javascript:__doPostBack('ctl00$RptProdGroups$ctl02$LnkActiveProdGroup','')");
                element.click();
            }

            if (document.getElementById("ctl00_C_UC_ENT_LIST1_ENTERPRISE_GDT_CODEFilterFld") != null)
            {
                document.getElementById("ctl00_C_UC_ENT_LIST1_ENTERPRISE_GDT_CODEFilterFld").innerText = this.txtma_so_thue.Text.Trim();
                var button = document.getElementById("ctl00_C_UC_ENT_LIST1_BtnFilter");
                button.click();
                Starttimer.Start();
            }
        }
        public System.Windows.Forms.HtmlDocument GetHtmlDocument(string html)
        {
            System.Windows.Forms.WebBrowser browser = new System.Windows.Forms.WebBrowser();
            browser.DocumentText = html;
            browser.Document.Write(html);
            return browser.Document;
        }
        private ClassMasothue GetDatafromSISAPI(string url)
        {
            ClassMasothue custom = null;
            try
            {
                WebRequest request = WebRequest.Create(url);
                //request.Credentials = CredentialCache.DefaultCredentials;
                //request.Timeout = 150;
                WebResponse response = request.GetResponse();
                string status = ((HttpWebResponse)response).StatusDescription;
                using (Stream dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    custom = JsonConvert.DeserializeObject<ClassMasothue>(responseFromServer);
                }
                this.newDataTable.Rows[0]["ten_kh"] = custom.TenChinhThuc;
                this.newDataTable.Rows[0]["doi_tac"] = custom.ChuDoanhNghiep;
                this.newDataTable.Rows[0]["dia_chi"] = custom.DiaChiGiaoDichChinh;
                this.newDataTable.Rows[0]["dien_thoai"] = custom.SoDienThoai;
                response.Close();
                return custom;
            }
            catch (Exception ex)
            {
                return custom;
            }

        }
        private Customerinfo GetData(string url)
        {
            //geturl();
            Customerinfo custom = null;
            try
            {
                WebRequest request = WebRequest.Create(url);
                //request.Credentials = CredentialCache.DefaultCredentials;
                //request.Timeout = 150;
                WebResponse response = request.GetResponse();
                string status = ((HttpWebResponse)response).StatusDescription;
                using (Stream dataStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(dataStream);
                    string responseFromServer = reader.ReadToEnd();
                    custom = JsonConvert.DeserializeObject<Customerinfo>(responseFromServer);
                }
                response.Close();
                return custom;
            }
            catch (Exception ex)
            {
                return custom;
            }

        }
        private void txtma_so_thue_LostFocus(object sender, RoutedEventArgs e)
        {
            /* if (!string.IsNullOrEmpty(this.txtma_so_thue.Text))
             {
                 this.newDataTable.Rows[0]["ten_kh"] = "";
                 this.newDataTable.Rows[0]["ten_kh2"] = "";
                 this.newDataTable.Rows[0]["dia_chi"] = "";
                 Isclick = false;
                 LoadDocument();
             }
            */
        }

        private void Webctl_LoadCompleted(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            geturl();
        }
        private void LoadApi()
        {
            if (!string.IsNullOrEmpty(this.txtma_so_thue.Text))
            {
                string url = "https://thongtindoanhnghiep.co/api/company/" + this.txtma_so_thue.Text.Trim();
                Customerinfo custom = GetData(url);
                if (custom != null)
                {
                    this.newDataTable.Rows[0]["ten_kh"] = custom.Title;
                    this.newDataTable.Rows[0]["ten_kh2"] = custom.TitleEn;
                    this.newDataTable.Rows[0]["doi_tac"] = custom.GiamDoc;
                    this.newDataTable.Rows[0]["dia_chi"] = custom.DiaChiNhanThongBaoThue;
                    this.newDataTable.Rows[0]["dien_thoai"] = custom.NoiNopThue_DienThoai;
                    this.newDataTable.Rows[0]["fax"] = custom.NoiNopThue_Fax;
                }
            }
        }

        private void txtma_so_thue_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtma_so_thue.Text))
            {

                if (e.Key == Key.F5)
                {
                    //LoadApi();
                    string mst = this.txtma_so_thue.Text.Trim();
                    ClassMasothue Curmst = GetDatafromSISAPI(url + "/" + mst);
                }
                if (e.Key == Key.F6)
                {
                    LoadDocument();
                }
            }
        }

        private void btnTracuu_Click(object sender, RoutedEventArgs e)
        {

            string mst = string.IsNullOrEmpty(this.txtma_so_thue.Text.Trim()) ? "0101218690" : this.txtma_so_thue.Text.Trim();
            bool isUpdate = false;
            if(StartUp.currActionTask == ActionTask.Copy || StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
            {
                isUpdate = true;
            }
            try
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo();
                // Arguments to pass to another process, separated by a space
                processStartInfo.Arguments = mst + " " + isUpdate.ToString().ToUpper();
                // Name of the exe file. Note: it will be copied by the builder (Visual Studio) to the projects bin folder, as it is dependent on the project "BlodDavid_CounterProject"
                processStartInfo.FileName = Path.Combine(StartupBase.SasObj.M_StartUp_Path, @"Net472Prog\SmSearchTax.exe");
                // To redirect the output stream, we need to set it to false
                processStartInfo.UseShellExecute = false;
                // We want to capture the data written to the output of the other process
                processStartInfo.RedirectStandardOutput = true;
                // We don't want to show the users another windows, just the output in our main app
                processStartInfo.CreateNoWindow = false;

                // Starting the another process
                var proc = Process.Start(processStartInfo);

                // Listening to the output stream
                while (!proc.StandardOutput.EndOfStream)
                {
                    string output = proc.StandardOutput.ReadToEnd();
                    if (!string.IsNullOrEmpty(output.Trim()) && isUpdate)
                    {
                        string XmlString = Base64Decode(output);

                        DataTable dt = new DataTable();
                        dt.ReadXml(new StringReader(XmlString.ToString()));
                        this.txtma_so_thue.Text = dt.Rows[0]["value"].ToString().Trim();
                        this.txtten_kh.Text = dt.Rows[3]["value"].ToString().Trim();
                        this.txtten_kh2.Text = dt.Rows[4]["value"].ToString().Trim();
                        this.txtten_kh3.Text = dt.Rows[4]["value"].ToString().Trim();
                        this.txtdia_chi.Text = dt.Rows[6]["value"].ToString().Trim();
                        this.txtdia_chi_gd.Text = dt.Rows[8]["value"].ToString().Trim();
                        this.txtdoi_tac.Text = dt.Rows[23]["value"].ToString().Trim();
                        this.txtngay_thanh_lap.Value = Convert.ToDateTime(dt.Rows[1]["value"].ToString().Trim());
                    }
                }
                // Don't let our program to exit until the other process finishes
                proc.WaitForExit();

                var handle = Process.GetCurrentProcess().MainWindowHandle;
                SasControls.ControlLib.WinAPISenkey.SenKey(ModifierKeys.None, Key.LeftShift);
                newForm.SetForegroundWindow(handle);
                newForm.SetActiveWindow(handle);
                newForm.SwitchToThisWindow(handle, true);

                this.Activate();
                this.Focus();
                if (isUpdate)
                {
                    this.txtma_so_thue.Focus();
                }
            }
            catch(Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
    }
}
