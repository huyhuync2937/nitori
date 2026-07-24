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

namespace Inqcs
{
    public partial class newForm : Form
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
            this._confirmGridview.ButtonType = 1;
        }

        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }
        private void newForm_Loaded(object sender, RoutedEventArgs e)
        {
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableView);           
            this.txtma_qcs.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_qcs");
            this.txtten_qcs.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_qcs");
          
            //this.txt_ghichu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ghi_chu");

            this.Title = SysFunc.Cat_Dau(StartUp.titleWindow);
            TextBox KeyTextBox = SysFunc.FindChild<TextBox>((DependencyObject)this, "txt" + StartUp.SqlTableKey);
            if (KeyTextBox != null)
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   KeyTextBox.SelectAll();
                   KeyTextBox.Focus();
               }));
            else
                Debug.Write("Findchild not found");
            switch (StartUp.currActionTask)
            {
                case ActionTask.View:
                    try
                    {
                        StartUp.LastEditRow = (DataRow)null;
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                        this._confirmGridview.ButtonType = 1;
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                        DataRow row = this.newDataTable.NewRow();
                        row["status"] = (object)"1";
                        row["ma_dvcs"] =(object)StartupBase.SasObj.M_ma_dvcs;
                        row["ma_qcs"] = (object)Createnew_ma();
                        this.newDataTable.Rows.Clear();
                        this.newDataTable.Rows.Add(row);
                        this.txtma_qcs.Focus();
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                        if (this.newDataTable.Rows.Count > 0)
                        {
                            this.newDataTable.Rows[0][StartUp.SqlTableKey] = (object)this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            this.OldRow = this.newDataTable.Copy();
                            this.txtma_qcs.Focus();
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                        this.newDataTable.Rows[0]["status"] = (object)"1";                       
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);                                
                            }
                            this.txtma_qcs.Focus();
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

            this._confirmGridview.DataContext = (object)this.newDataTable;
        }
       
        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtstatus.Text == ""))
                return;
            this.newDataTable.Rows[0]["status"] = (object)0;
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
        private void saveCustomer()
        {
            StartUp.LastEditRow = (DataRow)null;
            if (StartUp.currActionTask == ActionTask.View)
                return;
            int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            string str = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
            this.newDataTable.AcceptChanges();
           
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
         
                if (this.OldRow != null)
                    ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            else
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
            
                this.newDataTable.Rows[0]["date0"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time0"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id0"] = (object)num;
             
                ListFunc.inserRowInDataBase(StartUp.sqlTableName, this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            StartUp.LastEditRow = this.newDataTable.Select()[0];
        }
        private string Createnew_ma()
        {
            string str = "";
            if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
            {
                if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                {
                    str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);
                }
                if (string.IsNullOrEmpty(str))
                {
                    str = SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                }
                this.txtma_qcs.Focus();
                if (string.IsNullOrEmpty(str))
                    this.txtma_qcs.Focus();
            }
            return str;
        }
        public void EnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        //private void txtNgay_bh_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (this.txtNgay_bh.Value == DBNull.Value)
        //        this.txtNgay_bh.Value = (object)DateTime.Now;
        //    if (this.txtNgay_bh.IsFocusWithin)
        //        return;
        //}

        private bool CheckValid()
        {
            bool flag1 = true;
            if (flag1 && this.txtma_qcs.Text.Trim() == string.Empty)
            {              
                int num = (int)ExMessageBox.Show(1920, StartupBase.SasObj, "Chưa vào mã [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_qcs.Focus();
                flag1 = false;
            }
            if (flag1 && this.txtma_qcs.Text.Trim() == string.Empty)
            {
                int num = (int)ExMessageBox.Show(1922, StartupBase.SasObj, "Chưa vào tên [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_qcs.Focus();
                flag1 = false;
            }
            if (flag1 && this.txtma_qcs.Text.Trim() != string.Empty)
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtma_qcs.Text.TrimEnd());
                if (str != "" && flag1)
                {                   
                    int num = (int)ExMessageBox.Show(2220, StartupBase.SasObj, "Mã không được chứa các ký tự [" + str + "] !", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtma_qcs.SelectAll();
                    this.txtma_qcs.Focus();
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
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)this.txtma_qcs.Text.Trim();
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                           
                            int num = (int)ExMessageBox.Show(2225, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_qcs.SelectAll();
                            this.txtma_qcs.Focus();
                            flag1 = false;
                        }
                        if (flag1 && flag2)
                        {
                            string Value_old = "";
                            if (StartUp.currActionTask == ActionTask.Edit)
                                Value_old = this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            if (SysFunc.CheckStringContain(StartupBase.SasObj, StartUp.sqlTableName, StartUp.SqlTableKey, this.txtma_qcs.Text.Trim(), Value_old))
                            {
                              
                                int num = (int)ExMessageBox.Show(2230, StartupBase.SasObj, "Mã đã có hoặc mã lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_qcs.SelectAll();
                                this.txtma_qcs.Focus();
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
           
            return flag1;
        }

        private void FormList_Closed(object sender, EventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None)
                SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.currActionTask = ActionTask.None;
        }
        private void FormList_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Escape)
                return;
            StartUp.LastEditRow = (DataRow)null;
            this.Close();
        }

        private void _confirmGridview_OnCancel(object sender, RoutedEventArgs e)
        {
            SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.LastEditRow = (DataRow)null;
            this.Close();
        }

        private void _confirmGridview_OnOk(object sender, RoutedEventArgs e)
        {
            try
            {               
                if (StartUp.currActionTask == ActionTask.View)
                {
                    this.Close();
                    return;
                }    
                else if (StartUp.currActionTask == ActionTask.Edit)
                {
                    if (!CheckValid())
                        return;
                    saveCustomer();
                }
                else
                {
                    string ma_qcs = this.txtma_qcs.Text.Trim();
                    string _listSQlKeyValue1 = "" + ma_qcs;
                    string sqlfilter1 = StartUp.CreateSqlfilter(StartUp.SqlTableKey, _listSQlKeyValue1);
                    if (this.BindingSasObj.ExcuteReader(new SqlCommand("Select * from " + StartUp.sqlTableName + " WHERE " + sqlfilter1)).Tables[0].Rows.Count != 0)
                    {
                        int num = (int)ExMessageBox.Show(1900, StartupBase.SasObj, "Trùng mã [" + StartUp.TableName + "]!", " ", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtma_qcs.Focus();
                        return;
                    }
                    if (!CheckValid())
                        return;
                    saveCustomer();
                }
                StartUp.currActionTask = ActionTask.None;
                this.Close();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void txtstatus_LostFocus_1(object sender, RoutedEventArgs e)
        {

        }

        private void txtma_dvcs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }

        private void txtma_qcs_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
