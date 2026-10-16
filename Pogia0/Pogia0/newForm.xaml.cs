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

namespace Pogia0
{
    public partial class newForm : FormList
    {
        private DataTable newDataTable = new DataTable();
        private DataTable OldRow = (DataTable)null;
        private string[] OldKeyValue = new string[2];
        private string[] OldKeyCol = new string[2];
        private string strOldNgay_nhap = string.Empty;
        private string strOldMa_vt = string.Empty;
        private EditModeBindingObject FormInEditMode;
        private CodeValueBindingObject M_IP_GIA;
        private CodeValueBindingObject M_IP_GIA_NT;

        public newForm()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.FormInEditMode = (EditModeBindingObject)this.FindResource((object)"IsInEditMode");
            this.M_IP_GIA = (CodeValueBindingObject)this.FindResource((object)nameof(M_IP_GIA));
            this.M_IP_GIA_NT = (CodeValueBindingObject)this.FindResource((object)nameof(M_IP_GIA_NT));
            if (StartUp.currActionTask != ActionTask.View)
                return;
            this.ConfirmGV.ButtonType = 1;
        }

        private void LoadForm()
        {
            ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName);
            this.OldKeyCol = StartUp.SqlTableKey.Split(';');
            this.OldKeyValue = StartUp.currSqlTableKey.Split(';');
            this.Title = SysFunc.Cat_Dau(StartUp.titleWindow);
        }

        private void newForm_Loaded(object sender, RoutedEventArgs e)
        {
            this.M_IP_GIA.Text = this.BindingSasObj.GetOption("M_IP_GIA").ToString();
            this.M_IP_GIA_NT.Text = this.BindingSasObj.GetOption("M_IP_GIA_NT").ToString();
            this.LoadForm();
            TextBox child = SysFunc.FindChild<TextBox>((DependencyObject)this, "txt" + StartUp.SqlTableKey);
            if (child != null)
            {
                child.SelectAll();
                child.Focus();
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
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                    this.txtNgayHL.Focusable = false;
                    this.txtGia_VND.Focusable = false;
                    this.txtGia_NT.Focusable = false;
                    this.ConfirmGV.pnlButton.btnCancel.Focus();
                    break;
                case ActionTask.Add:
                    try
                    {
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        DataRow row = this.newDataTable.NewRow();
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            row["ngay_nhap"] = (object)DateTime.Today;
                            row["gia0"] = (object)0;
                            row["gia_nt0"] = (object)0;
                        }
                        this.newDataTable.Rows.Add(row);
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                    this.txtMa_vt.IsFocus = true;
                    break;
                case ActionTask.Edit:
                    try
                    {
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        if (this.newDataTable.Rows.Count > 0)
                            this.OldRow = this.newDataTable.Copy();
                        this.strOldMa_vt = this.OldKeyValue[0];
                        this.strOldNgay_nhap = this.OldKeyValue[1];
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                    this.txtMa_vt.IsFocus = true;
                    break;
                case ActionTask.Copy:
                    try
                    {
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableView);
                        this.strOldMa_vt = this.OldKeyValue[0];
                        this.strOldNgay_nhap = this.OldKeyValue[1];
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                    this.txtMa_vt.IsFocus = true;
                    break;
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
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
                this.newDataTable.Rows[0]["user_name"] = (object)str;
                if (this.OldRow != null)
                    ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj);
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
                ListFunc.inserRowInDataBase(StartUp.sqlTableName, this.newDataTable.Rows[0], StartupBase.SasObj);
            }
            StartUp.LastEditTable = this.newDataTable;
        }

        private bool CheckValid()
        {
            bool flag1 = true;
            if (this.txtMa_vt.Text.Trim() == string.Empty && flag1)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2370, StartupBase.SasObj, "Chưa vào mã vật tư!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_vt.Focus();
                flag1 = false;
            }
            if (this.txtNgayHL.dValue == new DateTime() && flag1)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2371, StartupBase.SasObj, "Chưa vào ngày hiệu lực!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgayHL.Focus();
                flag1 = false;
            }
            if (this.txtMa_vt.Text.Trim() != "" && flag1)
            {
                if (flag1)
                {
                    try
                    {
                        bool flag2 = true;
                        if (StartUp.currActionTask == ActionTask.Edit && flag1)
                        {
                            flag2 = false;
                            if (this.OldRow.Rows[0]["ma_vt"].ToString().Trim() + this.OldRow.Rows[0]["ngay_nhap"].ToString().Trim() !=  this.newDataTable.Rows[0]["ma_vt"].ToString().Trim() + this.newDataTable.Rows[0]["ngay_nhap"].ToString().Trim())
                                flag2 = true;
                        }
                        SqlCommand sqlcmd = new SqlCommand("select count(1)\tfrom " + StartUp.sqlTableName + " where ma_vt = @ma_vt and ngay_nhap = @ngay_nhap");                        
                        sqlcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)this.txtMa_vt.Text.Trim();
                        sqlcmd.Parameters.Add("@ngay_nhap", SqlDbType.Char).Value = (object)string.Format("{0:yyyyMMdd}", (object)this.txtNgayHL.dValue);
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num = (int)ExMessageBox.Show(2380, StartupBase.SasObj, "Mã đã có khai báo!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtMa_vt.IsFocus = true;
                            flag1 = false;
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
            if (sender is TextBox && (sender as TextBox).Name == "txtghi_chu" && Keyboard.Modifiers == ModifierKeys.Alt)
                return false;
            return base.IsEnterToPassObject(sender);
        }

        private void FormList_Closed(object sender, EventArgs e)
        {
            StartUp.currActionTask = ActionTask.None;
        }

        private void ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
        {
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

        private void txtMa_vt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMa_vt.RowResult == null)
                this.txtten_vt.Text = "";
            else
                this.txtten_vt.Text = StartupBase.M_LAN.Equals("V") ? this.txtMa_vt.RowResult["ten_vt"].ToString() : this.txtMa_vt.RowResult["ten_vt2"].ToString();
        }       
        private void txtGia_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }
    }
}
