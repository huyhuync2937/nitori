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

namespace Invt
{
    public partial class newForm : FormList
    {
        private DataTable newDataTable = new DataTable();
        private DataTable OldRow = (DataTable)null;
        private bool isError = false;
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

        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }

        private void LoadForm()
        {
            int formatDecimal = SysFunc.GetFormatDecimal(StartupBase.SasObj.GetOption("M_IP_SL").ToString());
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName);
            this.txtsl_min.Mask = "{double:" + (object)(ListFunc.GetLengthColumn(sqlTableFieldList, "sl_min") - formatDecimal) + "." + formatDecimal.ToString() + "}";
            this.txtsl_min.Mask = "{double:" + (object)(ListFunc.GetLengthColumn(sqlTableFieldList, "sl_min") - formatDecimal) + "." + formatDecimal.ToString() + "}";
            this.txttien.Mask = "{double:" + (object)(ListFunc.GetLengthColumn(sqlTableFieldList, "tien") - formatDecimal) + "." + '5' + "}";
            this.txtma_vt.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_vt");
            this.txtma_tra_cuu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ma_tra_cuu");
            this.txtten_vt.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_vt");
            this.txtten_vt2.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_vt2");
            this.txtdvt.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "dvt");
            this.txtghi_chu.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ghi_chu");
            //this.txtpart_no.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "part_no");
            this.Title = SysFunc.Cat_Dau(StartUp.titleWindow);
        }

        private void newForm_Loaded(object sender, RoutedEventArgs e)
        {
            this.LoadForm();
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
                        if (this.newDataTable.Rows.Count > 0)
                            row["tk_cl_vt"] = this.newDataTable.Rows[0]["tk_cl_vt"];
                        row["vt_ton_kho"] = (object)1;
                        row["gia_ton"] = (object)1;
                        row["sua_tk_vt"] = (object)0;
                        row["sl_min"] = (object)0;
                        row["sl_max"] = (object)0;
                        if (StartUp.Dvcs_ck == "1")
                            row["ma_dvcs"] = (object)this.BindingSasObj.M_ma_dvcs;
                        row["status"] = (object)"1";
                        row["doi_tuong_gt"] = (object)StartupBase.SasObj.GetOption("M_DOI_TUONG_GT").ToString().Trim();
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_vt.MaxLength)
                                    row[StartUp.SqlTableKey] = (object)str;
                            }
                            if (string.IsNullOrEmpty(row[StartUp.SqlTableKey].ToString().Trim()))
                                row[StartUp.SqlTableKey] = (object)SysFunc.GetNewMadm(StartupBase.SasObj, StartUp.sqlTableName);
                        }
                        if (row["tk_cl_vt"].ToString().Trim() == "")
                            row["tk_cl_vt"] = (object)StartupBase.SasObj.GetOption("M_TK_CL_VT").ToString().Trim();
                        this.newDataTable.Rows.Clear();
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
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
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
                            if (KeyTextBox != null)
                                KeyTextBox.IsReadOnly = true;
                            this.isError = true;
                            if (Convert.ToInt16(this.OldRow.Rows[0]["vt_ton_kho"]) != (short)1)
                            {

                            }
                        }
                        TextBox NameTextBox = SysFunc.FindChild<TextBox>((DependencyObject)this, "txtten_vt");
                        if ((KeyTextBox == null || KeyTextBox.IsReadOnly) && NameTextBox != null)
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                            {
                                NameTextBox.SelectAll();
                                NameTextBox.Focus();
                            }));
                        else if (KeyTextBox != null)
                            KeyTextBox.Focus();
                        this.ktraPS();
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
                        this.newDataTable.Rows[0]["doi_tuong_gt"] = (object)StartupBase.SasObj.GetOption("M_DOI_TUONG_GT").ToString().Trim();
                        if (StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1"))
                        {
                            if (!string.IsNullOrEmpty(StartUp.currSqlTableKey) && StartupBase.SasObj.DmdmInfo.Select("ma_dm like '" + StartUp.sqlTableName + "' and  increase_type = 2").Length > 0)
                            {
                                string str = SysFunc.IncreaseCode(StartupBase.SasObj, StartUp.currSqlTableKey, StartUp.SqlTableKey, StartUp.sqlTableName);
                                if (!string.IsNullOrEmpty(str) && str.Length <= this.txtma_vt.MaxLength)
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
            if (this.newDataTable.Columns.Contains("user_id3") && this.newDataTable.Rows[0]["user_id3"].ToString() == "")
                this.newDataTable.Rows[0]["user_id3"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
            this.gridMain.DataContext = this.gridMainB.DataContext = this.gridMainC.DataContext = (object)this.newDataTable;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
            {
                if (this.txtgia_ton.Text == "0")
                    this.txtgia_ton.Text = "";
                if (!string.IsNullOrEmpty(this.txtgia_ton.Text.Trim()))
                {
                    this.txtgia_ton.SearchInit();
                    if (this.txtgia_ton.RowResult != null)
                        this.txtgia_ton_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
                }
                this.txtma_dvcs.SearchInit();
                this.txtma_dvcs_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
                this.txtloai_vt.SearchInit();
                this.txtloai_vt_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
                this.txttk_vt.SearchInit();
                this.txttk_vt_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
            }));
        }

        public void ktraPS()
        {
            string str = this.newDataTable.Rows[0][StartUp.SqlTableKey].ToString().Trim();
            SqlCommand sqlcmd = new SqlCommand("select ma_vt from cdvtvv where ma_vt='" + str.Trim() + "'");
            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            try
            {
                if (!(table.Rows[0][0].ToString().Trim() == str.Trim()))
                    return;
                this.txtma_vt.IsReadOnly = true;
            }
            catch (Exception ex)
            {
            }
        }

        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtstatus.Text == ""))
                return;
            this.newDataTable.Rows[0]["status"] = (object)0;
        }

        private void saveCustomer()
        {
            StartUp.LastEditRow = (DataRow)null;
            if (StartUp.currActionTask == ActionTask.View)
                return;
            int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            string str = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
            this.newDataTable.AcceptChanges();
            if (this.newDataTable.Columns.Contains("search"))
                SysFunc.SetStrSearch(StartupBase.SasObj, StartUp.sqlTableName, ref this.newDataTable);
            this.FixEmptyDate(this.newDataTable.Rows[0], "ngay_hieu_luc");
            this.FixEmptyDate(this.newDataTable.Rows[0], "expiration");
            if (this.newDataTable.Rows[0]["vt_ton_kho"].ToString().Equals("0"))
                this.newDataTable.Rows[0]["gia_ton"] = (object)0;
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                this.newDataTable.Rows[0]["date"] = (object)DateTime.Now;
                this.newDataTable.Rows[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                this.newDataTable.Rows[0]["user_id"] = (object)num;
                this.newDataTable.Rows[0]["user_name"] = (object)str;

                foreach (DataColumn col in this.newDataTable.Columns)
                {
                    Debug.WriteLine($"{col.ColumnName} : {col.DataType}");
                }

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
            StartUp.LastEditRow = this.newDataTable.Select()[0];
        }

        private void FixEmptyDate(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName))
                return;
            object value = row[columnName];
            if (value == null || value == DBNull.Value)
                return;
            if (value is DateTime dateTime && dateTime < new DateTime(1753, 1, 1))
                row[columnName] = (object)DBNull.Value;
            else if (value is string text && string.IsNullOrEmpty(text))
                row[columnName] = (object)DBNull.Value;
        }

        private bool CheckValid()
        {
            bool flag1 = true;
            if (flag1 && this.txtma_vt.Text.Trim() == string.Empty)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2215, StartupBase.SasObj, "Chua v�o m� [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtma_vt.Focus();
                flag1 = false;
            }
            if (flag1 && this.txtma_vt.Text.Trim() != string.Empty)
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtma_vt.Text.TrimEnd());
                if (str != "" && flag1)
                {
                    this.TabInfor.SelectedIndex = 0;
                    int num = (int)ExMessageBox.Show(2220, StartupBase.SasObj, "M� kh�ng du?c ch?a c�c k� t? [" + str + "] !", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtma_vt.SelectAll();
                    this.txtma_vt.Focus();
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
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)this.txtma_vt.Text.Trim();
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0 && flag2)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            int num = (int)ExMessageBox.Show(2225, StartupBase.SasObj, "M� d� c� ho?c m� l?ng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtma_vt.SelectAll();
                            this.txtma_vt.Focus();
                            flag1 = false;
                        }
                        if (flag1 && flag2)
                        {
                            string Value_old = "";
                            if (StartUp.currActionTask == ActionTask.Edit)
                                Value_old = this.OldRow.Rows[0][StartUp.SqlTableKey].ToString().Trim();
                            if (SysFunc.CheckStringContain(StartupBase.SasObj, StartUp.sqlTableName, StartUp.SqlTableKey, this.txtma_vt.Text.Trim(), Value_old))
                            {
                                this.TabInfor.SelectedIndex = 0;
                                int num = (int)ExMessageBox.Show(2230, StartupBase.SasObj, "M� d� c� ho?c m� l?ng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtma_vt.SelectAll();
                                this.txtma_vt.Focus();
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
            if (flag1 && this.txtten_vt.Text.Trim() == string.Empty)
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2235, StartupBase.SasObj, "Chưa vào tên [" + StartUp.TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtten_vt.Focus();
                flag1 = false;
            }
            if (flag1 && this.isError && (this.OldRow.Rows[0]["dvt"].ToString().Trim() != "" && this.txtdvt.Text.Trim() == ""))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2240, this.BindingSasObj, "�� c� ph�t sinh, don v? t�nh kh�ng du?c d? tr?ng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtdvt.Focus();
                flag1 = false;
            }
            if (flag1 && this.txtvt_ton_kho.Value.ToString() == "1" && this.txtgia_ton.Text.Trim() == "")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2245, this.BindingSasObj, "Chưa chọn cách tính giá tồn kho!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtgia_ton.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && this.txtvt_ton_kho.Text.Trim() == "1" && this.txtloai_vt.Text.Trim() == "")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(4250, StartupBase.SasObj, "Chưa vào loại vật tư!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtloai_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txtloai_vt.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2250, StartupBase.SasObj, "Loại vật tư không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtloai_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && this.txttk_vt.Text.Trim() == string.Empty && this.txtvt_ton_kho.Text.Trim() == "1")
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2255, StartupBase.SasObj, "Chưa vào tài khoản kho!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txttk_vt.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2260, StartupBase.SasObj, "Tài khoản kho không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_vt.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2260, StartupBase.SasObj, "Tài khoản kho không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && StartUp.currActionTask == ActionTask.Edit)
            {
                if (this.OldRow.Rows[0]["tk_vt"].ToString().Trim() != this.newDataTable.Rows[0]["tk_vt"].ToString().Trim())
                {
                    try
                    {
                        SqlCommand sqlcmd = new SqlCommand("select Top(1)* from ct70 where ma_vt=@ma_vt and tk_vt=@tk_vt");
                        sqlcmd.Parameters.Add("@" + StartUp.SqlTableKey, SqlDbType.Char).Value = (object)StartUp.currSqlTableKey;
                        sqlcmd.Parameters.Add("@tk_vt", SqlDbType.Char).Value = (object)this.OldRow.Rows[0]["tk_vt"].ToString();
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                        {
                            this.TabInfor.SelectedIndex = 0;
                            if (ExMessageBox.Show(2265, StartupBase.SasObj, "T�i kho?n kho cu d� c� ph�t sinh, c� ti?p t?c kh�ng?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                            {
                                this.txttk_vt.IsFocus = true;
                                flag1 = false;
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                }
            }

            //
            if (flag1 && Checkdmtk(this.txttk_dt.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Tài khoản doanh thu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_dt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_tl.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2275, StartupBase.SasObj, "Tài khoản hàng bán bị trả lại không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_tl.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_cl_vt.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Tài khoản chênh lệch giá vật tư không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_cl_vt.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_ck.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2290, StartupBase.SasObj, "Tài khoản chiết khấu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_ck.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_nvl.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2295, StartupBase.SasObj, "Tài khoản nguyên vật liệu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_nvl.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_gv.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2300, StartupBase.SasObj, "Tài khoản giá vốn không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_gv.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_spdd.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2305, StartupBase.SasObj, "Tài khoản sản phẩm dở dang không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_spdd.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && Checkdmtk(this.txttk_km.Text))
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2310, StartupBase.SasObj, "T�i kho?n cp khuy?n m�i kh�ng h?p l?!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_km.IsFocus = true;
                flag1 = false;
            }

            //
            if (flag1 && !this.txttk_dt.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Tài khoản doanh thu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_dt.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_tl.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2275, StartupBase.SasObj, "Tài khoản hàng bán bị trả lại không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_tl.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_cl_vt.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Tài khoản chênh lệch giá vật tư không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_cl_vt.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_ck.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Tài khoản chiết khấu không hợp lệ!!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_ck.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_nvl.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2295, StartupBase.SasObj, "Tài khoản nguyên vật liệu không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_nvl.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_gv.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2300, StartupBase.SasObj, "Tài khoản giá vốn không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_gv.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_spdd.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2305, StartupBase.SasObj, "Tài khoản sản phẩm dở dang không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_spdd.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txttk_km.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2310, StartupBase.SasObj, "T�i kho?n cp khuy?n m�i kh�ng h?p l?!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txttk_km.IsFocus = true;
                flag1 = false;
            }

            if (flag1 && !this.txtnh_vt1.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2315, StartupBase.SasObj, "Nhóm vật tư 1 không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtnh_vt1.IsFocus = true;
                flag1 = false;
            }
            if (flag1 && !this.txtnh_vt2.CheckLostFocus())
            {
                this.TabInfor.SelectedIndex = 0;
                int num = (int)ExMessageBox.Show(2320, StartupBase.SasObj, "Nhóm vật tư 2 không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtnh_vt2.IsFocus = true;
                flag1 = false;
            }
            //if (flag1 && !this.txtnh_vt3.CheckLostFocus())
            //{
            //    this.TabInfor.SelectedIndex = 0;
            //    int num = (int)ExMessageBox.Show(2325, StartupBase.SasObj, "Nhóm vật tư 3 không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtnh_vt3.IsFocus = true;
            //    flag1 = false;
            //}
            //if (flag1 && this.newDataTable.Columns.Contains("dinh_dang_qcs"))
            //{
            //    string dinhDangQcs = Convert.ToInt32(this.newDataTable.Rows[0]["dinh_dang_qcs"]).ToString().Trim();
            //    if (dinhDangQcs != string.Empty && dinhDangQcs != "1" && dinhDangQcs != "2")
            //    {
            //        this.TabInfor.SelectedIndex = 2;
            //        int num = (int)ExMessageBox.Show(2325, StartupBase.SasObj,
            //            "Định dạng QCS chỉ được nhập 1 - Dọc hoặc 2 - Ngang!",
            //            "Xác nhận nhập liệu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //        this.txtdinh_dang_qcs.Focus();
            //        flag1 = false;
            //    }
            //}

            return flag1;
        }

        public void EnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        private void txtloai_vt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtloai_vt.RowResult == null)
                this.txtten_loai_vt.Text = "";
            else
                this.txtten_loai_vt.Text = StartupBase.M_LAN.Equals("V") ? this.txtloai_vt.RowResult["ten"].ToString() : this.txtloai_vt.RowResult["ten2"].ToString();
        }

        private void txttk_vt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txttk_vt.RowResult == null)
                this.txtten_tk_vt.Text = "";
            else
                this.txtten_tk_vt.Text = StartupBase.M_LAN.Equals("V") ? this.txttk_vt.RowResult["ten_tk"].ToString() : this.txttk_vt.RowResult["ten_tk2"].ToString();
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

        private void txtvt_ton_kho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtvt_ton_kho.Text.Trim() == ""))
                return;
            this.newDataTable.Rows[0]["vt_ton_kho"] = (object)0;
        }

        private void txtsua_tk_kho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtsua_tk_kho.IsFocusWithin || !(this.txtsua_tk_kho.Text.Trim() == ""))
                return;
            this.newDataTable.Rows[0]["sua_tk_vt"] = (object)0;
        }

        private void txtsl_min_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtsl_min.Text.Trim() == ""))
                return;
            this.txtsl_min.Value = (object)0;
        }

        private void txtsl_max_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtsl_max.Text.Trim() == ""))
                return;
            this.txtsl_max.Value = (object)0;
        }

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.Return))
                return;
            TextBox textBox = sender as TextBox;
            textBox.SelectedText = Environment.NewLine;
            ++textBox.SelectionStart;
            textBox.SelectionLength = 1;

            this.txtghi_chu.Focus();
        }

        protected override bool IsEnterToPassObject(object sender)
        {
            return (!(sender is TextBox) || !((sender as TextBox).Name == "txtghi_chu")) && base.IsEnterToPassObject(sender);
        }

        private void txtma_vt_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_vt.Text = this.txtma_vt.Text.Trim();
            this.txtma_vt.SelectAll();
        }

        private void txtma_tra_cuu_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtma_tra_cuu.Text = this.txtma_tra_cuu.Text.Trim();
            this.txtma_tra_cuu.SelectAll();
        }

        private void txtdvt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.isError)
            {
                if (!(this.OldRow.Rows[0]["dvt"].ToString().Trim() != "") || !(this.txtdvt.Text.Trim() == ""))
                    return;
                int num = (int)ExMessageBox.Show(2330, this.BindingSasObj, "�� c� ph�t sinh, don v? t�nh kh�ng du?c d? tr?ng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                if (!(this.txtdvt.Text.Trim() == ""))
                    return;
                this.newDataTable.Rows[0]["vt_ton_kho"] = (object)0;
            }
        }
        //private void txtpart_no_GotFocus(object sender, RoutedEventArgs e)
        //{
        //    this.txtpart_no.Text = this.txtpart_no.Text.Trim();
        //    this.txtpart_no.SelectAll();
        //}
        private void FormList_Closed(object sender, EventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None)
                SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.currActionTask = ActionTask.None;
        }

        private void txtgia_ton_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtgia_ton.RowResult == null)
                this.txtten_gia_ton.Text = "";
            else
                this.txtten_gia_ton.Text = !(StartupBase.M_LAN == "V") ? this.txtgia_ton.RowResult["ten_loai2"].ToString() : this.txtgia_ton.RowResult["ten_loai"].ToString();
        }

        private void ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
        {
            SysFunc.RollbackMadm(this.BindingSasObj, StartUp.sqlTableName);
            StartUp.LastEditRow = (DataRow)null;
            this.Close();
        }

        private void FormList_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Escape)
                return;
            StartUp.LastEditRow = (DataRow)null;
            this.Close();
        }

        private void txtvt_ton_kho_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            this.txtgia_ton.Text = Convert.ToInt32(e.NewValue) == 0 ? "" : "1";
            this.txtgia_ton.SearchInit();
            this.txtgia_ton_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
        }

        private void txtma_dvcs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.tblten_ma_dvcs.Text = this.txtma_dvcs.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtma_dvcs.RowResult["ten_dvcs"].ToString() : this.txtma_dvcs.RowResult["ten_dvcs2"].ToString());
        }

        private bool Checkdmtk(string strtk)
        {
            bool flag = false;
            if (!string.IsNullOrEmpty(strtk))
            {
                DataTable dt = this.BindingSasObj.ExcuteReader(new SqlCommand(string.Format("SELECT tk_me FROM dmtk where tk_me like '{0}' and tk_me <> ''", strtk.ToString().Trim()))).Tables[0];
                if (dt != null && dt.Rows.Count > 0)
                    flag = true;
            }
            else
            {
                flag = false;
            }
            return flag;
        }
    }
}

