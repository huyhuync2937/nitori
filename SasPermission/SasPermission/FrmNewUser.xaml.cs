using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Threading;

namespace SasPermission
{
    public partial class FrmNewUser : FormFilter
    {
        private DataTable OldRow = (DataTable)null;
        public int rowIndex = -1;
        public int M_LEN_PASS = 0;
        private EditModeBindingObject FormInEditMode;
        private DataTable newDataTable;

        public FrmNewUser()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.FormInEditMode = (EditModeBindingObject)this.FindResource((object)"IsInEditMode");
        }

        private void LoadForm()
        {
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName);
            this.txtuser_name.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "user_name");
            this.txtcomment.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "comment");
            this.chkboxis_admin.IsEnabled = StartUp.is_admin == 1;
            this.chkislocked.IsEnabled = StartUp.is_admin == 1;
            this.txtUser.Filter = StartUp.is_admin == 1 ? "" : "is_admin = 0";
            this.txtMa_dvcs.IsEnabled = StartUp.is_admin == 1;
            if (StartupBase.SasObj.Sysvars.ContainsKey("M_USER_DVCS") && StartupBase.SasObj.GetSysvar("M_USER_DVCS").ToString() != "1")
            {
                this.txtMa_dvcs.IsEnabled = false;
                this.Height -= 25.0;
            }
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               this.txtMa_dvcs.SearchInit();
               if (this.txtMa_dvcs.RowResult != null)                   
                this.lblTen_dvcs.Text = StartupBase.M_LAN == "V" ? this.txtMa_dvcs.RowResult["ten_dvcs"].ToString() : this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString();

               this.txtma_pb.SearchInit();
               if (this.txtma_pb.RowResult != null)
                   this.txtTenbpns.Text = StartupBase.M_LAN == "V" ? this.txtma_pb.RowResult["ten_bpns"].ToString() : this.txtma_pb.RowResult["ten_bpns2"].ToString();

               this.txtma_cv.SearchInit();
               if (this.txtma_cv.RowResult != null)
                   this.txtTencv.Text = StartupBase.M_LAN == "V" ? this.txtma_cv.RowResult["ten_cv"].ToString() : this.txtma_cv.RowResult["ten_cv2"].ToString();
           }));
        }

        private void FrmNewUser_Loaded(object sender, RoutedEventArgs e)
        {
            this.LoadForm();
            int.TryParse(StartupBase.SasObj.GetOption("M_LEN_PASS").ToString().Trim(), out M_LEN_PASS);
            M_LEN_PASS = M_LEN_PASS < 0 ? 0 : M_LEN_PASS;
            this.Title = StartUp.titleWindow;
            this.txtuser_name.Focus();
            switch (StartUp.currActionTask)
            {
                case ActionTask.Add:
                    try
                    {
                        this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
                        this.newDataTable.Rows.Add(this.newDataTable.NewRow());
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
                            this.OldRow = this.newDataTable.Copy();
                            if ((int)Convert.ToInt16(this.OldRow.Rows[0]["user_id"]) == StartUp.user_id)
                            {
                                this.chkboxis_admin.IsEnabled = false;
                                this.chkislocked.IsEnabled = false;
                                this.txtUser.Filter = "is_admin = 0";
                            }
                            this.newDataTable.Rows[0]["user_name"] = (object)this.newDataTable.Rows[0]["user_name"].ToString().Trim();
                            this.newDataTable.Rows[0]["comment"] = (object)this.newDataTable.Rows[0]["comment"].ToString().Trim();
                        }
                        this.txtpassword.IsEnabled = false;
                        this.txtgokt.IsEnabled = false;
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

        public void IsEnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.checkValid())
                return;
            this.saveUser();
            StartUp.currActionTask = ActionTask.None;
            this.Close();
        }

        private void saveUser()
        {
            this.newDataTable.AcceptChanges();
            string NonUpdateColumns = "user_id;duyet_dh";
            if (this.txtpassword.Password != "")
                this.newDataTable.Rows[0]["password"] = (object)this.Encrypt(this.txtpassword.Password);
            else
                NonUpdateColumns += ";password";
            bool? isChecked = this.chkboxis_admin.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
            {
                SqlCommand sqlcmd = new SqlCommand("select rights, r_read, r_add, r_edit, r_del, r_fav from " + StartUp.sqlTableName + " where admin_root = 1");
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                if (dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    this.newDataTable.Rows[0]["rights"] = (object)dataSet.Tables[0].Rows[0]["rights"].ToString();
                    this.newDataTable.Rows[0]["r_read"] = (object)dataSet.Tables[0].Rows[0]["r_read"].ToString();
                    this.newDataTable.Rows[0]["r_add"] = (object)dataSet.Tables[0].Rows[0]["r_add"].ToString();
                    this.newDataTable.Rows[0]["r_edit"] = (object)dataSet.Tables[0].Rows[0]["r_edit"].ToString();
                    this.newDataTable.Rows[0]["r_del"] = (object)dataSet.Tables[0].Rows[0]["r_del"].ToString();
                    this.newDataTable.Rows[0]["r_fav"] = (object)dataSet.Tables[0].Rows[0]["r_fav"].ToString();
                }
            }
            if (!string.IsNullOrEmpty(this.txtUser.Text.Trim()))
            {
                SqlCommand sqlcmd = new SqlCommand("select rights, r_read, r_add, r_edit, r_del,r_fav from " + StartUp.sqlTableName + " where user_name = @user_name");
                sqlcmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar)).Value = (object)this.txtUser.Text;
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                if (dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    this.newDataTable.Rows[0]["rights"] = (object)dataSet.Tables[0].Rows[0]["rights"].ToString();
                    this.newDataTable.Rows[0]["r_read"] = (object)dataSet.Tables[0].Rows[0]["r_read"].ToString();
                    this.newDataTable.Rows[0]["r_add"] = (object)dataSet.Tables[0].Rows[0]["r_add"].ToString();
                    this.newDataTable.Rows[0]["r_edit"] = (object)dataSet.Tables[0].Rows[0]["r_edit"].ToString();
                    this.newDataTable.Rows[0]["r_del"] = (object)dataSet.Tables[0].Rows[0]["r_del"].ToString();
                    this.newDataTable.Rows[0]["r_fav"] = (object)dataSet.Tables[0].Rows[0]["r_fav"].ToString(); 
                }
            }
            if (this.newDataTable.Columns.Contains("search"))
                SysFunc.SetStrSearch(StartupBase.SasObj, StartUp.sqlTableName, ref this.newDataTable);
            if (this.newDataTable.Columns.Contains("user_guid"))
            {
                if (StartUp.currActionTask == ActionTask.Add || Convert.IsDBNull(this.newDataTable.Rows[0]["user_guid"]))
                {
                    this.newDataTable.Rows[0]["user_guid"] = (object)Guid.NewGuid().ToString();
                }
               
            }   
            
            if (StartUp.currActionTask == ActionTask.Edit)
            {
                if (ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj, NonUpdateColumns) == 1)
                {
                    StartUp.dt.Rows[this.rowIndex].AcceptChanges();
                    StartUp.dt.Rows[this.rowIndex].BeginEdit();
                    StartUp.dt.Rows[this.rowIndex].ItemArray = this.newDataTable.Rows[0].ItemArray;
                    StartUp.dt.Rows[this.rowIndex].EndEdit();
                }
                DataRow row = StartUp.dt.Rows[this.rowIndex];
                if ((int)Convert.ToInt16(row["user_id"]) != StartUp.user_id)
                    return;
                string pass = BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(this.txtpassword.Password.ToLower())));
                StartupBase.SasObj.SetUserPassword(pass);
                StartupBase.SasObj.UserInfo.Rows[0]["password"] = row["password"] = (object)pass;
            }
            else if (ListFunc.inserRowInDataBase(StartUp.sqlTableName, this.newDataTable.Rows[0], StartupBase.SasObj, "user_id") == 1)
            {
                this.newDataTable.Rows[0].SetField<int>("user_id", ListFunc.GetValueIdentityCurrent(StartUp.sqlTableName, StartupBase.SasObj));
                StartUp.dt.Rows.Add(this.newDataTable.Rows[0].ItemArray);
                GhiLog();
            }
        }

        public void GhiLog()
        {
            string strsql = "INSERT INTO [dbo].[dmpaslog]([log_id],[machineid],[oldpass],[newpass],[oldname],[user_id],[user_id2],[date_log],[time_log],[status]) " +
                "VALUES(newid(),@machineid,@oldpass,@newpass,@oldname,@user_id,@user_id2,getdate(),convert(varchar(10), GETDATE(), 108),'1')";
            SqlCommand cmd = new SqlCommand(strsql);
            cmd.Parameters.Add("@machineid", SqlDbType.Char).Value = (object)System.Environment.MachineName.Trim();
            cmd.Parameters.Add("@oldpass", SqlDbType.NChar).Value = (object)"";
            cmd.Parameters.Add("@newpass", SqlDbType.NChar).Value = (object)this.txtpassword.Password.Trim();
            cmd.Parameters.Add("@oldname", SqlDbType.NChar).Value = (object)this.newDataTable.Rows[0]["user_name"].ToString().Trim();
            cmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)this.newDataTable.Rows[0]["user_id"];
            cmd.Parameters.Add("@user_id2", SqlDbType.Int).Value = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
            StartupBase.SasObj.ExcuteNonQuery(cmd);
        }
        private string Encrypt(string password)
        {
            password = password.ToLower();
            return BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(password)));
        }

        private bool checkValid()
        {
            bool flag1 = true;
            if (this.txtuser_name.Text.Trim() == string.Empty && flag1)
            {
                int num = (int)ExMessageBox.Show(635, StartupBase.SasObj, "Chưa vào tên!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtuser_name.Focus();
                flag1 = false;
            }
            if (this.txtuser_name.Text.Trim() != string.Empty && flag1)
            {
                string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtuser_name.Text.Trim());
                if (str != "" && flag1)
                {
                    this.TabInfor.SelectedIndex = 0;
                    int num = (int)ExMessageBox.Show(640, StartupBase.SasObj, "Tên không được chứa các ký tự [" + str + "]!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtuser_name.SelectAll();
                    this.txtuser_name.Focus();
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
                            if (this.OldRow.Rows[0][StartUp.SqlTableObjectName].ToString().Trim() != this.newDataTable.Rows[0][StartUp.SqlTableObjectName].ToString().Trim())
                                flag2 = true;
                        }
                        SqlCommand sqlcmd = new SqlCommand("select * from " + StartUp.sqlTableName + " where user_name=@user_name");
                        sqlcmd.Parameters.Add("@user_name", SqlDbType.Char).Value = (object)this.txtuser_name.Text.Trim();
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0 && flag2)
                        {
                            int num = (int)ExMessageBox.Show(645, StartupBase.SasObj, "Tên đã có!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtuser_name.SelectAll();
                            this.txtuser_name.Focus();
                            flag1 = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        ErrorLog.CatchMessage(ex);
                    }
                }
            }
            if (StartUp.currActionTask == ActionTask.Add && flag1)
            {
                // Kiểm tra pass
                if (M_LEN_PASS > 0)
                {
                    if (this.txtpassword.Password.Trim().Length < M_LEN_PASS)
                    {
                        int num0 = (int)ExMessageBox.Show(951, StartupBase.SasObj, "Mật khẩu chưa đủ độ dài!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtpassword.Focus();
                        return false;
                    }

                    if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsDigit)))
                    {
                        int num0 = (int)ExMessageBox.Show(952, StartupBase.SasObj, "Mật khẩu phải chứa số!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtpassword.Focus();
                        return false;
                    }
                    if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsUpper)))
                    {
                        int num0 = (int)ExMessageBox.Show(953, StartupBase.SasObj, "Mật khẩu phải chữ hoa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtpassword.Focus();
                        return false;
                    }

                    if (!(this.txtpassword.Password.Trim().ToCharArray().Any(char.IsLower)))
                    {
                        int num0 = (int)ExMessageBox.Show(954, StartupBase.SasObj, "Mật khẩu phải chữ thường!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtpassword.Focus();
                        return false;
                    }
                }
            }
            if (StartUp.currActionTask == ActionTask.Add && flag1 && this.txtpassword.Password == string.Empty)
            {
                int num = (int)ExMessageBox.Show(653, StartupBase.SasObj, "Chưa vào mật khẩu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtpassword.Focus();
                flag1 = false;
            }
            if (this.txtpassword.Password != this.txtgokt.Password && flag1)
            {
                int num = (int)ExMessageBox.Show(655, StartupBase.SasObj, "Đã gõ mã kiểm tra sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtgokt.SelectAll();
                this.txtgokt.Focus();
                flag1 = false;
            }
            return flag1;
        }

        private void txtMa_dvcs_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMa_dvcs.RowResult != null)
                this.lblTen_dvcs.Text = StartupBase.M_LAN == "V" ? this.txtMa_dvcs.RowResult["ten_dvcs"].ToString() : this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString();
            else
                this.lblTen_dvcs.Text = "";
        }
        private void txtma_bp_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_bp.RowResult != null)
                this.lblten_bp.Text = StartupBase.M_LAN == "V" ? this.txtma_bp.RowResult["ten_bp"].ToString() : this.txtma_bp.RowResult["ten_bp2"].ToString();
            else
                this.lblten_bp.Text = "";
        }
        private void txtma_kho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_kho.RowResult != null)
                this.lblten_kho.Text = StartupBase.M_LAN == "V" ? this.txtma_kho.RowResult["ten_kho"].ToString() : this.txtma_kho.RowResult["ten_kho2"].ToString();
            else
                this.lblten_kho.Text = "";
        }

        private void txtma_nx_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_nx.RowResult != null)
                this.lblten_nx.Text = StartupBase.M_LAN == "V" ? this.txtma_nx.RowResult["ten_nx"].ToString() : this.txtma_nx.RowResult["ten_nx2"].ToString();
            else
                this.lblten_nx.Text = "";
        }
        private void txtma_kh_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_kh.RowResult != null)
                this.lblten_kh.Text = StartupBase.M_LAN == "V" ? this.txtma_kh.RowResult["ten_kh"].ToString() : this.txtma_kh.RowResult["ten_kh2"].ToString();
            else
                this.lblten_kh.Text = "";
        }
        private void txtma_qs_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_qs.RowResult != null)
                this.lblten_qs.Text = StartupBase.M_LAN == "V" ? this.txtma_qs.RowResult["ten_qs"].ToString() : this.txtma_qs.RowResult["ten_qs2"].ToString();
            else
                this.lblten_qs.Text = "";
        }
        private void txtma_qs2_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_qs2.RowResult != null)
                this.lblten_qs2.Text = StartupBase.M_LAN == "V" ? this.txtma_qs2.RowResult["ten_qs"].ToString() : this.txtma_qs2.RowResult["ten_qs2"].ToString();
            else
                this.lblten_qs2.Text = "";
        }

        private void txtbpns_PreviewLostFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtma_pb.Text) && this.txtma_pb.RowResult != null)
            {
                if (StartupBase.M_LAN.Equals("V"))
                {
                    this.txtTenbpns.Text = this.txtma_pb.RowResult["ten_bpns"].ToString();

                }
                else
                {
                    this.txtTenbpns.Text = this.txtma_pb.RowResult["ten_bpns2"].ToString();
                }
            }
            else
            {
                this.txtTenbpns.Text = "";
            }
        }

        private void txtma_cv_PreviewLostFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtma_cv.Text) && this.txtma_cv.RowResult != null)
            {
                if (StartupBase.M_LAN.Equals("V"))
                {
                    this.txtTencv.Text = this.txtma_cv.RowResult["ten_cv"].ToString();

                }
                else
                {
                    this.txtTencv.Text = this.txtma_cv.RowResult["ten_cv2"].ToString();
                }
            }
            else
            {
                this.txtTencv.Text = "";
            }
        }
    }
}
