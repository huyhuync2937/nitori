using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Microsoft.Win32;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Linq;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using Excel = Microsoft.Office.Interop.Excel;

namespace SasPermission
{
    public partial class Permission : Form
    {
        public static string currUser_id = string.Empty;

        public Permission()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
        }

        private void Permission_Loaded(object sender, RoutedEventArgs e)
        {
            this.DisplayLanguage = StartupBase.M_LAN;
            this.GrdUserInfo.DataSource = (IEnumerable)StartUp.dt.DefaultView;
            if (StartUp.is_admin == 0)
                this.btnPhanQuyen.IsEnabled = false;
            if (this.GrdUserInfo.Records.Count <= 0)
                return;
            int M_PQ_WEB = 0;
            int M_PQ_APP = 0;
            int.TryParse(StartupBase.SasObj.GetOption("M_PQ_WEB").ToString().Trim(), out M_PQ_WEB);
            int.TryParse(StartupBase.SasObj.GetOption("M_PQ_APP").ToString().Trim(), out M_PQ_APP);
            if (M_PQ_WEB != 1)
                this.btnPhanQuyenW.Visibility = Visibility.Collapsed;
            if (M_PQ_APP != 1)
                this.btnPhanQuyenA.Visibility = Visibility.Collapsed;

            this.GrdUserInfo.ActiveRecord = (Record)(this.GrdUserInfo.Records[0] as DataRecord);
            this.GrdUserInfo.Focus();
        }

        private void btnThemNSD_Click(object sender, RoutedEventArgs e)
        {
            this.ThemNSD();
        }

        private void ThemNSD()
        {
            if (StartUp.currActionTask != ActionTask.None)
                return;
            int result = 0;
            int.TryParse(StartUp.dtRegInfo.Rows[14]["content"].ToString(), out result);

            if (result > 0)
            {
                SqlCommand sqlcmd = new SqlCommand("Select count(user_id) as sl_user from userinfo");
                DataTable tbluser = StartUp.SasObj.ExcuteReader(sqlcmd).Tables[0];
                if (tbluser.Rows.Count > 0)
                {
                    int sl_user = 0;
                    int.TryParse(tbluser.Rows[0]["sl_user"].ToString(), out sl_user);
                    if (sl_user >= result)
                    {
                        int num = (int)ExMessageBox.Show(650, StartupBase.SasObj, "Số lượng người sử dụng đã đủ số lượng đăng ký ([" + result.ToString().Trim() + "])!", "The number of users has reached the number of registrations", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return;
                    }
                }
            }

            StartUp.currActionTask = ActionTask.Add;
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Them moi nguoi su dung" : "Add new user";
            FrmNewUser frmNewUser = new FrmNewUser();
            frmNewUser.IsEnableEditMode(true);
            frmNewUser.ShowDialog();
            StartUp.currActionTask = ActionTask.None;
            frmNewUser.Closed += new EventHandler(this._form_Closed);
        }

        private void btnQuayRa_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnXoaNSD_Click(object sender, RoutedEventArgs e)
        {
            this.XoaNSD();
        }

        private void XoaNSD()
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            StartUp.currActionTask = ActionTask.Delete;
            if (ExMessageBox.Show(660, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                this.deleteRowByKey(StartUp.sqlTableName);
            StartUp.currActionTask = ActionTask.None;
            this.Reload();
        }

        private void deleteRowByKey(string tableName)
        {
            DataTable row = StartUp.GetRow(tableName);
            if (row.Rows.Count <= 0)
                return;
            int index = (this.GrdUserInfo.ActiveRecord as DataRecord).Index;
            if (ListFunc.deleteRowInDatabaseByKey(tableName, StartUp.SqlTableKey, row.Rows[0], StartupBase.SasObj) == 1)
            {
                StartUp.dt.Rows[index].Delete();
                StartUp.dt.AcceptChanges();
            }
        }

        private void btnSuaDoiNSD_Click(object sender, RoutedEventArgs e)
        {
            this.SuaDoiNSD();
        }

        private void SuaDoiNSD()
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            StartUp.currActionTask = ActionTask.Edit;
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            StartUp.titleWindow = StartupBase.M_LAN.Equals("V") ? "Sua thong tin nguoi su dung" : "Modify information of user";
            FrmNewUser frmNewUser = new FrmNewUser();
            frmNewUser.rowIndex = StartUp.dt.Rows.IndexOf(((this.GrdUserInfo.ActiveRecord as DataRecord).DataItem as DataRowView).Row);
            frmNewUser.IsEnableEditMode(true);
            frmNewUser.ShowDialog();
            StartUp.currActionTask = ActionTask.None;
            frmNewUser.Closed += new EventHandler(this._form_Closed);
        }

        private void btnChangePass_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            FrmChangePassword frmChangePassword = new FrmChangePassword();
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            DataTable row = StartUp.GetRow(StartUp.sqlTableName);
            if ((int)Convert.ToInt16(row.Rows[0]["user_id"]) != StartUp.user_id && StartUp.is_admin != 1)
            {
                int num = (int)ExMessageBox.Show(911, StartupBase.SasObj, "Không có quyền sửa mật khẩu của người sử dụng khác!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                StartUp.currActionTask = ActionTask.Edit;
                frmChangePassword.Title = StartupBase.M_LAN.Equals("V") ? "Sua mat khau" : "Change password";
                frmChangePassword.DataContext = (object)row;
                bool? nullable = frmChangePassword.ShowDialog();
                if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                {
                    DataRowView dataItem = (this.GrdUserInfo.ActiveRecord as DataRecord).DataItem as DataRowView;
                    if ((int)Convert.ToInt16(dataItem["user_id"]) == StartUp.user_id)
                    {
                        string pass = BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(frmChangePassword.txtpassword.Password.ToLower())));
                        StartupBase.SasObj.SetUserPassword(pass);
                        StartupBase.SasObj.UserInfo.Rows[0]["password"] = dataItem["password"] = (object)pass;
                    }
                }
                StartUp.currActionTask = ActionTask.None;
            }
        }

        private void _form_Closed(object sender, EventArgs e)
        {
            StartUp.currActionTask = ActionTask.None;
            this.Reload(); 
        }

        private void btnPhanQuyen_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            FrmPermission frmPermission = new FrmPermission();
            frmPermission.rowIndex = StartUp.dt.Rows.IndexOf(((this.GrdUserInfo.ActiveRecord as DataRecord).DataItem as DataRowView).Row);
            frmPermission.ShowDialog();
            this.Reload();
            frmPermission.Closed += new EventHandler(this._formPQ_Closed);
        }

        private void _formPQ_Closed(object sender, EventArgs e)
        {
            this.Reload();
        }

        private void Form_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.F4) && Keyboard.Modifiers == ModifierKeys.None)
                this.ThemNSD();
            if (Keyboard.IsKeyDown(Key.F3) && Keyboard.Modifiers == ModifierKeys.None)
                this.SuaDoiNSD();
            if (Keyboard.IsKeyDown(Key.F5) && Keyboard.Modifiers == ModifierKeys.None)
                this.btnChangePass_Click((object)this.btnChangePass, (RoutedEventArgs)null);
            if (!Keyboard.IsKeyDown(Key.F8) || Keyboard.Modifiers != ModifierKeys.None)
                return;
            this.XoaNSD();
        }

        private void Reload()
        {
            SqlCommand sqlcmd = new SqlCommand(StartUp.strCmdUserInfo);
            sqlcmd.CommandType = CommandType.Text;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            if (dataSet == null || dataSet.Tables.Count == 0)
                return;
            int index = 0;
            if (this.GrdUserInfo.ActiveRecord != null)
                index = this.GrdUserInfo.ActiveRecord.Index;
            StartUp.dt = dataSet.Tables[0];
            this.GrdUserInfo.DataSource = (IEnumerable)null;
            this.GrdUserInfo.DataSource = (IEnumerable)StartUp.dt.DefaultView;
            if (this.GrdUserInfo.Records.Count <= 0)
                return;
            if (index >= this.GrdUserInfo.Records.Count)
                this.GrdUserInfo.ActiveRecord = this.GrdUserInfo.Records[this.GrdUserInfo.Records.Count - 1];
            else
                this.GrdUserInfo.ActiveRecord = this.GrdUserInfo.Records[index];
        }

        private void GrdUserInfo_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            if (this.GrdUserInfo.Records.Count == 0 || this.GrdUserInfo.ActiveRecord == null)
                return;
            this.btnPhanQuyen.IsEnabled = true;
            int _user_id = (int)Convert.ToInt16((this.GrdUserInfo.ActiveRecord as DataRecord).Cells["user_id"].Value);
            if (_user_id == StartUp.user_id)
                this.btnPhanQuyen.IsEnabled = false;
            if (this.btnPhanQuyen.IsEnabled && StartUp.dt.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(_u => _u["is_admin"].ToString() == "1" && (int)Convert.ToInt16(_u["user_id"]) == _user_id)).Select<DataRow, object>((Func<DataRow, object>)(_u => _u["user_id"])).Count<object>() > 0)
                this.btnPhanQuyen.IsEnabled = false;
            DataRow[] dataRowArray = StartUp.dt.Select("user_id = " + (object)StartUp.user_id);
            if (dataRowArray.Length > 0)
            {
                if (dataRowArray[0]["r_add"].ToString().Contains(StartupBase.Menu_Id.Trim()))
                    this.btnThemNSD.IsEnabled = true;
                else
                    this.btnThemNSD.IsEnabled = false;
                if (dataRowArray[0]["r_edit"].ToString().Contains(StartupBase.Menu_Id.Trim()))
                    this.btnSuaDoiNSD.IsEnabled = true;
                else
                    this.btnSuaDoiNSD.IsEnabled = false;
                if (dataRowArray[0]["r_del"].ToString().Contains(StartupBase.Menu_Id.Trim()) && _user_id != StartUp.user_id)
                    this.btnXoaNSD.IsEnabled = true;
                else
                    this.btnXoaNSD.IsEnabled = false;
            }
            else
            {
                this.btnThemNSD.IsEnabled = true;
                this.btnSuaDoiNSD.IsEnabled = true;
                this.btnXoaNSD.IsEnabled = true;
            }
            if (!StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Contains("SE") || !(StartUp.dtRegInfo != null && StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Contains("SI")))
            {
                this.btnThemNSD.IsEnabled = false;
                this.btnXoaNSD.IsEnabled = false;
            }
            else
                if (StartUp.is_admin == 1)
                {
                this.btnThemNSD.IsEnabled = true;
                this.btnSuaDoiNSD.IsEnabled = true;
                this.btnXoaNSD.IsEnabled = true;
                this.btnPhanQuyen.IsEnabled = true;
            }    
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.is_admin == 0)
            {
                int num = (int)ExMessageBox.Show(680, StartupBase.SasObj, "Bạn không có quyền quản trị!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }    

            string directoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            SaveFileDialog saveFileDialog2 = saveFileDialog1;
            string str1 = this.Title.Trim().Replace(':', '_').Replace('/', '_');
            DateTime dateTime = DateTime.Now;
            dateTime = dateTime.Date;
            string str2 = dateTime.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern);
            string str3 = str1 + " (" + str2 + ")";
            saveFileDialog2.FileName = str3;
            saveFileDialog1.Filter = "Excel 2003 and before format|*.xls";
            saveFileDialog1.Title = "Save an file";
            if (!saveFileDialog1.ValidateNames)
                saveFileDialog1.FileName = "";
            bool? nullable = saveFileDialog1.ShowDialog();
            if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0 && !string.IsNullOrEmpty(saveFileDialog1.FileName))
                new GridViewExporter().Export((XamDataGrid)this.GrdUserInfo, saveFileDialog1.FileName, false, this.BindingSasObj);
            Directory.SetCurrentDirectory(directoryName);
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.is_admin == 0)
            {
                int num = (int)ExMessageBox.Show(680, StartupBase.SasObj, "Bạn không có quyền quản trị!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "xlsx (*.xlsx)|*.xlsx|xls (*.xls)|*.xls";
                string filepath = "";
                if (openFileDialog.ShowDialog() == true)
                {
                    filepath = openFileDialog.FileName;
                    if (File.Exists(filepath))
                    {
                        Selectdata(filepath);
                        //LoadExc(filepath);
                    }
                    else
                    {
                        string message = "Không tìm thấy file tại đường dẫn: " + filepath;                        
                        int num = (int)ExMessageBox.Show(670, StartupBase.SasObj, message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                }
            }
            catch(Exception ex)
            {
                string message = "Cấu trúc file Excel chưa đúng!";
                int num = (int)ExMessageBox.Show(730, StartupBase.SasObj, message, StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }
        void Selectdata(string filepath)
        {
            //string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq="+path+ ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
            string strConnString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" + filepath + ";Extensions=xls/xlsx;Persist Security Info=False";
            OdbcConnection oConn = new OdbcConnection(strConnString);
            List<string> lsodbc = GetSystemDriverList();
            bool hasdriver = true;
            foreach (string driver in lsodbc)
            {
                if (driver.Equals("Microsoft Excel Driver (.xls, .xlsx, .xlsm, .xlsb)"))
                {
                    hasdriver = true;
                }
            }
            if (!hasdriver)
            {
                ExMessageBox.Show(740, StartupBase.SasObj, "ODBC driver not found.", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }

            try
            {                
                oConn.Open();
                DataTable tblSchema = oConn.GetSchema("Tables");
                string sheetName = tblSchema.Rows[0]["TABLE_NAME"].ToString().Trim();
                tblSchema = null;

                string sql = "Select * From [" + sheetName.Trim() + "]";
                OdbcCommand oComm = new OdbcCommand(sql, oConn);

                DataSet ds = new DataSet();
                OdbcDataAdapter oAdapter = new OdbcDataAdapter(oComm);
                oAdapter.Fill(ds);
                ProcessData(ds.Tables[0]);
            }
            catch (IOException caught) {
                string message = "Cấu trúc file Excel chưa đúng!";
                ExMessageBox.Show(730, StartupBase.SasObj, message, StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            catch (OdbcException caught) {
                string message = "Cấu trúc file Excel chưa đúng!";
                ExMessageBox.Show(730, StartupBase.SasObj, message, StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            finally
            {
                oConn.Close();
            }            
        }

        bool ProcessData(DataTable tblexcel)
        {
            // Kiểm tra dư liệu Excel
            string strvalidname = "";
            string strnameexist = "";
            string strnamekthua = "";
            string strnamethieu = "";
            bool flag1 = true;

            for (int i=0; i< tblexcel.Rows.Count; i++)
            {
                string username = tblexcel.Rows[i]["Mã NSD"].ToString().Trim();
                if (!String.IsNullOrEmpty(username))
                {
                    string str = SysFunc.CheckInValidCode(StartupBase.SasObj, username);
                    if (str != "")
                    {
                        strvalidname += (strvalidname==""?"":";") + username;
                        flag1 = false;
                    }
                    else
                    {

                        // Ktra mã đã có
                        SqlCommand sqlcmd = new SqlCommand("select * from " + StartUp.sqlTableName + " where user_name=@user_name");
                        sqlcmd.Parameters.Add("@user_name", SqlDbType.Char).Value = (object)username;
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                        {
                            strnameexist += (strnameexist==""?"":";") + username;
                            flag1 = false;
                        }
                    }
                    // Ktra quyền thừa hưởng
                    string qkthua = tblexcel.Rows[i]["Quyền thừa hưởng"].ToString().Trim();

                    if (String.IsNullOrEmpty(qkthua))
                    { }
                    else
                    {
                        SqlCommand sqlcmd1 = new SqlCommand("select * from " + StartUp.sqlTableName + " where RTRIM(user_name)=RTRIM(@user_name)");
                        sqlcmd1.Parameters.Add("@user_name", SqlDbType.Char).Value = (object)qkthua;
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd1).Tables[0].Rows.Count <= 0)
                        {
                            strnamekthua += (strnamekthua==""?"":";") + qkthua;
                            flag1 = false;
                        }
                    }

                    if (String.IsNullOrEmpty(tblexcel.Rows[i]["Tên NSD"].ToString().Trim()))
                    {
                        strnamethieu += (strnamethieu==""?"":";") + qkthua;
                        flag1 = false;
                    }
                }
            }
            if (!flag1)
            {
                if (!String.IsNullOrEmpty(strvalidname))
                    ExMessageBox.Show(690, StartupBase.SasObj, "Tên chứa ký tự đặc biệt : [" + strvalidname + "]!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (!String.IsNullOrEmpty(strnameexist))
                    ExMessageBox.Show(700, StartupBase.SasObj, "Tên đã có : [" + strnameexist + "]!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (!String.IsNullOrEmpty(strnamekthua))
                    ExMessageBox.Show(710, StartupBase.SasObj, "Quyền thừa hưởng sai : [" + strnamekthua + "]!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (!String.IsNullOrEmpty(strnamethieu))
                    ExMessageBox.Show(710, StartupBase.SasObj, "Các mã bị sai tên : [" + strnamethieu + "]!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return false;
            }

            StartUp.currActionTask = ActionTask.Add;
            DataTable newDataTable;
            newDataTable = StartUp.GetRow(StartUp.sqlTableName);
            for (int i = 0; i < tblexcel.Rows.Count; i++)
            {
                string username = tblexcel.Rows[i]["Mã NSD"].ToString().Trim();
                if (!String.IsNullOrEmpty(username))
                {
                    DataRow nrow = newDataTable.NewRow();
                    nrow["password"] = (object)this.Encrypt("");
                    nrow["user_name"] = (object)username;
                    nrow["comment"] = (object)tblexcel.Rows[i]["Tên NSD"].ToString().Trim();
                    nrow["ma_cv"] = (object)tblexcel.Rows[i]["Chức vụ"].ToString().Trim();
                    nrow["ma_pb"] = (object)tblexcel.Rows[i]["Phòng ban"].ToString().Trim();
                    string isadmin = tblexcel.Rows[i]["Quyền quản lý"].ToString().Trim();
                    nrow["is_admin"] = (object)(isadmin.Contains("1")?1:0);
                    nrow["admin_root"] = (object)0;

                    nrow["ma_dvcs"] = (object)tblexcel.Rows[i]["Mã ĐVCS"].ToString().Trim();
                    nrow["e_mail"] = (object)tblexcel.Rows[i]["Thư điện tử"].ToString().Trim();                    
                    nrow["chuc_vu"] = (object)tblexcel.Rows[i]["Vị trí công việc"].ToString().Trim();
                    nrow["dien_thoai"] = (object)tblexcel.Rows[i]["Điện thoại"].ToString().Trim();

                    string userinh = tblexcel.Rows[i]["Quyền thừa hưởng"].ToString().Trim();
                    
                    if (isadmin.Contains("1"))
                    {
                        SqlCommand sqlcmd = new SqlCommand("select rights, r_read, r_add, r_edit, r_del, r_fav from " + StartUp.sqlTableName + " where admin_root = 1");
                        DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                        if (dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                        {
                            nrow["rights"] = (object)dataSet.Tables[0].Rows[0]["rights"].ToString();
                            nrow["r_read"] = (object)dataSet.Tables[0].Rows[0]["r_read"].ToString();
                            nrow["r_add"] = (object)dataSet.Tables[0].Rows[0]["r_add"].ToString();
                            nrow["r_edit"] = (object)dataSet.Tables[0].Rows[0]["r_edit"].ToString();
                            nrow["r_del"] = (object)dataSet.Tables[0].Rows[0]["r_del"].ToString();
                            nrow["r_fav"] = (object)dataSet.Tables[0].Rows[0]["r_fav"].ToString();
                        }
                    }
                    if (!string.IsNullOrEmpty(userinh))
                    {
                        SqlCommand sqlcmd = new SqlCommand("select rights, r_read, r_add, r_edit, r_del,r_fav from " + StartUp.sqlTableName + " where user_name = @user_name");
                        sqlcmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.NVarChar)).Value = (object)userinh;
                        DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                        if (dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                        {
                            nrow["rights"] = (object)dataSet.Tables[0].Rows[0]["rights"].ToString();
                            nrow["r_read"] = (object)dataSet.Tables[0].Rows[0]["r_read"].ToString();
                            nrow["r_add"] = (object)dataSet.Tables[0].Rows[0]["r_add"].ToString();
                            nrow["r_edit"] = (object)dataSet.Tables[0].Rows[0]["r_edit"].ToString();
                            nrow["r_del"] = (object)dataSet.Tables[0].Rows[0]["r_del"].ToString();
                            nrow["r_fav"] = (object)dataSet.Tables[0].Rows[0]["r_fav"].ToString();
                            nrow["user_name2"] = (object)userinh;
                        }
                    }
                    newDataTable.Rows.Add(nrow.ItemArray);
                }
            }
            if (newDataTable.Columns.Contains("search"))
                SetStrSearch(StartUp.sqlTableName, ref newDataTable);

            for (int i = 0; i < newDataTable.Rows.Count; i++)
            {
                if (ListFunc.inserRowInDataBase(StartUp.sqlTableName, newDataTable.Rows[i], StartupBase.SasObj, "user_id") == 1)
                {
                    newDataTable.Rows[i].SetField<int>("user_id", ListFunc.GetValueIdentityCurrent(StartUp.sqlTableName, StartupBase.SasObj));
                    StartUp.dt.Rows.Add(newDataTable.Rows[i].ItemArray);
                }    
            }

            StartUp.currActionTask = ActionTask.None;
            MessageBox.Show("Imported : " + (tblexcel.Rows.Count).ToString().Trim());
            return true;
        }

        public static void SetStrSearch(string table_name, ref DataTable tb)
        {
            string[] strArray = new string[0];
            string _str = "";
            string str1 = "";
            try
            {
                DataTable dataTable = ((IEnumerable<DataRow>)StartupBase.SasObj.DmdmInfo.Select("table_name = '" + table_name.Trim() + "'")).CopyToDataTable<DataRow>();
                if (tb.Columns.Contains("search") && dataTable.Rows.Count > 0 && tb.Rows.Count > 0)
                {
                    foreach (DataRow row in (InternalDataCollectionBase)dataTable.Rows)
                        str1 = str1 + row["field_search"].ToString().Trim() + ";";
                    string[] array = ((IEnumerable<string>)str1.Split(';')).Distinct<string>().ToArray<string>();
                    for (int irow = 0; irow <= tb.Rows.Count; irow++)
                    {
                        _str = "";
                        for (int index = 0; index < array.Length; ++index)
                        {
                            if (tb.Columns.Contains(array[index].Trim()))
                                _str = _str + SysFunc.Cat_Dau(tb.Rows[irow][array[index].Trim()].ToString().Trim()) + " ";
                        }
                        string str2 = SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, _str);
                        while (str2.IndexOf("  ") >= 0)
                            str2 = str2.Replace("  ", " ");
                        tb.Rows[irow]["search"] = (object)str2.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private string Encrypt(string password)
        {
            password = password.ToLower();
            return BitConverter.ToString(((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(new UnicodeEncoding().GetBytes(password)));
        }
        private void LoadExc(string filepath)
        {
            try
            {
                Excel.Application app = new Excel.Application();

                /*Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(@"C:\TMP\DSNSD.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);*/
                //string path = @"C:\TMP\DSNSD.xls";
                Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(filepath);
                Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)(wb.ActiveSheet);//(Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets["Sheet1"];

                Microsoft.Office.Interop.Excel.Range excelRange = sheet.UsedRange;
                string ma_nvi = (string)((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[1, 1]).Text;
                string ma_nvi2 = (string)((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[8, 2]).Text;
                bool isContinue = true;
                int i = 2;
                DataTable tblexcel = CreateUsertbl();
                while (isContinue)
                {
                    string maNSD = (string)((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 1]).Text;
                    string tenNSD = (string)((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 2]).Text;
                    if (!String.IsNullOrEmpty(maNSD) && !String.IsNullOrEmpty(tenNSD))
                    {
                        DataRow row = tblexcel.NewRow();
                        row["Mã NSD"] = (object)(maNSD);
                        row["Tên NSD"] = (object)(tenNSD);
                        row["Quyền quản lý"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 3]).Text.ToString().Contains("1") ? "1" : "0");
                        row["Quyền thừa hưởng"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 4]).Text.ToString());
                        row["Mã ĐVCS"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 5]).Text.ToString());
                        row["Thư điện tử"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 6]).Text.ToString());
                        row["Vị trí công việc"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 7]).Text.ToString());
                        row["Điện thoại"] = (object)(((Microsoft.Office.Interop.Excel.Range)excelRange.Cells[i, 8]).Text.ToString());
                        tblexcel.Rows.Add(row);
                        tblexcel.AcceptChanges();
                        i++;
                    }
                    else
                    {
                        isContinue = false;
                    }
                }
                app.Workbooks.Close();
                ProcessData(tblexcel);
            }
            catch (IOException caught)
            {
                string message = "Cấu trúc file Excel chưa đúng!" + caught.Message;
                ExMessageBox.Show(730, StartupBase.SasObj, message, StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            catch (OdbcException caught)
            {
                string message = "Cấu trúc file Excel chưa đúng!";
                ExMessageBox.Show(730, StartupBase.SasObj, message, StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            finally
            {

            }
        }
        public DataTable CreateUsertbl()
        {
            DataTable dt = new DataTable();
            DataColumn Name = new DataColumn("Mã NSD", typeof(string));
            DataColumn Ten = new DataColumn("Tên NSD", typeof(string));
            DataColumn Qql = new DataColumn("Quyền quản lý", typeof(string));
            DataColumn Qth = new DataColumn("Quyền thừa hưởng", typeof(string));
            DataColumn dvcs = new DataColumn("Mã ĐVCS", typeof(string));
            DataColumn mail = new DataColumn("Thư điện tử", typeof(string));
            DataColumn nhom = new DataColumn("Vị trí công việc", typeof(string));
            DataColumn phone = new DataColumn("Điện thoại", typeof(string));

            dt.Columns.Add(Name);
            dt.Columns.Add(Ten);
            dt.Columns.Add(Qql);
            dt.Columns.Add(Qth);
            dt.Columns.Add(dvcs);
            dt.Columns.Add(mail);
            dt.Columns.Add(nhom);
            dt.Columns.Add(phone);
            return dt;
        }

        public static List<String> GetSystemDriverList()
        {
            List<string> names = new List<string>();
            // get system dsn's
            Microsoft.Win32.RegistryKey reg = (Microsoft.Win32.Registry.LocalMachine).OpenSubKey("Software");
            if (reg != null)
            {
                reg = reg.OpenSubKey("ODBC");
                if (reg != null)
                {
                    reg = reg.OpenSubKey("ODBCINST.INI");
                    if (reg != null)
                    {

                        reg = reg.OpenSubKey("ODBC Drivers");
                        if (reg != null)
                        {
                            // Get all DSN entries defined in DSN_LOC_IN_REGISTRY.
                            foreach (string sName in reg.GetValueNames())
                            {
                                names.Add(sName);
                            }
                        }
                        try
                        {
                            reg.Close();
                        }
                        catch
                        { /* ignore this exception if we couldn't close */ }
                    }
                    }
                }

                return names;
            }

        private void btnPhanQuyenW_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            FrmPermissionW frmPermission = new FrmPermissionW();
            frmPermission.rowIndex = StartUp.dt.Rows.IndexOf(((this.GrdUserInfo.ActiveRecord as DataRecord).DataItem as DataRowView).Row);
            frmPermission.Is_mobile = "2";
            frmPermission.ShowDialog();
            this.Reload();
            frmPermission.Closed += new EventHandler(this._formPQ_Closed);
        }

        private void btnPhanQuyenA_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.currActionTask != ActionTask.None || this.GrdUserInfo.ActiveRecord == null)
                return;
            StartUp.currSqlTableKey = (this.GrdUserInfo.ActiveRecord as DataRecord).Cells[StartUp.SqlTableKey].Value.ToString();
            FrmPermissionA frmPermission = new FrmPermissionA();
            frmPermission.rowIndex = StartUp.dt.Rows.IndexOf(((this.GrdUserInfo.ActiveRecord as DataRecord).DataItem as DataRowView).Row);
            frmPermission.Is_mobile = "1";
            frmPermission.ShowDialog();
            this.Reload();
            frmPermission.Closed += new EventHandler(this._formPQ_Closed);
        }
    }
}
