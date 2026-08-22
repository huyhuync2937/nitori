using Infragistics.Windows.DataPresenter;
using Microsoft.Win32;
using SasControls;
using SasDataLib;
using SasDefine;
using SasFormBrowes;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace SasVoucherLib
{
    public partial class NewFrm : Form
    {
        public static readonly DependencyProperty currRecordProperty = DependencyProperty.Register(nameof(currRecord), typeof(DataRecord), typeof(NewFrm), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));

        private string M_LAN = string.Empty;
        public bool IsChanged = false;
        private EditModeBindingObject FormInEditMode;
        private bool flag = false;
        private DataTable newDataTable = new DataTable();
        DataTable Oldtable = (DataTable)null;
        string old_id;
        public DataRecord currRecord
        {
            get
            {
                return (DataRecord)this.GetValue(NewFrm.currRecordProperty);
            }
            set
            {
                this.SetValue(NewFrm.currRecordProperty, (object)value);
            }
        }

        public NewFrm()
        {
            this.InitializeComponent();
            this.FormInEditMode = (EditModeBindingObject)this.FindResource((object)"IsInEditMode");
            this.BindingSasObj = StartupBase.SasObj;
            this.M_LAN = this.BindingSasObj.GetOption("M_LAN").ToString();
            this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
            {
                Source = (object)this.FormInEditMode,
                Mode = BindingMode.TwoWay
            });
            if (VoucherConextMenu.oBrowse.ActiveRecord != null && VoucherConextMenu.oBrowse.ActiveRecord.Index != -1)
                this.currRecord = VoucherConextMenu.oBrowse.ActiveRecord;
            else if (VoucherConextMenu.oBrowse.frmBrw.oBrowse.Records.Count > 0)
            {
                this.currRecord = VoucherConextMenu.oBrowse.frmBrw.oBrowse.Records[0] as DataRecord;
                VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord = VoucherConextMenu.oBrowse.frmBrw.oBrowse.Records[0];
            }
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            SysFunc.LoadIcon((Window)this);
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, VoucherConextMenu.TableName);
            this.txtTenhs.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_hs");
            this.txtTenhs2.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_hs2");
            this.txtTenfile.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "ten_file");
            try
            {
                switch (VoucherConextMenu.currActionTask)
                {
                    case ActionTask.View:

                        if (!VoucherConextMenu.IsLoadFromLookUp)
                        {
                            this.newDataTable = VoucherConextMenu.GetRow(VoucherConextMenu.TableName);
                        }
                        else
                        {
                            if (VoucherConextMenu.dmctdocData.Rows.Count > 0)
                            {
                                this.newDataTable = VoucherConextMenu.dmctdocData;
                            }
                        }
                        if (this.newDataTable.Rows.Count > 0)
                        {
                            this.newDataTable.Rows[0][VoucherConextMenu.SqlTableKey] = (object)this.newDataTable.Rows[0][VoucherConextMenu.SqlTableKey].ToString().Trim();
                        }
                        this.Btnchonfile.IsEnabled = false;
                        this.Btnphanquyen.IsEnabled = false;
                        break;
                    case ActionTask.Add:

                        if (string.IsNullOrEmpty(VoucherConextMenu.currStt_rec))
                        {
                            VoucherConextMenu.currActionTask = ActionTask.None;
                            this.Close();
                        }
                        this.newDataTable = VoucherConextMenu.GetRow(VoucherConextMenu.TableName);
                        DataRow newrow = newDataTable.NewRow();
                        this.newDataTable.Rows.Clear();
                        this.newDataTable.Rows.Add(newrow);
                        this.newDataTable.Rows[0]["stt"] = (object)GetMaxStt(VoucherConextMenu.currStt_rec);
                        this.newDataTable.Rows[0]["status"] = (object)"1";
                        this.newDataTable.Rows[0]["ngay_ct"] = (object)VoucherConextMenu.currngay_ct;
                        this.newDataTable.Rows[0]["ma_ct"] = (object)VoucherConextMenu.currMa_ct;
                        break;
                    case ActionTask.Edit:

                        if (!VoucherConextMenu.IsLoadFromLookUp)
                        {
                            if (this.currRecord != null && this.currRecord.DataItem != null && this.currRecord.RecordType == RecordType.DataRecord)
                            {
                                this.newDataTable = VoucherConextMenu.GetRow(VoucherConextMenu.TableName);
                            }
                        }
                        else
                        {
                            if (VoucherConextMenu.dmctdocData.Rows.Count > 0)
                            {
                                this.newDataTable = VoucherConextMenu.dmctdocData;
                            }
                        }
                        if (this.newDataTable.Rows.Count > 0)
                        {
                            this.newDataTable.Rows[0][VoucherConextMenu.SqlTableKey] = (object)this.newDataTable.Rows[0][VoucherConextMenu.SqlTableKey].ToString().Trim();
                            old_id = this.newDataTable.Rows[0][VoucherConextMenu.SqlTableKey].ToString().Trim();
                            filedata = (byte[])this.newDataTable.Rows[0]["file_con"];
                        }
                        break;
                }
                this._confirmGridview.DataContext = (object)this.newDataTable;
                this.txtTenhs.Focus();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }
        int GetMaxStt(string stt_rec)
        {
            int maxstt = 0;
            string sql = "Select Isnull(max(stt),0) stt from " + VoucherConextMenu.TableName + " where stt_rec = '" + stt_rec + "'";
            DataTable tblmax = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql)).Tables[0];
            if (tblmax.Rows.Count != 0)
            {
                int.TryParse(tblmax.Rows[0]["stt"].ToString(), out maxstt);
            }
            maxstt += 1;
            return maxstt;
        }
        public void EnableEditMode(bool isEditMode)
        {
            this.FormInEditMode.IsEditMode = isEditMode;
        }

        private void _confirmGridview_OnOk(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                {
                    TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                    if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                        return;
                }

                if (VoucherConextMenu.currActionTask == ActionTask.Add || VoucherConextMenu.currActionTask == ActionTask.Edit)
                {
                    if (!Checkvalidtxt())
                        return;

                    saveCustomer();
                }
                if (!VoucherConextMenu.IsLoadFromLookUp)
                {                
                    VoucherConextMenu.dmctdocData = StartupBase.SasObj.ExcuteReader(VoucherConextMenu.MCmd).Tables[0];
                    VoucherConextMenu.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)VoucherConextMenu.dmctdocData.DefaultView;
                    for (int index = 0; index < VoucherConextMenu.dmctdocData.Rows.Count; ++index)
                    {
                        DataRow _dr = VoucherConextMenu.dmctdocData.Rows[index];
                        if (_dr["stt_rec"].ToString().Trim() == VoucherConextMenu.currStt_rec)
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveDataItem = (object)_dr));
                    }
                }
                VoucherConextMenu.currActionTask = ActionTask.None;
                this.Close();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }
        private void saveCustomer()
        {
            // StartUp.LastEditRow = (DataRow)null;
            if (VoucherConextMenu.currActionTask == ActionTask.View)
                return;
            int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            this.newDataTable.AcceptChanges();
            if (VoucherConextMenu.currActionTask == ActionTask.Edit)
            {
                string sql = "Update " + VoucherConextMenu.TableName + " set ten_hs=@ten_hs,ten_hs2=@ten_hs2,ten_file=@ten_file,file_con=@file_con,user_right=@user_right,date2=@date2,time2=@time2,user_id2=@user_id2 where file_id = @file_id";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add("@file_id", SqlDbType.VarChar).Value = (object)old_id;
                cmd.Parameters.Add("@ten_hs", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_hs"].ToString();
                cmd.Parameters.Add("@ten_hs2", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_hs2"].ToString();
                cmd.Parameters.Add("@ten_file", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_file"].ToString();
                cmd.Parameters.Add("@file_con", SqlDbType.Image).Value = (object)filedata;
                cmd.Parameters.Add("@user_right", SqlDbType.Char).Value = (object)this.newDataTable.Rows[0]["user_right"].ToString();
                cmd.Parameters.Add("@date2", SqlDbType.DateTime).Value = (object)DateTime.Now;
                cmd.Parameters.Add("@time2", SqlDbType.Char).Value = (object)DateTime.Now.ToString("HH:mm:ss");
                cmd.Parameters.Add("@user_id2", SqlDbType.Int).Value = (object)num;
                StartupBase.SasObj.ExcuteNonQuery(cmd);
            }
            else
            {
                string sql = "Insert into " + VoucherConextMenu.TableName + " (file_id,stt,stt_rec,ten_hs,ten_hs2,ten_file,file_con,user_right,status,ma_ct,ngay_ct,date2,time2,user_id2,date0,time0,user_id0) values (@file_id,@stt,@stt_rec,@ten_hs,@ten_hs2,@ten_file,@file_con,@user_right,@status,@ma_ct,@ngay_ct,@date2,@time2,@user_id2,@date0,@time0,@user_id0)";
                SqlCommand cmd = new SqlCommand(sql);
                cmd.Parameters.Add("@file_id", SqlDbType.VarChar).Value = (object)DateTime.Now.Ticks;
                cmd.Parameters.Add("@stt", SqlDbType.Char).Value = (object)this.newDataTable.Rows[0]["stt"].ToString().Trim();
                cmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = (object)VoucherConextMenu.currStt_rec.Trim();
                cmd.Parameters.Add("@ten_hs", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_hs"].ToString();
                cmd.Parameters.Add("@ten_hs2", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_hs2"].ToString();
                cmd.Parameters.Add("@ten_file", SqlDbType.NVarChar).Value = (object)this.newDataTable.Rows[0]["ten_file"].ToString();
                cmd.Parameters.Add("@status", SqlDbType.Char).Value = (object)this.newDataTable.Rows[0]["status"].ToString();
                cmd.Parameters.Add("@file_con", SqlDbType.Image).Value = (object)filedata;
                cmd.Parameters.Add("@user_right", SqlDbType.Char).Value = (object)this.newDataTable.Rows[0]["user_right"].ToString();
                cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)this.newDataTable.Rows[0]["ma_ct"].ToString();
                cmd.Parameters.Add("@ngay_ct", SqlDbType.DateTime).Value =(object)Convert.ToDateTime(this.newDataTable.Rows[0]["ngay_ct"]);
                cmd.Parameters.Add("@date2", SqlDbType.DateTime).Value = (object)DateTime.Now;
                cmd.Parameters.Add("@time2", SqlDbType.Char).Value = (object)DateTime.Now.ToString("HH:mm:ss");
                cmd.Parameters.Add("@user_id2", SqlDbType.Int).Value = (object)num;
                cmd.Parameters.Add("@date0", SqlDbType.DateTime).Value = (object)DateTime.Now;
                cmd.Parameters.Add("@time0", SqlDbType.Char).Value = (object)DateTime.Now.ToString("HH:mm:ss");
                cmd.Parameters.Add("@user_id0", SqlDbType.Int).Value = (object)num;
                StartupBase.SasObj.ExcuteNonQuery(cmd);
            }
            VoucherConextMenu.currActionTask = ActionTask.None;
            // StartUp.LastEditRow = this.newDataTable.Select()[0];          
        }
        private bool Checkvalidtxt()
        {
            if (string.IsNullOrEmpty(this.txtTenhs.Text))
            {
                flag = false;
                int num = (int)ExMessageBox.Show(1907, StartupBase.SasObj, "Chưa vào tên hs!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtTenhs.Focus();
                return false;
            }
            if (filedata.Length < 1)
            {
                flag = false;
                int num = (int)ExMessageBox.Show(1907, StartupBase.SasObj, "Chưa chọn file!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.Btnchonfile.Focus();
                return false;
            }

            return true;
        }
        private void _confirmGridview_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void VoucherAttachfile_Closed(object sender, EventArgs e)
        {
            this.Close();
        }

        private void VoucherAttachfile_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Escape)
            {
                Close();
            }
        }

        public static string ByteArrayToStr(byte[] barr)
        {
            UTF8Encoding encoding = new UTF8Encoding();
            return encoding.GetString(barr, 0, barr.Length);
        }
        byte[] filedata = new byte[] { };
        byte[] FiletoByte(string filepath)
        {
            byte[] bimage;
            using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read))
            {
                bimage = new byte[fs.Length];
                fs.Read(bimage, 0, Convert.ToInt32(fs.Length));
            }
            return bimage;
        }
        private void Btnchonfile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                this.newDataTable.Rows[0]["ten_file"] = (object)openFileDialog.SafeFileName;
                filedata = FiletoByte(openFileDialog.FileName);
            }
        }

        private void Btnphanquyen_Click(object sender, RoutedEventArgs e)
        {
            if (SysFunc.CheckPermission(StartupBase.SasObj, VoucherConextMenu.currActionTask, StartupBase.Menu_Id))
            {
                DanhSachUsers dsus = new DanhSachUsers();
                dsus.struserid = this.newDataTable.Rows[0]["user_right"].ToString().Trim();
                dsus.ShowDialog();
                this.newDataTable.Rows[0]["user_right"] = (object)dsus.struserid;
            }
        }
    }
}
