using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasControls.ControlLib;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace COSXLSX.CODMNVL
{
    public partial class FrmCapNhat : Form
    {
        private string sTag = string.Empty;
        private string sso_lsxOld = string.Empty;
        private string sMabphtOld = string.Empty;
        private string sMaspOld = string.Empty;
        private string sMahdOld = string.Empty;
        private string sMapxOld = string.Empty;
        private string sMavtOld = string.Empty;
        private string sMakyOld = StartUp.sMa_ky_loc;
        public bool isCloseForm = false;
        private AutoCompleteTextBox[] alist = (AutoCompleteTextBox[])null;
        public bool IsOk = false;
        private DataSet dsTmp;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject M_IP_SL;
        private EditModeBindingObject FormInEditMode;
        DataTable ttc = new DataTable();

        public FrmCapNhat()
        {
            this.InitializeComponent();
            this.FormInEditMode = (EditModeBindingObject)this.FindResource("IsInEditMode");
            if (StartUp.currActionTask == ActionTask.View)
                this.GridOkCancel.ButtonType = 1;
            this.alist = new AutoCompleteTextBox[3]
            {
                this.txtMa_sp,
                this.txtSo_lsx,
                this.txtMa_bpht
            };
        }

        public FrmCapNhat(string tag, string so_lsx, string mabpht, string masp, string mahd, string mapx, string ma_vt)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.sTag = tag;
            this.sso_lsxOld = so_lsx;
            this.sMabphtOld = mabpht;
            this.sMaspOld = masp;
            this.sMahdOld = mahd;
            this.sMapxOld = mapx;
            this.sMavtOld = ma_vt;
            this.FormInEditMode = (EditModeBindingObject)this.FindResource("IsInEditMode");
            if (StartUp.currActionTask == ActionTask.View)
                this.GridOkCancel.ButtonType = 1;
            this.alist = new AutoCompleteTextBox[4]
            {
                this.txtMa_sp,
                this.txtSo_lsx,
                this.txtMa_bpht,
                this.txtMa_hd
            };
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.GrdCt.Lan = StartupBase.M_LAN;
            this.LanguageProvider.Language = StartupBase.M_LAN;
            this.dsTmp = StartUp.DataSourceReport.Copy();
            this.dsTmp.Tables[1].DefaultView.RowFilter = string.Format("tag = '{0}'", this.sTag);


            DataTable dataTable = this.dsTmp.Tables[1].Copy();
            dataTable.Clear();

            for (int index = 0; index < this.dsTmp.Tables[1].Rows.Count; ++index)
            {
                if (this.dsTmp.Tables[1].Rows[index]["tag"].ToString() == this.sTag)
                {
                    DataRow row = this.dsTmp.Tables[1].Rows[index];
                    dataTable.ImportRow(row);
                }
            }
            int x = dataTable.Rows.Count;
            for (int index = 0; index < x; ++index)
            {
                if (dataTable.Rows[index]["tag"].ToString() == this.sTag && dataTable.Rows[index]["ma_vt"].ToString() == this.sMavtOld)
                {
                    DataRow row = dataTable.Rows[index];
                    dataTable.ImportRow(row);
                }
            }
            dataTable.AcceptChanges();
            int index1 = dataTable.Rows.IndexOf(dataTable.Select("ma_vt = '" + sMavtOld + "'").FirstOrDefault());
            if (index1 >= 0)
            {
                DataRow row = dataTable.Rows[index1];
                dataTable.Rows.Remove(row);
            }


            this.dsTmp.Tables[1].Clear();
            for (int index = dataTable.Rows.Count - 1; index >= 0; --index)
            {
                this.dsTmp.Tables[1].ImportRow(dataTable.Rows[index]);
            }

            if (dsTmp.Tables[1].DefaultView.Table.Rows.Count > 0)
            {
                ttc = dsTmp.Tables[1].DefaultView[0].Row.Table.Copy();
            }
            this.gridLayout10.DataContext = ttc.DefaultView;
            //this.gridLayout20.DataContext = dataTable1.DefaultView;
            //this.GrdCt.DataSource = (IEnumerable)dataTable1.DefaultView;
            this.gridLayout20.DataContext = ttc.DefaultView;
            this.GrdCt.DataSource = (IEnumerable)this.dsTmp.Tables[1].DefaultView;
            this.Voucher_Lan0 = (CodeValueBindingObject)this.FindResource("Voucher_Lan0");
            this.Voucher_Lan0.Value = StartupBase.M_LAN.Equals("V");
            this.M_IP_SL = (CodeValueBindingObject)this.FindResource("M_IP_SL");
            //this.M_IP_SL.Text = this.BindingSasObj.GetOption("M_IP_SL").ToString();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.txtMa_bpht != null)
               {
                   if (StartUp.sMa_bpht_loc.Trim() != "")
                   {
                       if (StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
                           this.txtMa_bpht.Text = StartUp.sMa_bpht_loc.Trim();
                       this.txtMa_bpht.IsReadOnly = true;
                   }
                   this.txtMa_bpht.SearchInit();
                   if (this.txtMa_bpht.RowResult != null)
                       this.lblTen_bpht.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMa_bpht.RowResult["ten_bpht2"].ToString() : this.txtMa_bpht.RowResult["ten_bpht"].ToString();
               }
               if (this.txtSo_lsx != null)
               {
                   if (StartUp.sso_lsx_loc.Trim() != "")
                   {
                       if (StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
                           this.txtSo_lsx.Text = StartUp.sso_lsx_loc.Trim();
                       this.txtSo_lsx.IsReadOnly = true;
                   }
                   this.txtSo_lsx.SearchInit();
                   if (this.txtSo_lsx.RowResult != null)
                   {
                       this.lblTenLsx.Text = !StartupBase.M_LAN.Equals("V") ? this.txtSo_lsx.RowResult["ten_lsx2"].ToString() : this.txtSo_lsx.RowResult["ten_lsx"].ToString();
                       this.lblNgayLsx.Value = this.txtSo_lsx.RowResult["ngay_lkh"];
                   }
               }
               if (this.txtMa_hd != null)
               {
                   if (StartUp.sMa_hd_loc.Trim() != "")
                   {
                       if (StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
                           this.txtMa_hd.Text = StartUp.sMa_hd_loc.Trim();
                       this.txtMa_hd.IsReadOnly = true;
                   }
                   this.txtMa_hd.SearchInit();
                   if (this.txtMa_hd.RowResult != null)
                   {
                       this.lblNgayCt.Value = this.txtMa_hd.RowResult["ngay_ct"];
                   }
               }
               if (this.txtMa_sp == null)
                   return;
               if (StartUp.sMa_sp_loc.Trim() != "")
               {
                   if (StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
                       this.txtMa_sp.Text = StartUp.sMa_sp_loc.Trim();
                   this.txtMa_sp.IsReadOnly = true;
               }
               if (this.txtMa_Chuyen == null)
                   return;
               if (sMapxOld.Trim() != "")
               {
                   if (StartUp.currActionTask == ActionTask.Add || StartUp.currActionTask == ActionTask.Edit)
                       this.txtMa_Chuyen.Text = sMapxOld.Trim();
                   //this.txtMa_Chuyen.IsReadOnly = true;
               }
               else if ((!this.txtMa_sp.IsReadOnly || !this.txtMa_Chuyen.IsReadOnly) && StartUp.currActionTask != ActionTask.View)
                   this.txtMa_sp.IsFocus = true;
               this.txtMa_Chuyen.IsFocus = true;
               this.txtMa_sp.SearchInit();
               this.txtMa_Chuyen.SearchInit();
               this.txtMa_sp_PreviewLostFocus(this.txtMa_sp, (KeyboardFocusChangedEventArgs)null);
               this.txtMa_Chuyen_PreviewLostFocus(this.txtMa_Chuyen, (KeyboardFocusChangedEventArgs)null);
           }));
        }

        private void Form_Closed(object sender, EventArgs e)
        {
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
                if (!this.checkValid())
                    return;
                if (!StartUp.isNew)
                {
                    SqlCommand sqlcmd1 = new SqlCommand();
                    sqlcmd1.CommandText = "delete [cosxlsx-Dmdmvtct] Where ma_sp = @Masp and ma_ky = @Maky and ma_px = @Mapx; ";
                    sqlcmd1.Parameters.Add("@Masp", SqlDbType.Char).Value = this.sMaspOld;
                    sqlcmd1.Parameters.Add("@Maky", SqlDbType.Char).Value = this.sMakyOld;
                    sqlcmd1.Parameters.Add("@Mapx", SqlDbType.Char).Value = this.sMapxOld;
                    StartupBase.SasObj.ExcuteNonQuery(sqlcmd1);
                }

                for (int index = 0; index < this.dsTmp.Tables[1].DefaultView.Count; index++)
                {
                    if (StartUp.isNew)
                    {
                        dsTmp.Tables[1].DefaultView[index]["user_id0"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                        dsTmp.Tables[1].DefaultView[index]["time0"] = (object)DateTime.Now.ToString("HH:mm:ss");
                        dsTmp.Tables[1].DefaultView[index]["date0"] = (object)DateTime.Now;
                    }
                    dsTmp.Tables[1].DefaultView[index]["user_id2"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                    dsTmp.Tables[1].DefaultView[index]["time2"] = (object)DateTime.Now.ToString("HH:mm:ss");
                    dsTmp.Tables[1].DefaultView[index]["date2"] = (object)DateTime.Now;

                    dsTmp.Tables[1].DefaultView[index]["ma_sp"] = txtMa_sp.Text;
                    dsTmp.Tables[1].DefaultView[index]["ma_px"] = txtMa_Chuyen.Text;
                    ListFunc.inserRowInDataBase("[cosxlsx-Dmdmvtct]", dsTmp.Tables[1].DefaultView.Table.Rows[index], StartupBase.SasObj);
                }
                this.Close();
                this.IsOk = true;
            }
        }

        public bool checkValid()
        {
            if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return false;
            }
            if (string.IsNullOrEmpty(this.txtMa_sp.Text.Trim().ToString()))
            {
                int num = (int)ExMessageBox.Show(1490, StartupBase.SasObj, "Chưa vào mã sản phẩm!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_sp.IsFocus = true;
                return false;
            }
            //if (string.IsNullOrEmpty(this.txtSo_lsx.Text.Trim().ToString()) && string.IsNullOrEmpty(this.txtMa_bpht.Text.Trim().ToString()) && string.IsNullOrEmpty(this.txtMa_hd.Text.Trim().ToString()))
            //{
            //    int num = (int)ExMessageBox.Show(14, StartupBase.SasObj, "Chưa vào số lệnh sản xuất hoăc Chưa vào số đơn hàng hoăc Chưa vào bộ phận hoạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtSo_lsx.IsFocus = true;
            //    return false;
            //}
            //if (StartUp.sMa_sp_loc.Trim() != "" && this.txtMa_sp.Text.Trim() != StartUp.sMa_sp_loc.Trim())
            //{
            //    int num = (int)ExMessageBox.Show(1493, StartupBase.SasObj, "Mã sản phẩm khác mã sản phẩm của điều kiện lọc!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtMa_sp.IsFocus = true;
            //    return false;
            //}
            //if (StartUp.sso_lsx_loc.Trim() != "" && this.txtSo_lsx.Text.Trim() != StartUp.sso_lsx_loc.Trim())
            //{
            //    int num = (int)ExMessageBox.Show(1494, StartupBase.SasObj, "Số lệnh sản xuất khác số lệnh sản xuất của điều kiện lọc!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtSo_lsx.IsFocus = true;
            //    return false;
            //}
            //if (StartUp.sMa_bpht_loc.Trim() != "" && this.txtMa_bpht.Text.Trim() != StartUp.sMa_bpht_loc.Trim())
            //{
            //    int num = (int)ExMessageBox.Show(1495, StartupBase.SasObj, "Mã bộ phận hạch toán khác mã bộ phận hạch toán của điều kiện lọc!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtMa_bpht.IsFocus = true;
            //    return false;
            //}
            //if (StartUp.sMa_hd_loc.Trim() != "" && this.txtMa_hd.Text.Trim() != StartUp.sMa_hd_loc.Trim())
            //{
            //    int num = (int)ExMessageBox.Show(1496, StartupBase.SasObj, "Số đơn hàng khác số đơn hàng của điều kiện lọc!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txtMa_hd.IsFocus = true;
            //    return false;
            //}
            if (StartUp.isNew)
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "Select ma_ky,ma_sp From [cosxlsx-Dmdmvt] Where ma_sp = @Masp and ma_ky = @Maky and ma_px = @Mapx";
                //sqlcmd.Parameters.Add("@so_lsx", SqlDbType.Char).Value = this.txtSo_lsx.Text.Trim().ToString();
                //sqlcmd.Parameters.Add("@Mabpht", SqlDbType.Char).Value = this.txtMa_bpht.Text.Trim().ToString();
                sqlcmd.Parameters.Add("@Masp", SqlDbType.Char).Value = this.txtMa_sp.Text.Trim().ToString();
                sqlcmd.Parameters.Add("@Mapx", SqlDbType.Char).Value = this.txtMa_Chuyen.Text.Trim().ToString();
                //sqlcmd.Parameters.Add("@Mahd", SqlDbType.Char).Value = this.txtMa_hd.Text.Trim().ToString();
                sqlcmd.Parameters.Add("@Maky", SqlDbType.Char).Value = StartUp.sMa_ky_loc;
                DataTable x = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                {
                    int num3 = (int)ExMessageBox.Show(58, StartupBase.SasObj, "Mã sản phẩm trong chuyền cùng kỳ đã tồn tại!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtSo_lsx.IsFocus = true;
                    return false;
                }
            }
            DataTable dataTable = this.dsTmp.Tables[1].Copy();
            dataTable.DefaultView.RowFilter = this.dsTmp.Tables[1].DefaultView.RowFilter;
            dataTable.DefaultView.Sort = "ma_vt";
            if (dataTable.DefaultView.Count > 0)
            {
                for (int index = 0; index < dataTable.DefaultView.Count - 1; ++index)
                {
                    string str1 = dataTable.DefaultView[index]["ma_vt"].ToString().Trim();
                    string str2 = dataTable.DefaultView[index + 1]["ma_vt"].ToString().Trim();
                    if (str1.Equals(str2))
                    {
                        int num = (int)ExMessageBox.Show(1500, StartupBase.SasObj, "Vào trùng mã vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return false;
                    }
                }
                for (int index = 0; index < dataTable.DefaultView.Count; ++index)
                {
                    if (Decimal.Parse(dataTable.DefaultView[index]["sl_dm"].ToString()) == new Decimal(0))
                    {
                        int num2 = (int)ExMessageBox.Show(1520, StartupBase.SasObj, "Chưa nhập số lượng định mức kế hoạch!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["sl_dm"];
                        this.GrdCt.Focus();
                        return false;
                    }
                }
                return true;
            }
            int num4 = (int)ExMessageBox.Show(1525, StartupBase.SasObj, "Chưa vào chi tiết vật tư không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
            this.GrdCt.Focus();
            return false;
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            if (this.GrdCt.ActiveCell != null)
            {
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox txt = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        DataRecord activeRecord = this.GrdCt.ActiveRecord as DataRecord;
                        DataRowView dataItem = activeRecord.DataItem as DataRowView;
                        if (e.Editor.Value is DBNull || string.IsNullOrEmpty(e.Editor.Value.ToString().Trim()))
                        {
                            if (ExMessageBox.Show(795, StartupBase.SasObj, "Có nhập tiếp không?", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                            {
                                if (activeRecord != null)
                                {
                                    int num1 = 0;
                                    Cell activeCell = this.GrdCt.ActiveCell;
                                    num1 = activeRecord.Index;
                                    if (activeRecord.Index == 0)
                                    {
                                        if (this.GrdCt.Records.Count == 1)
                                            this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                                    }
                                    else if (activeRecord.Index == this.GrdCt.Records.Count - 1)
                                        num1 = activeRecord.Index - 1;
                                    int num2 = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                                    if (num2 >= 0)
                                    {
                                        this.dsTmp.Tables[1].Rows.Remove(this.dsTmp.Tables[1].DefaultView[activeRecord.Index].Row);
                                        this.dsTmp.Tables[1].AcceptChanges();
                                        if (this.GrdCt.Records.Count > 0)
                                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                                    }
                                }
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridOkCancel.pnlButton.btnOk.Focus()));
                                break;
                            }
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                           {
                               txt.IsShowGrid = true;
                               this.GrdCt.ActiveCell = e.Cell;
                               this.GrdCt.ExecuteCommand(DataPresenterCommands.StartEditMode);
                           }));
                            break;
                        }
                        if (txt.RowResult != null)
                        {
                            e.Cell.Record.Cells["Ten_Vt"].Value = txt.RowResult["Ten_Vt"];
                            //e.Cell.Record.Cells["Ten_Vt2"].Value = txt.RowResult["Ten_Vt2"];
                            e.Cell.Record.Cells["Dvt"].Value = txt.RowResult["Dvt"];
                            break;
                        }
                        break;
                }
            }
        }

        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCt();
            return true;
        }

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.GridOkCancel.pnlButton.btnOk.Focus()), DispatcherPriority.Background);
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt"].Value == null || activeRecord.Cells["ma_vt"].Value.ToString() == "")
                        break;
                    this.NewRowCt();
                    this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                    this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"];
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(1530, StartupBase.SasObj, "Có chắc chắn xóa không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord1))
                        break;
                    int num = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                    int index = activeRecord1.Index;
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num >= 0)
                    {
                        this.dsTmp.Tables[1].Rows.Remove(this.dsTmp.Tables[1].DefaultView[index].Row);
                        this.dsTmp.Tables[1].AcceptChanges();
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[index > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : index];
                        else
                            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.GridOkCancel.pnlButton.btnOk.Focus()), DispatcherPriority.Background);
                    }
                    break;
            }
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl) || (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt"].Value == null || activeRecord.Cells["ma_vt"].Value.ToString() == ""))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            this.GrdCt.ActiveCell = (this.GrdCt.Records[this.GrdCt.Records.Count - 1] as DataRecord).Cells["ma_vt"];
        }

        private void txtSo_lsx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtSo_lsx != null)
            {
                if (this.txtSo_lsx.RowResult != null)
                {
                    this.lblTenLsx.Text = !StartupBase.M_LAN.Equals("V") ? this.txtSo_lsx.RowResult["ten_lsx2"].ToString() : this.txtSo_lsx.RowResult["ten_lsx"].ToString();
                    this.lblNgayLsx.Value = this.txtSo_lsx.RowResult["ngay_lkh"];
                }
                else
                {
                    this.lblTenLsx.Text = "";
                    this.lblNgayLsx.Value = (object)DBNull.Value;
                }
            }
            else
            {
                this.lblTenLsx.Text = "";
                this.lblNgayLsx.Value = (object)DBNull.Value;
            }
        }
        private void txtMa_hd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_hd != null)
            {
                if (this.txtMa_hd.RowResult != null)
                {
                    this.lblNgayCt.Value = this.txtMa_hd.RowResult["ngay_ct"];
                }
                else
                {
                    this.lblNgayCt.Value = (object)DBNull.Value;
                }
            }
            else
            {
                this.lblNgayCt.Value = (object)DBNull.Value;
            }
        }
        private void txtMa_bpht_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_bpht != null)
            {
                if (this.txtMa_bpht.RowResult != null)
                {
                    if (StartupBase.M_LAN.Equals("V"))
                        this.lblTen_bpht.Text = this.txtMa_bpht.RowResult["ten_bpht"].ToString();
                    else
                        this.lblTen_bpht.Text = this.txtMa_bpht.RowResult["ten_bpht2"].ToString();
                }
                else
                    this.lblTen_bpht.Text = "";
            }
            else
                this.lblTen_bpht.Text = "";
        }

        private void txtMa_sp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_sp != null)
            {
                if (this.txtMa_sp.RowResult == null)
                    return;
                this.lblTenSp.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMa_sp.RowResult["ten_vt2"].ToString() : this.txtMa_sp.RowResult["ten_vt"].ToString();
            }
            else
            {
                this.lblTenSp.Text = "";
            }
        }

        private void NewRowCt()
        {
            try
            {
                int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                DataRow row = this.dsTmp.Tables[1].NewRow();
                row["tag"] = this.sTag;
                row["sl_dm"] = 0;
                row["date2"] = (object)DateTime.Now;
                row["time2"] = (object)DateTime.Now.ToString("HH:mm:ss");
                row["user_id2"] = (object)num;
                row["date0"] = (object)DateTime.Now;
                row["time0"] = (object)DateTime.Now.ToString("HH:mm:ss");
                row["user_id0"] = (object)num;
                row["ma_ky"] = StartUp.sMa_ky_loc;
                this.dsTmp.Tables[1].Rows.Add(row);

            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Form_Closing(object sender, CancelEventArgs e)
        {
        }

        private void GridOkCancel_GotFocus(object sender, RoutedEventArgs e)
        {
        }

        private void pnlButton_GotFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtDenngay_GotFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtDenngay_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtTungay_GotFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtTungay_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtDoi_tuong_gt_PreviewLostKeyboardFocus(
          object sender,
          KeyboardFocusChangedEventArgs e)
        {
        }

        private void GrdCt_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.FormInEditMode.IsEditMode)
                return;
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               if (this.GrdCt.Records.Count != 0)
                   return;
               this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
           }), DispatcherPriority.Background);
        }

        private void txtMa_Chuyen_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_Chuyen != null)
            {
                if (this.txtMa_Chuyen.RowResult == null)
                    return;
                this.lblTen_Chuyen.Text = !StartupBase.M_LAN.Equals("V") ? this.txtMa_Chuyen.RowResult["ten_px"].ToString() : this.txtMa_Chuyen.RowResult["ten_px"].ToString();
            }
            else
            {
                this.lblTen_Chuyen.Text = "";
            }
        }
    }
}
