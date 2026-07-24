using Infragistics.Windows.Controls;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using SasControls;
using SasControls.ControlLib;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using SasLib;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Infragistics.Windows.Themes;

namespace QCCNCC
{
    public partial class FrmPoctpna : FormTrans
    {
        private bool txtDiaChiFocusable = true;

        public static int iRow = 0;
        public static int OldiRow = 0;
        public string Old_ma_kho = string.Empty;
        private string ma_hd;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Lan0;

        private DataSet DsVitual;
        private DataSet dsCheckData;
        public FrmPoctpna()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_So_ct = this.txtSo_ct;
        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.BindingSasObj = StartupBase.SasObj;
                StartUp.M_AR_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_CK"));
                StartUp.M_AR_TT = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_TT"));
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");

                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmPoctpna.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCt.Lan = StartUpTrans.M_LAN;

                this.LanguageProvider.Language = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.LoadData();
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                }

                if (StartUp.M_AR_CK == 0)
                {
                    this.GrdLayout00.RowDefinitions[3].Height = new GridLength(114.0);
                }
                this.txtMa_qs.Filter = "ma_cts like '"+StartUp.Ma_ct+"%' and status = 1";
                this.SetFocusToolbar();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void LoadData()
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout10.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout30.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            //this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;          
        }

        private void V_Dau()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Truoc()
        {
            if (FrmPoctpna.iRow <= 1)
                return;
            --FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Sau()
        {
            if (FrmPoctpna.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Cuoi()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Moi()
        {
            try
            {
                string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
                FormTrans.currActionTask = ActionTask.Add;
                if (string.IsNullOrEmpty(str))
                    return;

                this.DsVitual = StartUpTrans.DsTrans.Copy();
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)str;
                row["loai_phieu"] = (object)"PNA";
                row["loai_yc"] = (object)"IQC";
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["ngay_ct"] = DateTime.Now.Date;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ma_qs"].ToString().Trim());
                }
                row["status"] = (object)"1";

                StartUpTrans.DsTrans.Tables[0].Rows.Add(row);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.NewRowCt();
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmPoctpna.OldiRow = FrmPoctpna.iRow;
                FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            FrmPoctpnaCopy frmPoctpnaCopy = new FrmPoctpnaCopy();
            frmPoctpnaCopy.Closed += new EventHandler(this._formcopy_Closed);
            frmPoctpnaCopy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!FrmPoctpnaCopy.isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                //this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.TxtPhieuhang.Focus()));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow].ItemArray;
                row1["stt_rec"] = (object)str;
                row1["ngay_ct"] = (object)FrmPoctpnaCopy.ngay_ct;
                row1["ngay_lct"] = (object)FrmPoctpnaCopy.ngay_ct;
                row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                {
                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"))
                    {
                        DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                        row2.ItemArray = dataRow.ItemArray;
                        row2["stt_rec"] = (object)str;
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                    }
                }
                FrmPoctpna.OldiRow = FrmPoctpna.iRow;
                FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmPoctpna.IsInEditMode.Value = true;
            }
        }
        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (this.txtMa_kh.RowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim())))
                return;
            if (this.M_LAN.ToUpper().Equals("V"))
            {
                this.txtTen_kh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
                StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ten_kh2"] = (object)this.txtMa_kh.RowResult["ten_kh2"].ToString();
            }
            else
            {
                this.txtTen_kh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
                StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ten_kh"] = (object)this.txtMa_kh.RowResult["ten_kh"].ToString();
            }
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            if (string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_kh.RowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            //StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"] = (object)this.txtMa_kh.RowResult["ma_thck"].ToString().Trim();
            this.txtMaSoThue.Text = this.txtMa_kh.RowResult["ma_so_thue"].ToString();
        }
        private void V_Sua()
        {
            this.ma_hd = this.txtSo_ct.Text.Trim();
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num1 = (int)ExMessageBox.Show(2215, StartupBase.SasObj, "Không có dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)))
            {
                int num2 = (int)ExMessageBox.Show(2220, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                // this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.TxtYeucau.Focus()));
                FormTrans.currActionTask = ActionTask.Edit;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                FrmPoctpna.IsInEditMode.Value = true;
            }
        }

        private void V_Huy()
        {
            FrmPoctpna.IsInEditMode.Value = false;
            if (this.DsVitual == null || StartUpTrans.DsTrans.Tables[0].Rows.Count <= 0)
                return;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmPoctpna.iRow = FrmPoctpna.OldiRow;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString());
                        break;
                    }
                    break;
                case ActionTask.Edit:
                    FormTrans.currActionTask = ActionTask.View;
                    string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                    }
                    StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmPoctpna.iRow);
                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row1, FrmPoctpna.iRow);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    break;
            }
        }
        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            if (Keyboard.IsKeyDown(Key.Tab) && Keyboard.Modifiers == ModifierKeys.Shift)
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Shift, Key.Tab);
            else
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }
        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUp.DeleteVoucher(_stt_rec);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec='" + _stt_rec + "'");
                StartUpTrans.DsTrans.Tables[0].Rows.Remove(dataRowArray[0]);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    FrmPoctpna.iRow = FrmPoctpna.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmPoctpna.iRow - 1 : FrmPoctpna.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString());
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            FormTrans.currActionTask = ActionTask.None;
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmTim frmTim = new FrmTim(StartupBase.SasObj, StartUpTrans.filterId, StartUp.tableList);
                SysFunc.LoadIcon((Window)frmTim);
                frmTim.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Xem()
        {
            FormTrans.currActionTask = ActionTask.View;
            if (StartUp.M_AR_CK == 0 && !StartUp.HiddenFieldIsSetted)
            {
                StartUp.stringBrowse1 = StartUp.EditCkFields(StartUp.stringBrowse1);
                StartUp.stringBrowse2 = StartUp.EditCkFields(StartUp.stringBrowse2);
            }
            StartUp.HiddenFieldIsSetted = true;
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.ListFieldSum = "t_so_luong;t_so_luong1";
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "QCCNCC_4";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string str = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmPoctpna.iRow = index + 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_In()
        {
            new FrmIn().ShowDialog();
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {

        }

        private void NewRowCt()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                int result = 0;
                int num1 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (result >= num1 ? result : num1) + 1;
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num2);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                dataRow["so_luong"] = dataRow["so_luong1"] = (object)0;

                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCt();
            return true;
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmPoctpna.IsInEditMode.Value)
                    return;
                if (this.GrdCt.ActiveCell != null && StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "ma_vt_i":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null)
                            {
                                e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl1.RowResult["ten_vt"];
                                e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl1.RowResult["ten_vt2"];
                            }
                            else
                            {
                                e.Cell.Record.Cells["ten_vt"].Value = "";
                                e.Cell.Record.Cells["ten_vt2"].Value = "";
                            }
                            break;
                        case "so_luong":
                            try
                            {
                                this.Sum_ALL();
                                break;
                            }
                            catch (Exception ex)
                            {
                                ErrorLog.CatchMessage(ex);
                            }
                            break;
                        case "so_luong1":
                            try
                            {
                                this.Sum_ALL();
                                break;
                            }
                            catch (Exception ex)
                            {
                                ErrorLog.CatchMessage(ex);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }       
      
        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Sum_ALL();
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt_i"].Value == null || activeRecord.Cells["ma_vt_i"].Value.ToString() == "")
                        break;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt_i"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCt, "ma_vt_i");
                            break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(2225, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord1))
                        break;
                    int num1 = 0;
                    Cell activeCell = this.GrdCt.ActiveCell;
                    if (activeRecord1.Index == 0)
                    {
                        if (this.GrdCt.Records.Count == 1)
                            this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord1.Index == this.GrdCt.Records.Count - 1)
                        num1 = activeRecord1.Index - 1;
                    int num2 = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[activeRecord1.Index].Row);
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[num1 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : num1];
                        this.Sum_ALL();
                    }
                    break;
            }
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt_i"].Value == null || activeRecord.Cells["ma_vt_i"].Value.ToString() == ""))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            this.GrdCt.ActiveCell = (this.GrdCt.Records[this.GrdCt.Records.Count - 1] as DataRecord).Cells["ma_vt_i"];
        }
        private void V_Nhan()
        {
            try
            {
                StartUpTrans.DsTrans.Tables[0].AcceptChanges();  
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                bool flag = false;
                if (!this.IsSequenceSave)
                {
                    Decimal result1 = new Decimal(0);
                    Decimal result2 = new Decimal(0);
                    Decimal result3 = new Decimal(0);
                    Decimal result4 = new Decimal(0);

                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    //else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_ph"].ToString()))
                    //{
                    //    int num = (int)ExMessageBox.Show(2220, StartupBase.SasObj, "Chưa vào phiếu hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    //    this.TxtPhieuhang.Focus();
                    //    flag = true;
                    //}
                    else
                    {

                        int num1;
                        if (StartUp.M_NGAY_BAT_DAU.HasValue)
                        {
                            if (this.txtNgay_ct.IsValueValid)
                            {
                                DateTime dValue1 = this.txtNgay_ct.dValue;
                                DateTime? nullable = StartUp.M_NGAY_BAT_DAU;
                                if ((nullable.HasValue ? (dValue1 < nullable.GetValueOrDefault() ? 1 : 0) : 0) == 0)
                                {
                                    DateTime dValue2 = this.txtNgay_ct.dValue;
                                    nullable = StartUp.M_NGAY_KET_THUC;
                                    num1 = (nullable.HasValue ? (dValue2 > nullable.GetValueOrDefault() ? 1 : 0) : 0) == 0 ? 1 : 0;
                                    goto label_14;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_14:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            this.txtNgay_ct.Focus();
                        }
                        else if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt_i"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(2250, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCt.Focus();
                            flag = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(2262, StartupBase.SasObj, "Chưa vào số c.từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag = true;
                        }
                    }
                    if (!flag && StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    {
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt_i"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào chi tiết, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt_i"];
                                this.GrdCt.Focus();
                                return;
                            }
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào số lượng, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["so_luong"];
                                this.GrdCt.Focus();
                                return;
                            }
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong1"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào số lượng mẫu yêu cầu, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["so_luong"];
                                this.GrdCt.Focus();
                                return;
                            }
                        }
                    }
                }
                if (!flag)
                {
                    if (!this.IsSequenceSave)
                    {
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_hd_i"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_hd"];
                        }
                        this.Sum_ALL();

                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    //StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_hd"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();

                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable1.Rows[0]["status"] = (object)1;
                    if (!DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id"))
                    {
                        int num1 = (int)ExMessageBox.Show(2125, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else
                    {
                        DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            if (!this.IsSequenceSave)
                                dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;
                            LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                        }

                        if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            int num1 = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                        else
                        {
                            if (!this.IsSequenceSave && !flag)
                            {
                                this.dsCheckData = StartUp.CheckData();
                                this.dsCheckData.Tables[0].AcceptChanges();
                                if (this.dsCheckData.Tables.Count > 0)
                                {
                                    foreach (DataRowView dataRowView in this.dsCheckData.Tables[0].DefaultView)
                                    {
                                        if (!flag)
                                        {
                                            switch (dataRowView[0].ToString())
                                            {
                                                case "PH01":
                                                    if (StartUpTrans.M_trung_so.Equals("1"))
                                                    {
                                                        if (ExMessageBox.Show(2290, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                        {
                                                            this.txtSo_ct.SelectAll();
                                                            this.txtSo_ct.Focus();
                                                            flag = true;
                                                            break;
                                                        }
                                                        break;
                                                    }
                                                    if (StartUpTrans.M_trung_so.Equals("2"))
                                                    {
                                                        int num2 = (int)ExMessageBox.Show(2295, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag = true;
                                                        break;
                                                    }
                                                    break;
                                            }
                                            this.dsCheckData.Tables[0].Rows.Remove(dataRowView.Row);
                                        }
                                        else
                                            break;
                                    }
                                }
                            }
                            if (!flag)
                            {
                                StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                new Thread((ThreadStart)(() => this.Post())).Start();
                                if (!this.IsSequenceSave)
                                {
                                    int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                    if (FrmPoctpna.iRow != pos)
                                    {
                                        DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                        DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                        row2.ItemArray = row1.ItemArray;
                                        if (FrmPoctpna.iRow > pos)
                                            StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                        else
                                            StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                        StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                        FrmPoctpna.iRow = pos;
                                    }
                                    FormTrans.currActionTask = ActionTask.View;
                                    FrmPoctpna.IsInEditMode.Value = false;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public new string GetLastSoct(SasObject SasObj, string ma_qs)
        {
            try
            {
                string cmdText = "SELECT MAX(so_ct) FROM ph50 WHERE ma_qs='" + ma_qs.Trim() + "';";
                DataRow row = SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0].Rows[0];
                if (row[0] != null && row[0] != DBNull.Value)
                    return row[0].ToString();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return "";
        }


        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                return;
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
            {

                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim()) || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"].ToString().Trim()))
                {
                    this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
                }
                else
                    this.txtSo_ct.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim();
            }
            if (this.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtSo_ct.Text, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
            {
                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
            }
        }

        private void txtSo_ct_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtSo_ct.MaxLength = ListFunc.GetLengthColumn(ListFunc.GetSqlTableFieldList(StartupBase.SasObj, "v_ph50"), "so_ct");
        }
        private void Sum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong1"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong1)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));

        }

        private bool GrdCp_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D3);
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            return false;
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = new Decimal(0);
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.IsKeyDown(Key.Return) && (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt)))
            {
                TextBox textBox = sender as TextBox;
                textBox.SelectedText = Environment.NewLine;
                ++textBox.SelectionStart;
                textBox.SelectionLength = 1;
                e.Handled = true;
            }
            else
            {
                if (!Keyboard.IsKeyDown(Key.Return))
                    return;
                (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
                e.Handled = true;
            }
        }
        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value != DBNull.Value)
                return;
            this.txtNgay_ct.Value = (object)DateTime.Now;
        }

        private void txtSo_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Edit || !(this.txtSo_ct.Text.Trim() != this.ma_hd))
                return;
            //int num = (int)ExMessageBox.Show(2377, StartupBase.SasObj, "Đơn hàng đã có phát sinh, không thể sửa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //this.txtSo_ct.Text = this.ma_hd;
        }


        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[QCCNCC-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private void txtHan_tt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Return)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.Up:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveUp((DataRecord)this.GrdCt.ActiveRecord);
                    break;
                case Key.Down:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveDown((DataRecord)this.GrdCt.ActiveRecord);
                    break;
            }
        }



        //private void txtMa_bpxl_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //    if (!string.IsNullOrEmpty(this.txtMa_bpxl.Text) && this.txtMa_bpxl.RowResult != null)
        //    {
        //        if (StartupBase.M_LAN.Equals("V"))
        //        {
        //            this.tblTenbpxl.Text = this.txtMa_bpxl.RowResult["ten_bpns"].ToString();
        //        }
        //        else
        //        {
        //            this.tblTenbpxl.Text = this.txtMa_bpxl.RowResult["ten_bpns2"].ToString();
        //        }
        //    }
        //}

        //private void txtMa_bpyc_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //    if (!string.IsNullOrEmpty(this.txtMa_bpyc.Text) && this.txtMa_bpyc.RowResult != null)
        //    {
        //        if (StartupBase.M_LAN.Equals("V"))
        //        {
        //            this.tblTenbpyc.Text = this.txtMa_bpyc.RowResult["ten_bpns"].ToString();
        //        }
        //        else
        //        {
        //            this.tblTenbpyc.Text = this.txtMa_bpyc.RowResult["ten_bpns2"].ToString();
        //        }
        //    }
        //}

        private void BtnChonphieu_Click(object sender, RoutedEventArgs e)
        {
            FrmLoc Locdonhang = new FrmLoc();
            Locdonhang.ShowDialog();
            if (StartUp.isOk)
            {
                if (StartUp.HDBData.Tables[1].DefaultView.Count > 0)
                {
                    if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"] + "'"))
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    }
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_ph"] = StartUp.HDBData.Tables[1].DefaultView[0]["so_ct"].ToString().Trim();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh_ph"] = StartUp.HDBData.Tables[0].DefaultView[0]["ma_kh"].ToString();
                    foreach (DataRow r in StartUp.HDBData.Tables[1].DefaultView.ToTable().Rows)
                    {
                        try
                        {
                            DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                            dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                            int result = 0;
                            int num1 = 0;
                            if (this.GrdCt.Records.Count > 0)
                            {
                                string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                                if (str != null)
                                    int.TryParse(str.ToString(), out result);
                            }
                            int num2 = (result >= num1 ? result : num1) + 1;
                            dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num2);
                            dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                            dataRow["ma_vt_i"] = r["ma_vt"];
                            dataRow["ten_vt"] = r["ten_vt"];
                            dataRow["dvt1"] = r["dvt1"];                          
                            dataRow["so_luong"] = dataRow["so_luong1"] = r["so_luong"];
                            dataRow["so_ct_ph"] = r["so_ct"]; 
                         
                            FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                            StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                        }
                    }
                    this.Sum_ALL();                 
                }
                //this.TxtPhieuhang.Focus();
            }
        }
    }
}
