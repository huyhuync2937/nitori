using Infragistics.Windows.Controls;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using InvtLib;
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
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using System.Reflection;
using System.IO;

namespace INCTPND
{
    public partial class FrmINCTPND : FormTrans
    {
        public int so_ct_px_length = 12;
        public static int iRow = 0;
        public static int OldiRow = 0;
        public string Old_ma_kho = string.Empty;
        private bool txtDiaChiFocusable = true;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject IsCheckedPn_gia_tb;
        private DataSet dsCheckData;
        private DataSet DsVitual;

        public FrmINCTPND()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.cbMa_nt;
            this.C_So_ct = this.txtSo_ct;
        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.BindingSasObj = StartupBase.SasObj;
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmINCTPND.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmINCTPND.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.IsCheckedPn_gia_tb = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedPn_gia_tb");
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmINCTPND.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                this.LanguageProvider.Language = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.IsCheckedPn_gia_tb.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["pn_gia_tb"].ToString() == "1";
                    this.UpdateTonKho();
                }
                this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
                this.SetFocusToolbar();
                //Quan ly lo
                if (StartUp.M_QL_LO_CK.ToString().Trim() != "1")
                {
                    if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_lo")))
                    {
                        this.GrdCt.FieldLayouts[0].Fields["ma_lo"].Visibility = Visibility.Collapsed;
                    }
                    if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ngay_hh")))
                    {
                        this.GrdCt.FieldLayouts[0].Fields["ngay_hh"].Visibility = Visibility.Collapsed;
                    }
                }
                if (!(StartupBase.SasObj.GetOption("M_TON_KHO13").ToString() != "1") || !this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ton13")))
                    return;
                this.GrdCt.FieldLayouts[0].Fields["ton13"].Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void LoadData()
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0 ASC";
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout10.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout20.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout21.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtStatus.IsEnabled = false;
        }

        private void V_Dau()
        {
            FrmINCTPND.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Truoc()
        {
            if (FrmINCTPND.iRow <= 1)
                return;
            --FrmINCTPND.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Sau()
        {
            if (FrmINCTPND.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmINCTPND.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Cuoi()
        {
            FrmINCTPND.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Moi()
        {
            try
            {
                string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
                FormTrans.currActionTask = ActionTask.Add;
                if (string.IsNullOrEmpty(str))
                    return;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_gd.IsFocus = true));
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)str;
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["ma_gd"] = (object)StartUp.M_ma_gd;
                row["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                row["status"] = StartUpTrans.DmctInfo["ma_post"];
                row["sua_tien"] = (object)0;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["ma_qs"].ToString().Trim());
                }
                row["ty_giaf"] = !row["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)StartUp.GetRates(row["ma_nt"].ToString().Trim(), Convert.ToDateTime(row["ngay_ct"]).Date) : (object)1;
                row["t_tien"] = (object)0;
                row["t_tien_nt"] = (object)0;
                row["t_so_luong"] = (object)0;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.IsCheckedPn_gia_tb.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["pn_gia_tb"].ToString() == "1";
                this.NewRowCt();
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.txtngay_lct.Text = "";
                FrmINCTPND.OldiRow = FrmINCTPND.iRow;
                FrmINCTPND.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmINCTPND.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
                this.ChkSuaTien.IsChecked = new bool?(false);
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
            FrmINCTPNDCopy frmInctpndCopy = new FrmINCTPNDCopy();
            frmInctpndCopy.Closed += new EventHandler(this._formcopy_Closed);
            frmInctpndCopy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!FrmINCTPNDCopy.isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_gd.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow].ItemArray;
                row1["stt_rec"] = (object)str;
                row1["ngay_ct"] = (object)FrmINCTPNDCopy.ngay_ct;
                row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];

                row1["stt_rec_px"] = "";
                row1["ma_qs_px"] = "";
                row1["so_ct_px"] = "";
                row1["ma_ct_px"] = "";
                row1["loai_xnvl"] = 0;
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
                FrmINCTPND.OldiRow = FrmINCTPND.iRow;
                FrmINCTPND.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmINCTPND.IsInEditMode.Value = true;
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
        }

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num1 = (int)ExMessageBox.Show(420, StartupBase.SasObj, "Không có dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)))
            {
                int num2 = (int)ExMessageBox.Show(425, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_gd.IsFocus = true));
                FormTrans.currActionTask = ActionTask.Edit;
                this.IsCheckedPn_gia_tb.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["pn_gia_tb"].ToString() == "1";
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                FrmINCTPND.IsInEditMode.Value = true;
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            }
        }

        private void V_Huy()
        {
            FrmINCTPND.IsInEditMode.Value = false;
            if (this.DsVitual == null || StartUpTrans.DsTrans.Tables[0].Rows.Count <= 0)
                return;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmINCTPND.iRow = FrmINCTPND.OldiRow;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString());
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
                    StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmINCTPND.iRow);
                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row1, FrmINCTPND.iRow);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    break;
            }
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUpTrans.UpdateTkSd13(1, 0);
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
                    FrmINCTPND.iRow = FrmINCTPND.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmINCTPND.iRow - 1 : FrmINCTPND.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString());
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
            string strBrowse;
            string strBrowseCt;
            if (StartUpTrans.M_LAN.Equals("V"))
            {
                strBrowse = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|')[0];
                strBrowseCt = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|')[1];
            }
            else
            {
                strBrowse = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|')[0];
                strBrowseCt = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|')[1];
            }
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, strBrowse, strBrowseCt, "stt_rec");
            formView.ListFieldSum = "t_tien_nt;t_tien";
            formView.frmBrw.Title = this.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString());
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "INCTPND_5";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int dataItemIndex = (formView.DataGrid.ActiveRecord as DataRecord).DataItemIndex;
            if (dataItemIndex >= 0)
            {
                string str = (formView.DataGrid.DataSource as DataView)[dataItemIndex]["stt_rec"].ToString();
                FrmINCTPND.iRow = dataItemIndex + 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_In()
        {
            new FrmIn().ShowDialog();
        }

        private void V_Nhan()
        {
            try
            {
                bool flag = false;
                if (!this.IsSequenceSave)
                {
                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(430, StartupBase.SasObj, "Chưa vào mã khách hàng!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag = true;
                    }
                    else if (string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                    {
                        int num = (int)ExMessageBox.Show(435, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtNgay_ct.Focus();
                        flag = true;
                    }
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
                                    goto label_15;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_15:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            this.txtNgay_ct.Focus();
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(6101, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            this.txtMa_qs.IsFocus = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(440, StartupBase.SasObj, "Chưa vào số chứng từ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag = true;
                        }
                    }
                }
                if (!flag)
                {
                    if (!this.IsSequenceSave)
                    {
                        this.Sum_ALL();
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                        {
                            for (int i = 0; i < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++i)
                            {
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["ma_vt"].ToString()))
                                {
                                    int num = (int)ExMessageBox.Show(455, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["ma_vt"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["ma_kho_i"].ToString()))
                                {
                                    int num = (int)ExMessageBox.Show(460, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["ma_kho_i"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["tk_vt"].ToString().Trim()))
                                {
                                    int num = (int)ExMessageBox.Show(465, StartupBase.SasObj, "Chưa vào tk nợ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["tk_vt"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["ma_nx_i"].ToString().Trim()))
                                {
                                    int num = (int)ExMessageBox.Show(470, StartupBase.SasObj, "Chưa vào tk có!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["ma_nx_i"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                                if (int.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["gia_ton"].ToString()) == 3 && Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["so_luong"].ToString()) == new Decimal(0))
                                {
                                    int num = (int)ExMessageBox.Show(475, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["so_luong"];
                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                    return;
                                }
                                if (StartUpTrans.DsTrans.Tables[1].DefaultView[i]["loai_vt"].ToString().Equals("51") && this.txtMa_gd.Text.Trim().Equals("4") && Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[i]["so_luong"].ToString()) == new Decimal(0))
                                {
                                    int num = (int)ExMessageBox.Show(476, StartupBase.SasObj, "Chưa có số lượng nhập kho!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                                   {
                                       this.GrdCt.ActiveCell = (this.GrdCt.Records[i] as DataRecord).Cells["so_luong"];
                                       this.GrdCt.Focus();
                                   }));
                                    return;
                                }
                            }
                        }
                        else
                        {
                            int num = (int)ExMessageBox.Show(490, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.GrdCt.Focus();
                            flag = true;
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
                            }
                            Decimal num = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                            bool? isChecked = this.ChkSuaTien.IsChecked;
                            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && num != new Decimal(0))
                                this.CanBangTien();
                            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        }
                        DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                        LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                        if (!this.IsSequenceSave)
                            LocalTable1.Rows[0]["status"] = (object)0;
                        DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                        DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            if (!this.IsSequenceSave)
                                dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;

                            if (dataRowView["dvt1"].ToString().Trim() == dataRowView["dvt"].ToString().Trim())
                            {
                                dataRowView["he_so1"] = 0;
                                dataRowView["so_luong1"] = 0;
                            }
                            //decimal resulHeso1 = 0;
                            //decimal.TryParse(dataRowView["he_so1"].ToString(), out resulHeso1);
                            //if (resulHeso1 != new decimal(0))
                            if ((decimal)dataRowView["he_so1"] != new decimal(0))
                            {
                                dataRowView["so_luong1"] = (Decimal)(Convert.ToDecimal(dataRowView["so_luong"]) * Convert.ToDecimal(dataRowView["he_so1"]));
                            }

                            LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                        }
                        if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            int num1 = (int)ExMessageBox.Show(495, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                                        if (ExMessageBox.Show(500, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
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
                                                        int num2 = (int)ExMessageBox.Show(505, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag = true;
                                                        break;
                                                    }
                                                    break;
                                                case "CT01":
                                                    int int16_1 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num3 = (int)ExMessageBox.Show(510, StartupBase.SasObj, "Tk nợ là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_1] as DataRecord).Cells["tk_vt"];
                                                    this.GrdCt.Focus();
                                                    break;
                                                case "CT02":
                                                    int int16_2 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num4 = (int)ExMessageBox.Show(515, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_2] as DataRecord).Cells["ma_nx_i"];
                                                    this.GrdCt.Focus();
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
                                string stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                new Thread((ThreadStart)(() =>
                               {
                                   this.Post();
                                   if (this.IsSequenceSave)
                                       return;
                                   this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                   {
                       if (!StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(stt_rec1))
                           return;
                       this.UpdateTonKho();
                   }));
                               })).Start();
                                if (!this.IsSequenceSave)
                                {
                                    int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                    if (FrmINCTPND.iRow != pos)
                                    {
                                        DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                        DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                        row2.ItemArray = row1.ItemArray;
                                        if (FrmINCTPND.iRow > pos)
                                            StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                        else
                                            StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                        StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                        FrmINCTPND.iRow = pos;
                                    }
                                    FormTrans.currActionTask = ActionTask.View;
                                    FrmINCTPND.IsInEditMode.Value = false;
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

        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[INCTPND-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            if (!menuItemName.Equals("btnSave"))
                this.UpdateTonKho();
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count <= 0)
                return;
            this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
        }

        private void NewRowCt()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                int result = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                    dataRow["ma_kho_i"] = (this.GrdCt.Records[0] as DataRecord).Cells["ma_kho_i"].Value;
                    dataRow["ma_nx_i"] = (this.GrdCt.Records[0] as DataRecord).Cells["ma_nx_i"].Value;
                }
                else
                {
                    dataRow["ma_kho_i"] = (object)"";
                    dataRow["ma_nx_i"] = (object)"";
                }
                int num = result + 1;
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                dataRow["so_luong"] = (object)0;
                dataRow["gia_nt"] = (object)0;
                dataRow["tien_nt"] = (object)0;
                dataRow["tien"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["ton13"] = (object)0;
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmINCTPND.IsInEditMode.Value || this.txtMa_kh.RowResult == null)
                return;
            string str1 = this.txtMa_kh.RowResult["ten_kh"].ToString().Trim();
            string str2 = this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim();
            this.tblTen_kh.Text = StartUpTrans.M_LAN.Equals("V") ? str1 : str2;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["ten_kh"] = (object)str1;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["ten_kh2"] = (object)str2;
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            if (string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_kh.RowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
        }

        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value == DBNull.Value)
                this.txtNgay_ct.Value = (object)DateTime.Now;
            if (this.txtNgay_ct.IsFocusWithin || FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy)
                return;
            if (!(this.txtNgay_ct.dValue == new DateTime()))
                ;
            if (StartUpTrans.M_ngay_lct.Equals("0") && !string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                this.txtngay_lct.Value = this.txtNgay_ct.Value;
        }

        private void txtngay_lct_GotFocus(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.M_ngay_lct.Equals("0") && !string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                this.txtngay_lct.Value = this.txtNgay_ct.Value;
            else if (string.IsNullOrEmpty(this.txtngay_lct.Text.ToString()))
                this.txtngay_lct.Value = this.txtNgay_ct.Value;
        }

        private void txtngay_lct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtngay_lct.IsFocusWithin || (FormTrans.currActionTask != ActionTask.Copy && FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit || this.txtNgay_ct.dValue.Date.Equals(this.txtngay_lct.dValue.Date)))
                return;
            int num = (int)ExMessageBox.Show(525, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmINCTPND.IsInEditMode.Value || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
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
            this.txtSo_ct.MaxLength = ListFunc.GetLengthColumn(ListFunc.GetSqlTableFieldList(StartupBase.SasObj, "v_PH71"), "so_ct");
        }

        private void cbMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null)
                return;
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (this.cbMa_nt.RowResult == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.cbMa_nt.RowResult["loai_tg"];
            if (this.cbMa_nt.RowResult["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                this.txtTy_gia.Value = (object)1;
            else
                this.txtTy_gia.Value = (object)StartUp.GetRates(this.cbMa_nt.RowResult["ma_nt"].ToString().Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrdLayoutNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.View || FormTrans.currActionTask == ActionTask.None || !(this.txtTy_gia.OldValue != this.txtTy_gia.nValue))
                return;
            this.TyGiaValueChange();
        }

        public void TyGiaValueChange()
        {
            if (this.cbMa_nt.RowResult != null)
                this.txtTy_gia.Value = this.cbMa_nt.RowResult["ma_nt"].ToString().Trim() == StartUpTrans.M_ma_nt0 ? (object)1 : this.txtTy_gia.Value;
            if (string.IsNullOrEmpty(this.txtTy_gia.Text.ToString()))
                this.txtTy_gia.Value = (object)0;
            try
            {
                if (FormTrans.currActionTask == ActionTask.Delete || !FrmINCTPND.IsInEditMode.Value || (this.txtTy_gia.Value == null || this.txtTy_gia.Value == DBNull.Value || !(this.txtTy_gia.nValue != new Decimal(0))) || (this.txtTy_gia.Value == null || this.txtTy_gia.Value == DBNull.Value || !(this.txtTy_gia.nValue != new Decimal(0)) || (this.GrdCt.Records.Count <= 0 || (this.GrdCt.DataSource as DataView).Table.DefaultView[0]["ma_vt"] == DBNull.Value)))
                    return;
                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);
                Decimal num4 = new Decimal(0);
                Decimal num5 = new Decimal(0);
                Decimal nValue = this.txtTy_gia.nValue;
                num5 = this.txtT_Tien_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_Tien_nt.Value);
                for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                {
                    if ((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt"].Value != DBNull.Value)
                    {
                        Decimal num6 = (this.GrdCt.DataSource as DataView)[index]["so_luong"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["so_luong"].Value);
                        Decimal num7 = (this.GrdCt.DataSource as DataView)[index]["gia_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["gia_nt"].Value);
                        if (num6 * num7 != new Decimal(0))
                        {
                            num2 = SysFunc.Round(num6 * num7, StartUpTrans.M_ROUND_NT);
                            (this.GrdCt.DataSource as DataView)[index]["tien_nt"] = (object)num2;
                        }
                        if (nValue * num7 != new Decimal(0))
                            (this.GrdCt.DataSource as DataView)[index]["gia"] = (object)SysFunc.Round(nValue * num7, StartUpTrans.M_ROUND_GIA);
                        if (nValue * num2 != new Decimal(0))
                            (this.GrdCt.DataSource as DataView)[index]["tien"] = (object)SysFunc.Round(nValue * num2, StartUpTrans.M_ROUND);
                    }
                }
                this.Sum_ALL();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmINCTPND.IsInEditMode.Value || (this.GrdCt.ActiveCell == null || StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) != null))
                    return;
                Decimal result1;
                Decimal result2;
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl1.RowResult["ten_vt"];
                            e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl1.RowResult["ten_vt2"];
                            e.Cell.Record.Cells["dvt"].Value = autoCompleteControl1.RowResult["dvt"];
                            e.Cell.Record.Cells["loai_vt"].Value = autoCompleteControl1.RowResult["loai_vt"];
                            (e.Cell.Record.DataItem as DataRowView)["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                            if (string.IsNullOrEmpty(e.Cell.Record.Cells["tk_vt"].Value.ToString().Trim()) || autoCompleteControl1.RowResult["sua_tk_vt"].ToString().Trim().Equals("0"))
                                e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl1.RowResult["tk_vt"];
                            if (string.IsNullOrEmpty(e.Cell.Record.Cells["ma_nx_i"].Value.ToString().Trim()))
                                e.Cell.Record.Cells["ma_nx_i"].Value = autoCompleteControl1.RowResult["tk_spdd"];
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                            if (autoCompleteControl2 != null)
                            {
                                autoCompleteControl2.SearchInit();
                                if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                                    e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl2.RowResult["tk_dl"];
                            }
                          (e.Cell.Record.DataItem as DataRowView)["sua_tk_vt"] = autoCompleteControl1.RowResult["sua_tk_vt"];
                            e.Cell.Record.Cells["gia_ton"].Value = autoCompleteControl1.RowResult["gia_ton"];
                            (e.Cell.Record.DataItem as DataRowView)["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                            if (autoCompleteControl1.RowResult["vt_ton_kho"].ToString().Equals("0"))
                            {
                                e.Cell.Record.Cells["so_luong"].Value = (object)0;
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
                                e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                e.Cell.Record.Cells["gia"].Value = (object)0;
                            }
                            if (this.ParseInt(autoCompleteControl1.RowResult["vt_ton_kho"], 0) == 1)
                            {
                                if (StartUp.M_QL_LO_CK.ToString().Trim() != "1")
                                {
                                    e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                }
                                else
                                {
                                    e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13Lo(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString(), e.Cell.Record.Cells["ma_lo"].Value.ToString());
                                }
                            }
                            else
                                e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;

                            if (autoCompleteControl1.RowResult["loai_vt"].ToString().Trim().Equals("51") && this.txtMa_gd.Text.Trim().Equals("4"))
                            {
                                if (e.Cell.Record.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_sp")))
                                    e.Cell.Record.Cells["ma_sp"].Value = e.Editor.Value;
                                else
                                    (e.Cell.Record.DataItem as DataRowView)["ma_sp"] = e.Editor.Value;
                            }


                            DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                            CellCollection cells = e.Cell.Record.Cells;
                            if (string.IsNullOrEmpty(cells["dvt1"].Value.ToString()))
                            {
                                cells["dvt1"].Value = cells["dvt"].Value;
                                AutoCompleteTextBox autoCompleteControlmavtdvt1 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["dvt1"]).Editor as ControlHostEditor);
                                if (autoCompleteControlmavtdvt1 != null)
                                {
                                    autoCompleteControlmavtdvt1.SearchInit();
                                    if (autoCompleteControlmavtdvt1.RowResult != null && (autoCompleteControlmavtdvt1.RowResult["hs_qd"] != DBNull.Value))
                                    {
                                        dataItem1["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                    }
                                    else
                                    {
                                        dataItem1["he_so1"] = (decimal)0;
                                    }
                                }
                            }
                            else
                            {
                                AutoCompleteTextBox autoCompleteControlmavtdvt1 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["dvt1"]).Editor as ControlHostEditor);
                                if (autoCompleteControlmavtdvt1 != null)
                                {
                                    autoCompleteControlmavtdvt1.SearchInit();
                                    if (autoCompleteControlmavtdvt1.RowResult == null)
                                    {
                                        cells["dvt1"].Value = cells["dvt"].Value;
                                        autoCompleteControlmavtdvt1.SearchInit();
                                        if (autoCompleteControlmavtdvt1.RowResult != null && (autoCompleteControlmavtdvt1.RowResult["hs_qd"] != DBNull.Value))
                                        {
                                            dataItem1["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                        }
                                        else
                                        {
                                            dataItem1["he_so1"] = (decimal)0;
                                        }
                                    }
                                    else
                                    {
                                        if ((autoCompleteControlmavtdvt1.RowResult["hs_qd"] != DBNull.Value))
                                            dataItem1["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                        else
                                            dataItem1["he_so1"] = (decimal)0;
                                    }
                                }
                            }
                            break;
                        }
                        break;
                    case "dvt1":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControlDvt1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        DataRowView dataItemDvt1 = e.Cell.Record.DataItem as DataRowView;
                        if (autoCompleteControlDvt1.IsDataChanged)
                        {
                            if (autoCompleteControlDvt1 != null)
                            {
                                autoCompleteControlDvt1.SearchInit();
                                if (autoCompleteControlDvt1.RowResult != null && (autoCompleteControlDvt1.RowResult["hs_qd"] != DBNull.Value))
                                {
                                    dataItemDvt1["he_so1"] = (decimal)autoCompleteControlDvt1.RowResult["hs_qd"];
                                }
                                else
                                {
                                    dataItemDvt1["he_so1"] = (decimal)0;
                                }
                            }
                            break;
                        }
                        break;
                    case "ma_sp":
                        if (e.Editor.Value == null)
                            break;
                        CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]);
                        if (cellValuePresenter != null)
                        {
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(cellValuePresenter.Editor as ControlHostEditor);
                            if (autoCompleteControl2.RowResult != null)
                            {
                                if (autoCompleteControl2.RowResult["loai_vt"].ToString().Trim().Equals("51") && this.txtMa_gd.Text.Trim().Equals("4"))
                                {
                                    e.Editor.Value = (object)autoCompleteControl2.Text.Trim();
                                    e.Editor.IsReadOnly = true;
                                }
                                else
                                    e.Editor.IsReadOnly = false;
                            }
                            break;
                        }
                        break;
                    case "ma_kho_i":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                        if (autoCompleteControl3 != null)
                        {
                            autoCompleteControl3.SearchInit();
                            if (autoCompleteControl3.RowResult != null && autoCompleteControl3.RowResult["sua_tk_vt"].ToString().Equals("0"))
                                e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl3.RowResult["tk_vt"];
                        }
                        AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl4 != null)
                        {
                            autoCompleteControl4.SearchInit();
                            if (autoCompleteControl4.RowResult != null && (autoCompleteControl4.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl4.RowResult["tk_dl"].ToString().Trim())))
                                e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl4.RowResult["tk_dl"];
                        }
                        if (autoCompleteControl3.RowResult != null)
                        {
                            if (this.ParseInt(autoCompleteControl3.RowResult["vt_ton_kho"], 0) == 1)
                            {
                                if (StartUp.M_QL_LO_CK != "1")
                                {
                                    e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                }
                                else
                                {
                                    e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13Lo(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_lo"].ToString());
                                }
                            }
                            else
                                e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;
                            break;
                        }
                        break;
                    case "ma_lo":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControlLo = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControlLo != null)
                        {
                            autoCompleteControlLo.SearchInit();
                            if (autoCompleteControlLo.RowResult != null)
                            {
                                e.Cell.Record.Cells["ngay_hh"].Value = autoCompleteControlLo.RowResult["ngay_hh"] != DBNull.Value ? autoCompleteControlLo.RowResult["ngay_hh"] : (object)DBNull.Value;
                            }
                            else
                            {
                                e.Cell.Record.Cells["ngay_hh"].Value = (object)DBNull.Value;
                                e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;
                            }
                        }
                        AutoCompleteTextBox autoCompleteControVT = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                        if (autoCompleteControVT.RowResult == null)
                            autoCompleteControVT.SearchInit();
                        if (autoCompleteControVT.RowResult != null)
                        {
                            if (this.ParseInt(autoCompleteControVT.RowResult["vt_ton_kho"], 0) == 1)
                            {
                                if (!string.IsNullOrEmpty(e.Cell.Record.Cells["ma_vt"].Value.ToString()) && !string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kho_i"].Value.ToString()))
                                {
                                    if (StartUp.M_QL_LO_CK != "1")
                                    {
                                        e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                    }
                                    else
                                    {
                                        e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13Lo(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_lo"].ToString());
                                    }
                                }
                                else
                                {
                                    e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;
                                }
                            }
                            else
                                e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;
                        }
                        break;
                    case "so_luong":
                        try
                        {
                            Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                            if (int.Parse(e.Cell.Record.Cells["gia_ton"].Value.ToString()) == 3 && nValue1 == new Decimal(0))
                            {
                                int num = (int)ExMessageBox.Show(530, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                break;
                            }
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result3;
                                Decimal.TryParse(e.Cell.Record.Cells["gia_nt"].Value.ToString(), out result3);
                                Decimal result4;
                                Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result4);
                                Decimal result5;
                                Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result5);
                                Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result1);
                                Decimal nValue2 = this.txtTy_gia.nValue;
                                if (nValue1 == new Decimal(0))
                                {
                                    result3 = new Decimal(0);
                                    e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                    result4 = new Decimal(0);
                                    e.Cell.Record.Cells["gia"].Value = (object)0;
                                }
                                if (nValue1 * result3 != new Decimal(0))
                                {
                                    result5 = SysFunc.Round(nValue1 * result3, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)result5;
                                }
                                if (result5 * nValue2 != new Decimal(0))
                                {
                                    Decimal num = SysFunc.Round(result5 * nValue2, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien"].Value = (object)num;
                                }
                                else if (nValue1 * result4 != new Decimal(0))
                                {
                                    Decimal num = SysFunc.Round(nValue1 * result4, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien"].Value = (object)num;
                                }
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result5, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(result3, StartUpTrans.M_ROUND_GIA);
                                }
                            }
                            this.Sum_ALL();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                    case "gia_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                            Decimal result3;
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result3);
                            Decimal result4;
                            Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result4);
                            Decimal result5;
                            Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result5);
                            Decimal result6;
                            Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result6);
                            Decimal nValue2 = this.txtTy_gia.nValue;
                            if (result3 * nValue1 != new Decimal(0))
                            {
                                bool? isChecked = this.ChkSuaTien.IsChecked;
                                if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                                {
                                    result5 = SysFunc.Round(result3 * nValue1, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)result5;
                                }
                            }
                            if (nValue1 * nValue2 != new Decimal(0))
                            {
                                result4 = SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND_GIA);
                                e.Cell.Record.Cells["gia"].Value = (object)result4;
                            }
                            if (result5 * nValue2 != new Decimal(0))
                            {
                                result6 = SysFunc.Round(result5 * nValue2, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["tien"].Value = (object)result6;
                            }
                            else if (result3 * result4 != new Decimal(0))
                            {
                                result6 = SysFunc.Round(result3 * result4, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["tien"].Value = (object)result6;
                            }
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result5, StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(nValue1, StartUpTrans.M_ROUND_GIA);
                            }
                        }
                        this.Sum_ALL();
                        break;
                    case "tien_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                            Decimal result3;
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result3);
                            Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result2);
                            Decimal result4;
                            Decimal.TryParse(e.Cell.Record.Cells["gia_nt"].Value.ToString(), out result4);
                            Decimal result5;
                            Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result5);
                            Decimal nValue2 = this.txtTy_gia.nValue;
                            if (nValue1 * nValue2 != new Decimal(0))
                            {
                                result5 = SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["tien"].Value = (object)result5;
                            }
                            if (result4 == new Decimal(0))
                            {
                                if (result3 != new Decimal(0))
                                    result4 = SysFunc.Round(nValue1 / result3, StartUpTrans.M_ROUND_GIA);
                                e.Cell.Record.Cells["gia_nt"].Value = (object)result4;
                            }
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(nValue1, StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND_GIA);
                            }
                        }
                        this.Sum_ALL();
                        break;
                    case "gia":
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal nValue = (e.Editor as NumericTextBox).nValue;
                            Decimal result3;
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result3);
                            Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result1);
                            if (result3 * nValue != new Decimal(0))
                            {
                                Decimal num = SysFunc.Round(result3 * nValue, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["tien"].Value = (object)num;
                            }
                            if (!(result3 != new Decimal(0)))
                                ;
                        }
                        this.Sum_ALL();
                        break;
                    case "tien":
                        if (e.Cell.IsDataChanged)
                        {
                            result1 = (e.Editor as NumericTextBox).nValue;
                            Decimal result3;
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result3);
                            Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result2);
                            if (!(result3 != new Decimal(0)))
                                ;
                        }
                        this.Sum_ALL();
                        break;
                    case "ma_vv_i":
                        if (e.Cell.IsDataChanged)
                        {
                            this.UpdateTonKho();
                            break;
                        }
                        break;
                }
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

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmINCTPND.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    DataRecord activeRecord1 = this.GrdCt.ActiveRecord as DataRecord;
                    if (activeRecord1.Cells["ma_vt"].Value == null || activeRecord1.Cells["ma_vt"].Value.ToString() == "")
                        break;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCt, "ma_vt");
                            break;
                    }
                    break;
                case Key.F5:
                    if (this.GrdCt.ActiveRecord != null)
                    {
                        this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                        CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell((this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"]);
                        if (cellValuePresenter != null && cellValuePresenter.Editor is ControlHostEditor editor)
                        {
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(editor);
                            if (string.IsNullOrEmpty(autoCompleteControl.Text.Trim()))
                            {
                                int num1 = (int)ExMessageBox.Show(535, StartupBase.SasObj, "Chưa nhập mã vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else if (autoCompleteControl != null)
                            {
                                if (autoCompleteControl.CheckLostFocus())
                                {
                                    string ma_vt = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"].Value.ToString();
                                    string ten_vt = (this.GrdCt.ActiveRecord as DataRecord).Cells["ten_vt"].Value.ToString();
                                    if (!StartUpTrans.M_LAN.Equals("V"))
                                        ten_vt = (this.GrdCt.ActiveRecord as DataRecord).Cells["ten_vt2"].Value.ToString();
                                    string ma_kho = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_kho_i"].Value.ToString();
                                    DataTable inctpndPx = StartUp.GetINCTPND_PX(ma_vt, ma_kho);
                                    if (inctpndPx.Rows.Count > 0)
                                    {
                                        FrmINCTPND_PX frmInctpndPx = new FrmINCTPND_PX(inctpndPx, ten_vt);
                                        frmInctpndPx.ShowDialog();
                                        int index = this.GrdCt.ActiveRecord.Index;
                                        if (index >= 0 && index <= this.GrdCt.Records.Count - 1)
                                        {
                                            DataRowView drvFrmInctpndPx = frmInctpndPx.drvFrmINCTPND_PX;
                                            if (drvFrmInctpndPx != null)
                                            {
                                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmInctpndPx["gia"];
                                                    (this.GrdCt.DataSource as DataView)[index]["gia"] = drvFrmInctpndPx["gia"];
                                                }
                                                else
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmInctpndPx["gia_nt"];
                                                    (this.GrdCt.DataSource as DataView)[index]["gia"] = drvFrmInctpndPx["gia"];
                                                }
                                                Decimal num2 = new Decimal(0);
                                                Decimal num3 = new Decimal(0);
                                                Decimal num4 = new Decimal(0);
                                                Decimal num5 = new Decimal(0);
                                                Decimal num6 = new Decimal(0);
                                                Decimal num7 = new Decimal(0);
                                                Decimal num8 = this.ParseDecimal((object)(this.GrdCt.DataSource as DataView)[index]["gia_nt"].ToString(), new Decimal(0));
                                                num4 = this.ParseDecimal((object)(this.GrdCt.DataSource as DataView)[index]["gia"].ToString(), new Decimal(0));
                                                Decimal num9 = this.ParseDecimal((object)(this.GrdCt.DataSource as DataView)[index]["so_luong"].ToString(), new Decimal(0));
                                                Decimal nValue = this.txtTy_gia.nValue;
                                                int num10;
                                                if (num8 * num9 != new Decimal(0))
                                                {
                                                    bool? isChecked = this.ChkSuaTien.IsChecked;
                                                    num10 = (isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 ? 1 : 0;
                                                }
                                                else
                                                    num10 = 1;
                                                if (num10 == 0)
                                                {
                                                    num5 = SysFunc.Round(num8 * num9, StartUpTrans.M_ROUND_NT);
                                                    (this.GrdCt.DataSource as DataView)[index]["tien_nt"] = (object)num5;
                                                }
                                                int num11;
                                                if (num5 * nValue != new Decimal(0))
                                                {
                                                    bool? isChecked = this.ChkSuaTien.IsChecked;
                                                    num11 = (isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 ? 1 : 0;
                                                }
                                                else
                                                    num11 = 1;
                                                if (num11 == 0)
                                                    (this.GrdCt.DataSource as DataView)[index]["tien"] = (object)SysFunc.Round(num5 * nValue, StartUpTrans.M_ROUND);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int num12 = (int)ExMessageBox.Show(540, StartupBase.SasObj, "Không có phiếu xuất cho vật tư này!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    }
                                }
                                else
                                {
                                    int num13 = (int)ExMessageBox.Show(545, StartupBase.SasObj, "Không có phiếu xuất cho vật tư này!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                            }
                        }
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(550, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord2))
                        break;
                    int num14 = 0;
                    Cell activeCell = this.GrdCt.ActiveCell;
                    if (activeRecord2.Index != 0 && activeRecord2.Index == this.GrdCt.Records.Count - 1)
                        num14 = activeRecord2.Index - 1;
                    int num15 = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num15 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[activeRecord2.Index].Row);
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[num14 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : num14];
                        this.Sum_ALL();
                    }
                    break;
            }
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmINCTPND.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
        }

        private void ChkSuaTien_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            bool? isChecked = this.ChkSuaTien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.TyGiaValueChange();
        }

        private void ChkNhapGiaTB_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedPn_gia_tb.Value = this.ChkNhapGiaTB.IsChecked.Value;
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

        private void txtMa_gd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_gd.RowResult == null)
                return;
            this.tblTen_gd.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_gd.RowResult["ten_gd"].ToString() : this.txtMa_gd.RowResult["ten_gd2"].ToString();
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (ma_nt == StartUpTrans.M_ma_nt0)
            {
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["sua_tk_vt"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["sua_tk_vt"].Settings.CellMaxWidth = 0.0;
                this.txtTy_gia.IsReadOnly = true;
            }
            else
            {
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["sua_tk_vt"].Visibility = Visibility.Hidden;
                FieldSettings settings1 = this.GrdCt.FieldLayouts[0].Fields["tien"].Settings;
                FieldLength? width = this.GrdCt.FieldLayouts[0].Fields["tien"].Width;
                double num1 = width.Value.Value;
                settings1.CellMaxWidth = num1;
                FieldSettings settings2 = this.GrdCt.FieldLayouts[0].Fields["gia"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["gia"].Width;
                double num2 = width.Value.Value;
                settings2.CellMaxWidth = num2;
                this.GrdCt.FieldLayouts[0].Fields["sua_tk_vt"].Settings.CellMaxWidth = 0.0;
                this.txtTy_gia.IsReadOnly = false;
            }
            this.LanguageProvider.ChangeLanguage((Visual)this.GrdCt, this.LanguageID.Trim() + ".TabInfo.tabItemHT", StartUpTrans.M_LAN, false);
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = new Decimal(0);
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void Sum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        public void CanBangTien()
        {
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal result5 = new Decimal(1);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)SysFunc.Round(result1 * result5, StartUpTrans.M_ROUND);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"], new Decimal(0)) != new Decimal(0))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"] = (object)SysFunc.Round(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"], new Decimal(0)) + (this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"], new Decimal(0)) - result4), StartUpTrans.M_ROUND);
                    break;
                }
            }
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        private void UpdateTonKho()
        {
            string str1 = "";
            string str2 = "";
            string str3 = "";
            string strlo = "";
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseInt(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["vt_ton_kho"], 0) == 1)
                {
                    str1 = str1 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString().Trim();
                    str2 = str2 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString().Trim();
                    str3 = str3 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vv_i"].ToString().Trim();
                    strlo = strlo + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_lo"].ToString().Trim();
                }
                else
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = (object)DBNull.Value;
            }
            if (!string.IsNullOrEmpty(str1) && !string.IsNullOrEmpty(str2))
            {
                string ma_kho = str1.Substring(1);
                string ma_vt = str2.Substring(1);
                string ma_vv = str3.Substring(1);
                string ma_lo = strlo.Substring(1);
                DataTable listTon13 = StartUp.M_QL_LO_CK.ToString() != "1" ? InFuncLib.GetListTon13(StartupBase.SasObj, ma_kho, ma_vt, ma_vv) : InFuncLib.GetListTon13Lo(StartupBase.SasObj, ma_kho, ma_vt, ma_vv, ma_lo);
                if (listTon13 != null && listTon13.Rows.Count > 0)
                {
                    for (int index = 0; index < listTon13.Rows.Count; ++index)
                    {
                        string str4 = listTon13.Rows[index]["ma_kho"].ToString().Trim();
                        listTon13.Rows[index]["ma_kho"] = (object)str4;
                        string str5 = listTon13.Rows[index]["ma_vt"].ToString().Trim();
                        listTon13.Rows[index]["ma_vt"] = (object)str5;
                        string str6 = listTon13.Rows[index]["ma_vv"].ToString().Trim();
                        listTon13.Rows[index]["ma_vv"] = (object)str6;
                        if (StartUp.M_QL_LO_CK.ToString().Trim() == "1")
                        {
                            string str7 = listTon13.Rows[index]["ma_lo"].ToString().Trim();
                            listTon13.Rows[index]["ma_lo"] = (object)str7;
                        }
                    }
                    listTon13.AcceptChanges();
                }
                if (listTon13 != null)
                {
                    for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                    {
                        string str4 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString().Trim();
                        string str5 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString().Trim();
                        string str6 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vv_i"].ToString().Trim();
                        string str7 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_lo"].ToString().Trim();
                        SqlCommand sqlcmd = new SqlCommand("exec [CheckTonvv] @Ma_kho");
                        sqlcmd.Parameters.Add("@Ma_kho", SqlDbType.VarChar).Value = str4;
                        if (this.ParseInt(StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows[0][0], 0) == 0)
                            str6 = "";
                        DataRow[] dataRowArray = StartUp.M_QL_LO_CK.ToString() != "1" ? listTon13.Select("ma_kho LIKE '" + str4 + "' AND ma_vt LIKE '" + str5 + "' AND ma_vv LIKE '" + str6 + "'") : listTon13.Select("ma_kho LIKE '" + str4 + "' AND ma_vt LIKE '" + str5 + "' AND ma_vv LIKE '" + str6 + "' AND ma_lo LIKE '" + str7 + "'");
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = dataRowArray.Length <= 0 ? 0 : dataRowArray[0]["ton13"];
                    }
                }
            }
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        private void GrdCt_CellDeactivating(object sender, CellDeactivatingEventArgs e)
        {
            try
            {
                if (!FrmINCTPND.IsInEditMode.Value || e.Cell.Field.Name != "ma_vt")
                    return;
                CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell(e.Cell);
                if (cellValuePresenter == null || cellValuePresenter.Editor == null)
                    return;
                AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(cellValuePresenter.Editor as ControlHostEditor);
                if (autoCompleteControl1 == null)
                    return;
                DataRowView dataItem = e.Cell.Record.DataItem as DataRowView;
                autoCompleteControl1.SearchInit();
                if (autoCompleteControl1.RowResult != null)
                {
                    dataItem["sua_tk_vt"] = autoCompleteControl1.RowResult["sua_tk_vt"];
                    if (dataItem["sua_tk_vt"].ToString() != "1" && autoCompleteControl1.RowResult["tk_vt"].ToString() != "")
                        e.Cell.Record.Cells["tk_vt"].Value = (object)autoCompleteControl1.RowResult["tk_vt"].ToString();
                }
                AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                if (autoCompleteControl2 != null)
                {
                    autoCompleteControl2.SearchInit();
                    if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                    {
                        e.Cell.Record.Cells["tk_vt"].Value = (object)autoCompleteControl2.RowResult["tk_dl"].ToString();
                        dataItem["tk_vt_dmvt"] = (object)autoCompleteControl2.RowResult["tk_dl"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void FormMain_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.D5 || Keyboard.Modifiers != ModifierKeys.Control)
                return;
            this.txtMa_gd.IsFocus = false;
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmINCTPND.IsInEditMode.Value)
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

        private void BtnXuatNVL_Click(object sender, RoutedEventArgs e)
        {
            if (FrmINCTPND.IsInEditMode.Value)
                return;

            string newstt_recPXD = string.Empty;

            DataTable dataTable = (DataTable)null;
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString()))
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = string.Format("SELECT a.stt_rec,a.ma_ct,a.ma_gd,a.ma_qs,a.so_ct,a.ma_nt,a.ong_ba,a.dien_giai,a.ma_kho,a.ma_kh, b.ma_nx_i as ma_nx FROM PH84 a left join ct84 b on a.stt_rec = b.stt_rec WHERE a.stt_rec LIKE '{0}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString());
                dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];
            }

            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                newstt_recPXD = dataTable.Rows[0]["stt_rec"].ToString();
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("ma_ct", typeof(string));
            dt.Columns.Add("stt_rec", typeof(string));
            dt.Columns.Add("stt_recpx", typeof(string));
            dt.Columns.Add("ma_qs", typeof(string));
            dt.Columns.Add("so_ct", typeof(string));
            dt.Columns.Add("ma_nt", typeof(string));
            dt.Columns.Add("ty_gia", typeof(Decimal));
            dt.Columns.Add("ty_giaf", typeof(Decimal));
            dt.Columns.Add("nguoinop", typeof(string));
            dt.Columns.Add("lydonop", typeof(string));
            dt.Columns.Add("ma_gd", typeof(string));
            dt.Columns.Add("loai_xnvl", typeof(int));
            dt.Columns.Add("ma_kho", typeof(string));
            dt.Columns.Add("ma_nx", typeof(string));
            DataRow row = dt.NewRow();
            dt.Rows.Add(row);

            FrmTaoPXNVL frmTaoPXNVL = new FrmTaoPXNVL();
            frmTaoPXNVL.tbInfoPT = dataTable;
            frmTaoPXNVL.DataContext = (object)dt.DefaultView;
            frmTaoPXNVL.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"].ToString();
            frmTaoPXNVL.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"].ToString().Trim().PadLeft(frmTaoPXNVL.txtso_ct_pt.MaxLength);
            frmTaoPXNVL.txtnguoi_nop.Text = this.txtOng_ba.Text;
            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString().Trim() != "1" && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString().Trim() != "2")
            {
                frmTaoPXNVL.kind = 1;
            }
            else
            {
                frmTaoPXNVL.kind = Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"]);
            }

            frmTaoPXNVL.Ma_nt_ht = this.cbMa_nt.Text;
            frmTaoPXNVL.so_hd = this.txtSo_ct.Text.Trim();
            frmTaoPXNVL.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
            frmTaoPXNVL.ShowDialog();
            if (!frmTaoPXNVL.isOk)
            {
                return;
            }
            else
            {
                bool isSuccess = false;
                FrmWaiting frmWaiting = new FrmWaiting(60600.0);
                frmWaiting.Show();
                if (frmWaiting.PBar.Value <= 60000.0)
                    frmWaiting.Set(frmWaiting.PBar.Value + 1.0);

                isSuccess = true;
                dt = dt.Copy();
                if (string.IsNullOrEmpty(newstt_recPXD))
                    newstt_recPXD = DataProvider.NewTrans(StartupBase.SasObj, "PXD", StartUpTrans.Ws_Id);

                dt.Rows[0]["ma_ct"] = "PXD";
                dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dt.Rows[0]["stt_recpx"] = newstt_recPXD;
                dt.Rows[0]["ma_qs"] = frmTaoPXNVL.txtMa_qs_pt.Text;
                dt.Rows[0]["so_ct"] = frmTaoPXNVL.txtso_ct_pt.Text.PadLeft(frmTaoPXNVL.txtso_ct_pt.MaxLength, ' ');
                dt.Rows[0]["ma_nt"] = frmTaoPXNVL.txtMa_nt.Text;
                dt.Rows[0]["ty_gia"] = frmTaoPXNVL.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.Rate;
                dt.Rows[0]["ty_giaf"] = frmTaoPXNVL.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.RateF;
                dt.Rows[0]["nguoinop"] = frmTaoPXNVL.txtnguoi_nop.Text;
                dt.Rows[0]["lydonop"] = frmTaoPXNVL.txtlydo_nop.Text;
                dt.Rows[0]["ma_gd"] = frmTaoPXNVL.txtMa_gd.Text;
                dt.Rows[0]["loai_xnvl"] = frmTaoPXNVL.txtKind.Text.Equals("1") ? 1 : 2;
                dt.Rows[0]["ma_kho"] = frmTaoPXNVL.txtMa_kho.Text;
                dt.Rows[0]["ma_nx"] = frmTaoPXNVL.txtMa_nx.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"] = frmTaoPXNVL.txtMa_qs_pt.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"] = frmTaoPXNVL.txtso_ct_pt.Text.Trim().PadLeft(frmTaoPXNVL.txtso_ct_pt.MaxLength);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_px"] = "PXD";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"] = newstt_recPXD;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"] = frmTaoPXNVL.txtKind.Text.Equals("1") ? 1 : 2;
                string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                this.so_ct_px_length = StartupBase.SasObj.GetDatabaseFieldLength("so_ct");

                try
                {
                    this.CreatePXD(dt);
                    //Kiểm tra xem đã tao hóa đơn thành công không
                    SqlCommand sqlcmd3 = new SqlCommand();
                    sqlcmd3.CommandText = string.Format("SELECT TOP 1 stt_rec FROM PH84 WHERE stt_rec LIKE '{0}' UNION ALL SELECT TOP 1 stt_rec FROM CT84 WHERE stt_rec LIKE '{0}'", newstt_recPXD.ToString());
                    DataTable dtkt = StartupBase.SasObj.ExcuteReader(sqlcmd3).Tables[0];
                    if (dtkt != null && dtkt.Rows.Count >= 2)
                    {
                        this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                        {
                            DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + _stt_rec1 + "'");
                            if (dataRowArray.Length == 1)
                            {
                                dataRowArray[0]["stt_rec_px"] = newstt_recPXD;
                                dataRowArray[0]["so_ct_px"] = dt.Rows[0]["so_ct"].ToString().Trim().PadLeft(this.so_ct_px_length);
                                dataRowArray[0]["ma_ct_px"] = dt.Rows[0]["ma_ct"];
                                dataRowArray[0]["ma_qs_px"] = dt.Rows[0]["ma_qs"];
                                dataRowArray[0]["loai_xnvl"] = Convert.ToInt32(dt.Rows[0]["loai_xnvl"].ToString());
                            }
                        }));
                        isSuccess = true;
                    }
                    else
                    {
                        this.DeleteVoucherPXD(newstt_recPXD);

                        this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                        {
                            DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + _stt_rec1 + "'");
                            if (dataRowArray.Length == 1)
                            {
                                dataRowArray[0]["stt_rec_px"] = "";
                                dataRowArray[0]["so_ct_px"] = "";
                                dataRowArray[0]["ma_ct_px"] = "";
                                dataRowArray[0]["ma_qs_px"] = "";
                                dataRowArray[0]["loai_xnvl"] = 0;
                            }
                        }));
                        isSuccess = false;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                }
                frmWaiting.Set(60600.0);
                frmWaiting.Close();

                if(isSuccess)
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Successfully created NVL export slip!" : "Tạo phiếu xuất NVL thành công!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                else
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Create an NVL output slip with error. Please check the quota declaration!" : "Tạo phiếu xuất NVL bị lỗi.Hãy kiểm tra lại khai báo định mức!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());

            }
        }

        private void CreatePXD(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[INCTPND-CREATEPXD] @Stt_rec, @Stt_recpx, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct, @loai_xnvl, @ma_kho, @ma_nx");
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = dt.Rows[0]["stt_rec"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@Stt_recpx", SqlDbType.VarChar).Value = dt.Rows[0]["stt_recpx"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@ma_qs", SqlDbType.VarChar).Value = dt.Rows[0]["ma_qs"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@so_ct", SqlDbType.VarChar).Value = dt.Rows[0]["so_ct"].ToString().Trim();
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nt"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@ty_gia", SqlDbType.Decimal).Value = dt.Rows[0]["ty_gia"];
                sqlcmd.Parameters.Add("@ty_giaf", SqlDbType.Decimal).Value = dt.Rows[0]["ty_giaf"];
                sqlcmd.Parameters.Add("@nguoinop", SqlDbType.NVarChar).Value = dt.Rows[0]["nguoinop"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@lydonop", SqlDbType.NVarChar).Value = dt.Rows[0]["lydonop"];
                sqlcmd.Parameters.Add("@ma_gd", SqlDbType.VarChar).Value = dt.Rows[0]["ma_gd"].ToString().Trim(); 
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = dt.Rows[0]["ma_ct"].ToString().Trim();
                sqlcmd.Parameters.Add("@loai_xnvl", SqlDbType.TinyInt).Value = dt.Rows[0]["loai_xnvl"];
                sqlcmd.Parameters.Add("@ma_kho", SqlDbType.VarChar).Value = dt.Rows[0]["ma_kho"].ToString().Trim();
                sqlcmd.Parameters.Add("@ma_nx", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nx"].ToString().Trim();
                //ErrorLog.AddLog(StartupBase.SasObj.SqlString(sqlcmd));
                //string mparams = "'"+ dt.Rows[0]["stt_rec"].ToString().Trim() + "','" + dt.Rows[0]["stt_recpx"].ToString().Trim() + "','" + dt.Rows[0]["ma_qs"].ToString().Trim() + "','" + dt.Rows[0]["so_ct"].ToString().Trim() + "','" + dt.Rows[0]["ma_nt"].ToString().Trim() + "','" + dt.Rows[0]["ty_gia"].ToString().Trim() + "','" + dt.Rows[0]["ty_giaf"].ToString().Trim() + "','" + dt.Rows[0]["nguoinop"].ToString().Trim() + "','" + dt.Rows[0]["lydonop"].ToString().Trim() + "','" + dt.Rows[0]["ma_gd"].ToString().Trim() + "','" + dt.Rows[0]["ma_ct"].ToString().Trim() + "','" + dt.Rows[0]["loai_xnvl"].ToString().Trim() + "','" + dt.Rows[0]["ma_kho"].ToString().Trim() + "','" + dt.Rows[0]["ma_nx"].ToString().Trim() + "'";
                //Console.WriteLine(mparams);
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public void DeleteVoucherPXD(string _stt_rec)
        {
            try
            {
                string format = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd = new SqlCommand(string.Format(format, "[DeleteVoucher]"));
                sqlcmd.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = "PXD";
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = _stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void BtnViewNVL_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString().Trim()))
                return;
            SqlCommand sqlcmd1 = new SqlCommand();
            sqlcmd1.CommandText = "Select count(1) from ph84 WHERE stt_Rec = @stt_rec_px";
            sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_px", SqlDbType.VarChar)).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString();
            if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
            {
                if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu xuất NVL không tồn tại, có xóa thông tin phiếu xuất NVL trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) != MessageBoxResult.Yes)
                    return;
                SqlCommand sqlcmd2 = new SqlCommand();
                sqlcmd2.CommandText = "UPDATE ph74 Set stt_rec_px = '', so_ct_px = '', ma_ct_px = '', ma_qs_px = '', loai_xnvl = 0 WHERE stt_rec = @stt_rec_px; ";
                sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_px", SqlDbType.VarChar)).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"] = 0;
            }
            else
                SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "PXD", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString(), Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), this.BindingSasObj.M_ProcessName);
        }
    }
}
