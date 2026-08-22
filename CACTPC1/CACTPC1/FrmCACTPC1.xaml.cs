using Apttpb;
using ArapLib;
using CatgLib;
using Infragistics.Windows.Controls;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using SasControls;
using SasControls.ControlLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using SasVoucherLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace CACTPC1
{
    public partial class FrmCACTPC1 : FormTrans, IPhValue
    {
        public static int iRow = 0;
        public int iOldRow = 0;
        private bool IsLastGtCell = false;
        private string stt_rec_magd1 = string.Empty;
        private bool txtDiaChiFocusable = true;
        public static CodeValueBindingObject IsInEditModeThue;
        public static CodeValueBindingObject IsInEditMode;
        public static CodeValueBindingObject IsTheodoipt;
        public static CodeValueBindingObject IsSUA_MANT_PTC;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Ma_nt;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject IsCheckedSua_tggs;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject Loai_tg;
        private CodeValueBindingObject Ip_tien_hd;
        public static CodeValueBindingObject Ma_GD_Value;
        private DataSet DsVitual;
        private DataSet dsCheckData;

        public FrmCACTPC1()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtMa_qs.Filter = !(StartUpTrans.Ma_ct == "BN1") ? "ma_cts like '%PC1%' and status = 1" : "ma_cts like '%BN1%' and status = 1";
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.cbMa_nt;
            this.C_So_ct = this.txtSo_ct;
            if (!StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Equals("FK") && (StartUp.dtRegInfo == null || !StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Equals("FK")))
                return;
            this.btnSoHD.Visibility = Visibility.Collapsed;
            this.lblSo_ct_tt.Visibility = Visibility.Collapsed;
            this.ChkSua_tggs.Visibility = Visibility.Collapsed;
        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                if (this.PhData.Rows.Count > 1)
                    FrmCACTPC1.iRow = this.PhData.Rows.Count - 1;
                FrmCACTPC1.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Ma_nt = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tggs = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tggs");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                FrmCACTPC1.IsInEditModeThue = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditModeThue");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                FrmCACTPC1.Ma_GD_Value = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ma_GD_Value");
                this.Loai_tg = (CodeValueBindingObject)this.FormMain.FindResource((object)"Loai_tg");
                this.Ip_tien_hd = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ip_tien_hd");
                this.Ip_tien_hd.Text = StartUp.M_IP_TIEN_HD;
                FrmCACTPC1.IsTheodoipt = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsTheodoipt");
                FrmCACTPC1.IsTheodoipt.Text = StartUp.M_CA_THEO_DOI_PT;
                FrmCACTPC1.IsTheodoipt.Value = StartUp.M_CA_THEO_DOI_PT == "1";
                FrmCACTPC1.IsSUA_MANT_PTC = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsSUA_MANT_PTC");
                FrmCACTPC1.IsSUA_MANT_PTC.Text = StartUp.M_CA_THEO_DOI_PT;
                FrmCACTPC1.IsSUA_MANT_PTC.Value = StartUp.M_SUA_MANT_PTC == "1";

                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmCACTPC1.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                this.Voucher_Ma_nt.Text = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                this.GrdCt.Lan = this.M_LAN;
                this.GrdCtgt.Lan = this.M_LAN;
                this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                FreeCodeFieldLib.InitFreeCodeField(this.BindingSasObj, (BasicGridView)this.Grdhd, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(this.BindingSasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(this.BindingSasObj, (BasicGridView)this.GrdCtChi, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(this.BindingSasObj, (BasicGridView)this.GrdCtgt, StartUpTrans.Ma_ct, 2);
                if (this.PhData.Rows.Count > 0)
                {
                    this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                    this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                    this.GrdLayout00.DataContext = (object)this.PhView;
                    this.GrdCt.DataSource = (IEnumerable)this.CtView;
                    this.GrdCtChi.DataSource = (IEnumerable)this.CtView;
                    this.Grdhd.DataSource = (IEnumerable)this.CtView;
                    this.GrdCtgt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[2].DefaultView;
                    this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
                    if (StartUpTrans.tbStatus.DefaultView.Count == 1)
                        this.txtStatus.IsEnabled = false;
                    this.PhView.ListChanged += new ListChangedEventHandler(this.phValue_ListChanged);
                    this.CtView.ListChanged += new ListChangedEventHandler(this.DefaultView_ListChanged);
                    this.IsCheckedSua_tggs.Value = this.PhView[0]["sua_tggs"].ToString() == "1";
                    this.IsCheckedSua_tien.Value = this.PhView[0]["sua_tien"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = true;
                    this.Loai_tg.Text = this.PhView[0]["loai_tg"].ToString();
                }
                this.Voucher_Ma_nt0.Text = this.PhView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = this.PhView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                FrmCACTPC1.Ma_GD_Value.Text = this.PhView[0]["ma_gd"].ToString();
                FrmCACTPC1.Ma_GD_Value.Value = this.txtMa_gd.Text.Equals("8");
                this.SetStatusVisibleField();
                if (this.txtMa_gd.Text.Equals("3"))
                {
                    this.txtMa_kh.AllowEmty = true;                   
                }
                else
                {
                    this.txtMa_kh.AllowEmty = false;
                    if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                        this.txtSoDuKH.Visibility = this.tblSoDuKH.Visibility = Visibility.Hidden;
                    this.LoadDataDu13();
                }
                this.SetFocusToolbar();

                //Quan ly mẫu HĐ
                if (StartUp.M_MAU_THUE_CK.ToString().Trim() != "1")
                {
                    if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "so_seri0")))
                    {
                        this.GrdCtgt.FieldLayouts[0].Fields["so_seri0"].Visibility = Visibility.Collapsed;
                    }
                    if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "kh_mau_hd")))
                    {
                        this.GrdCtgt.FieldLayouts[0].Fields["kh_mau_hd"].Visibility = Visibility.Collapsed;
                    }
                }

                if (this.Grdhd.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_vv_i")))
                {
                    this.Grdhd.FieldLayouts[0].Fields["ma_vv_i"].Settings.AllowEdit = new bool?(false);
                    this.Grdhd.FieldLayouts[0].Fields["ma_vv_i"].Settings.EditorStyle = (Style)null;
                }
                this.UpdateTotalHT();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.TabInfo_SelectionChanged((object)null, (SelectionChangedEventArgs)null)));
                string str = StartupBase.SasObj.GetOption("M_DC_THUE_CK").ToString();
                if (!(str == "2") && !(str == "3"))
                {
                    if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ten_vt_t")))
                    {
                        this.GrdCt.FieldLayouts[0].Fields["ten_vt_t"].Visibility = Visibility.Collapsed;
                        this.GrdCt.FieldLayouts[0].Fields["ten_vt_t"].Width = new FieldLength?(new FieldLength(0.0));
                        this.GrdCt.FieldLayouts[0].Fields["dvt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["dvt2"].Width = new FieldLength?(new FieldLength(0.0));
                        this.GrdCt.FieldLayouts[0].Fields["so_luong2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["so_luong2"].Width = new FieldLength?(new FieldLength(0.0));
                        if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia")))
                        {
                            this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Collapsed;
                            this.GrdCt.FieldLayouts[0].Fields["gia"].Width = new FieldLength?(new FieldLength(0.0));
                        }
                        if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia_nt")))
                        {
                            this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Collapsed;
                            this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Width = new FieldLength?(new FieldLength(0.0));
                        }
                    }
                    if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ten_vt")))
                    {
                        this.GrdCtgt.FieldLayouts[0].Fields["ten_vt"].Visibility = Visibility.Collapsed;
                        this.GrdCtgt.FieldLayouts[0].Fields["ten_vt"].Width = new FieldLength?(new FieldLength(0.0));
                        this.GrdCtgt.FieldLayouts[0].Fields["dvt"].Visibility = Visibility.Hidden;
                        this.GrdCtgt.FieldLayouts[0].Fields["dvt"].Width = new FieldLength?(new FieldLength(0.0));
                        this.GrdCtgt.FieldLayouts[0].Fields["so_luong"].Visibility = Visibility.Hidden;
                        this.GrdCtgt.FieldLayouts[0].Fields["so_luong"].Width = new FieldLength?(new FieldLength(0.0));
                        if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia")))
                        {
                            this.GrdCtgt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Collapsed;
                            this.GrdCtgt.FieldLayouts[0].Fields["gia"].Width = new FieldLength?(new FieldLength(0.0));
                        }
                        if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia_nt")))
                        {
                            this.GrdCtgt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Collapsed;
                            this.GrdCtgt.FieldLayouts[0].Fields["gia_nt"].Width = new FieldLength?(new FieldLength(0.0));
                        }
                    }
                }
                if (str == "1" || str == "3")
                    return;
                if (this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "dia_chi_t")))
                {
                    this.GrdCt.FieldLayouts[0].Fields["dia_chi_t"].Visibility = Visibility.Collapsed;
                    this.GrdCt.FieldLayouts[0].Fields["dia_chi_t"].Width = new FieldLength?(new FieldLength(0.0));
                }
                if (this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "dia_chi")))
                {
                    this.GrdCtgt.FieldLayouts[0].Fields["dia_chi"].Visibility = Visibility.Collapsed;
                    this.GrdCtgt.FieldLayouts[0].Fields["dia_chi"].Width = new FieldLength?(new FieldLength(0.0));
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void phValue_ListChanged(object sender, ListChangedEventArgs e)
        {
            this.IsCheckedSua_tggs.Value = this.PhView[0]["sua_tggs"].ToString() == "1";
        }

        private void DefaultView_ListChanged(object sender, ListChangedEventArgs e)
        {
            this.SetStatusVisibleField();
        }

        private void V_Truoc()
        {
            if (FrmCACTPC1.iRow <= 1)
                return;
            --FrmCACTPC1.iRow;
            this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
        }

        private void V_Sau()
        {
            if (FrmCACTPC1.iRow >= this.PhData.Rows.Count - 1)
                return;
            ++FrmCACTPC1.iRow;
            this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
        }

        private void V_Dau()
        {
            FrmCACTPC1.iRow = this.PhData.Rows.Count < 2 ? 0 : 1;
            this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
        }

        private void V_Cuoi()
        {
            FrmCACTPC1.iRow = this.PhData.Rows.Count - 1;
            this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
        }

        private void V_Huy()
        {
            FrmCACTPC1.IsInEditMode.Value = false;
            if (this.DsVitual == null || this.PhData.Rows.Count <= 0)
                return;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (this.PhData.Rows.Count > 0)
                    {
                        FrmCACTPC1.iRow = this.iOldRow;
                        this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                        this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                        break;
                    }
                    break;
                case ActionTask.Edit:
                    FormTrans.currActionTask = ActionTask.View;
                    string stt_rec = this.PhView[0]["stt_rec"].ToString();
                    this.PhView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                    this.CtView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                    if (this.CtData.Rows.Count > 0)
                    {
                        foreach (DataRow row in this.CtData.Select("stt_rec='" + stt_rec + "'"))
                            this.CtData.Rows.Remove(row);
                    }
                    if (StartUpTrans.DsTrans.Tables[2].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + stt_rec + "'"))
                            StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                    }
                    this.PhData.Rows.RemoveAt(FrmCACTPC1.iRow);
                    DataRow row1 = this.PhData.NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    this.PhData.Rows.InsertAt(row1, FrmCACTPC1.iRow);
                    this.PhView.RowFilter = "stt_rec= '" + stt_rec + "'";
                    this.CtView.RowFilter = "stt_rec= '" + stt_rec + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
                    this.CtData.Merge(this.DsVitual.Tables[1]);
                    StartUpTrans.DsTrans.Tables[2].Merge(this.DsVitual.Tables[2]);
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                   {
                       if (!FrmCACTPC1.Ma_GD_Value.Text.Equals("1"))
                           return;
                       stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                       StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                       StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString();
                       int num = (int)StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                   }));
                    break;
            }
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(this.PhView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = this.PhView[0]["stt_rec"].ToString();
                string maGd = this.Ma_gd;
                string maNt = this.Ma_nt;
                StartUp.DeleteVoucher(_stt_rec);
                this.PhView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                this.CtView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.PhData.Rows[0]["stt_rec"].ToString() + "'";
                this.PhData.Rows.RemoveAt(FrmCACTPC1.iRow);
                if (StartUpTrans.DsTrans.Tables[2].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                }
                if (this.PhData.Rows.Count > 0)
                {
                    FrmCACTPC1.iRow = FrmCACTPC1.iRow > this.PhData.Rows.Count - 1 ? FrmCACTPC1.iRow - 1 : FrmCACTPC1.iRow;
                    this.PhView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                    this.CtView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.Stt_rec + "'";
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            FormTrans.currActionTask = ActionTask.View;
        }

        private void V_In()
        {
            try
            {
                this.tiHT.Focus();
                StartUp.In();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(this.PhView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            FrmCACTPC1Copy frmCactpC1Copy = new FrmCACTPC1Copy();
            frmCactpC1Copy.Closed += new EventHandler(this._formcopy_Closed);
            frmCactpC1Copy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!(sender as FrmCACTPC1Copy).isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                DataRow row1 = this.PhData.NewRow();
                row1.ItemArray = this.PhData.Rows[FrmCACTPC1.iRow].ItemArray;
                row1["stt_rec"] = (object)str;
                row1["ngay_ct"] = (object)FrmCACTPC1Copy.ngay_ct;
                row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];
                row1["so_ct_tt"] = (object)"";
                this.PhData.Rows.Add(row1);
                if (this.CtView.Count > 0)
                {
                    foreach (DataRow dataRow in this.CtData.Select("stt_rec='" + this.PhView[0]["stt_rec"].ToString() + "'"))
                    {
                        DataRow row2 = this.CtData.NewRow();
                        row2.ItemArray = dataRow.ItemArray;
                        row2["stt_rec"] = (object)str;
                        this.CtData.Rows.Add(row2);
                    }
                }
                DataRow row3 = StartUpTrans.DsTrans.Tables[2].NewRow();
                row3["stt_rec"] = (object)str;
                StartUpTrans.DsTrans.Tables[2].Rows.Add(row3);
                this.iOldRow = FrmCACTPC1.iRow;
                FrmCACTPC1.iRow = this.PhData.Rows.Count - 1;
                this.PhView.RowFilter = "stt_rec= '" + str + "'";
                this.CtView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmCACTPC1.IsInEditMode.Value = true;
                this.SetStatusVisibleField();
                this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            }
        }

        private void NewRowCt()
        {
            DataRow dataRow = this.CtData.NewRow();
            dataRow["stt_rec"] = this.PhView[0]["stt_rec"];
            int result1 = 0;
            int result2 = 0;
            if (this.GrdCt.Records.Count > 0)
            {
                string str = this.CtData.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result1);
            }
            if (this.GrdCtgt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result2);
            }
            int num = (result1 >= result2 ? result1 : result2) + 1;
            dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
            dataRow["loai_hd"] = (object)0;
            dataRow["tien_nt"] = (object)0;
            dataRow["tien"] = (object)0;
            dataRow["thue_nt"] = (object)0;
            dataRow["thue"] = (object)0;
            dataRow["ma_ms"] = (object)"";
            int count = this.CtView.Count;
            if (count > 0)
            {
                dataRow["dien_giaii"] = this.CtView[count - 1].Row["dien_giaii"];
                dataRow["ty_giahtf2"] = (object)0;
                dataRow["ty_gia_ht2"] = (object)0;
            }
            else
            {
                dataRow["dien_giaii"] = this.PhView[0].Row["dien_giai"];
                dataRow["ty_giahtf2"] = (object)0;
                dataRow["ty_gia_ht2"] = (object)0;
            }
            FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, this.CtView, dataRow, 1);
            this.CtData.Rows.Add(dataRow);
        }

        private void NewRowCtChi()
        {
            DataRow dataRow = this.CtData.NewRow();
            dataRow["stt_rec"] = this.PhView[0]["stt_rec"];
            int result1 = 0;
            int result2 = 0;
            if (this.GrdCtChi.Records.Count > 0)
            {
                string str = this.CtData.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result1);
            }
            if (this.GrdCtgt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result2);
            }
            int num = (result1 >= result2 ? result1 : result2) + 1;
            dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
            dataRow["tien_nt"] = (object)0;
            dataRow["tien"] = (object)0;
            dataRow["thue_nt"] = (object)0;
            dataRow["thue"] = (object)0;
            int count = this.CtView.Count;
            if (count > 0)
            {
                dataRow["dien_giaii"] = this.CtView[count - 1].Row["dien_giaii"];
                dataRow["ty_giahtf2"] = (object)0;
                dataRow["ty_gia_ht2"] = (object)0;
            }
            else
            {
                dataRow["dien_giaii"] = this.PhView[0].Row["dien_giai"];
                dataRow["ty_giahtf2"] = (object)0;
                dataRow["ty_gia_ht2"] = (object)0;
            }
            FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, this.CtView, dataRow, 1);
            this.CtData.Rows.Add(dataRow);
        }

        private bool GrdCtChi_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCtChi();
            return true;
        }

        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCt();
            return true;
        }

        private DataRecord CalculateHd2(DataRecord record)
        {
            Decimal num1;
            if (record.Cells["thue_suat"].Value != DBNull.Value)
            {
                num1 = new Decimal(0);
                Decimal result1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);
                Decimal result2;
                Decimal.TryParse(record.Cells["thue_suat"].Value.ToString(), out result2);
                Decimal.TryParse(record.Cells["tt_nt"].Value.ToString(), out result1);
                Decimal nValue = this.txtTy_gia_ht.nValue;
                Decimal num4 = result1 / Decimal.Add(result2 / new Decimal(100), 1);
                Decimal num5 = this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num4, this.M_Round) : SysFunc.Round(num4, this.M_Round_nt);
                record.Cells["tien_nt"].Value = (object)num5;
                record.Cells["thue_nt"].Value = (object)(result1 - num5);
                if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    Decimal num6 = SysFunc.Round(num5 * nValue, this.M_Round);
                    if (num6 > new Decimal(0))
                    {
                        Decimal num7 = new Decimal(0);
                        Decimal num8 = new Decimal(0);
                        Decimal num9 = new Decimal(0);
                        Decimal num10 = new Decimal(0);
                        Decimal num11 = new Decimal(0);
                        Decimal num12 = SysFunc.Round(num6 * result2 / new Decimal(100), this.M_Round);
                        Decimal num13 = SysFunc.Round(result1 * nValue, this.M_Round);
                        Decimal num14 = num13 - (num6 + num12);
                        Decimal num15 = SysFunc.Round(num14 / Decimal.Add(result2 / new Decimal(100), 1), this.M_Round);
                        Decimal num16 = num14 - num15;
                        record.Cells["tien"].Value = (object)(num6 + num15);
                        record.Cells["thue"].Value = (object)(num12 + num16);
                        record.Cells["tt"].Value = (object)num13;
                    }
                }
                else
                {
                    record.Cells["tien"].Value = (object)num5;
                    record.Cells["thue"].Value = (object)(result1 - num5);
                    record.Cells["tt"].Value = (object)result1;
                }
            }
            else
            {
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                num1 = new Decimal(0);
                Decimal.TryParse(record.Cells["tt_nt"].Value.ToString(), out result1);
                if (record.Cells["thue_nt"].Value != DBNull.Value)
                    Decimal.TryParse(record.Cells["thue_nt"].Value.ToString(), out result2);
                Decimal num2 = result1 - result2;
                record.Cells["tien_nt"].Value = (object)num2;
                if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    Decimal num3 = new Decimal(0);
                    Decimal num4 = SysFunc.Round(this.txtTy_gia_ht.nValue * num2, this.M_Round);
                    if (num4 > new Decimal(0))
                    {
                        Decimal result3 = new Decimal(0);
                        record.Cells["tien"].Value = (object)num4;
                        if (record.Cells["thue"].Value != DBNull.Value)
                            Decimal.TryParse(record.Cells["thue"].Value.ToString(), out result3);
                        record.Cells["tt"].Value = (object)(num4 + result3);
                    }
                }
                else
                {
                    record.Cells["tien"].Value = (object)num2;
                    record.Cells["thue"].Value = (object)result2;
                    record.Cells["tt"].Value = (object)result1;
                }
            }
            return record;
        }

        public string GetLoaiTk(AutoCompleteTextBox txt)
        {
            return txt.RowResult.Table.Columns.Contains("loai_cl_no") ? txt.RowResult["loai_cl_no"].ToString().Trim() : StartupBase.SasObj.ExcuteScalar(new SqlCommand("SELECT loai_cl_no FROM dmtk WHERE tk = '" + txt.Text.Trim() + "'")).ToString().Trim();
        }

        private void GrdCtChi_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmCACTPC1.IsInEditMode.Value || this.GrdCtChi.ActiveCell == null || this.CtView.Count <= this.GrdCtChi.ActiveRecord.Index || this.CtData.GetChanges(DataRowState.Deleted) != null)
                    return;
                switch (e.Cell.Field.Name)
                {
                    case "tk_i":
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.Text == "" || autoCompleteControl1.RowResult == null)
                            break;
                        e.Cell.Record.Cells["ten_tk"].Value = autoCompleteControl1.RowResult["ten_tk"];
                        e.Cell.Record.Cells["ten_tk2"].Value = autoCompleteControl1.RowResult["ten_tk2"];
                        if (e.Cell.DataPresenter.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "loai_cl_no")))
                            e.Cell.Record.Cells["loai_cl_no"].Value = (object)this.GetLoaiTk(autoCompleteControl1);
                        break;
                    case "ma_kh_i":
                        AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl2.RowResult == null)
                            break;
                        e.Cell.Record.Cells["ten_kh_i"].Value = autoCompleteControl2.RowResult["ten_kh"];
                        e.Cell.Record.Cells["ten_kh2_i"].Value = autoCompleteControl2.RowResult["ten_kh2"];
                        break;
                    case "tien_nt":
                        if (!e.Cell.IsDataChanged)
                            break;
                        if (!string.IsNullOrEmpty(e.Editor.Text))
                        {
                            e.Cell.Record.Cells["tt_nt"].Value = e.Editor.Value;
                            if (!this.ChkSuaTien.IsChecked.Value)
                            {
                                Decimal result1 = new Decimal(0);
                                Decimal.TryParse((e.Cell.Record.DataItem as DataRowView)["ty_gia_ht2"].ToString(), out result1);
                                Decimal result2;
                                Decimal.TryParse(e.Editor.Text, out result2);
                                Decimal nValue1 = this.txtTy_gia.nValue;
                                Decimal nValue2 = this.txtTy_gia_ht.nValue;
                                Decimal num1 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(result2 * nValue2, this.M_Round) : SysFunc.Round(result2 * nValue1, this.M_Round);
                                Decimal num2 = this.GrdCtChi.FieldLayouts[0].Fields["ty_giahtf2"].Visibility != Visibility.Visible ? SysFunc.Round(result2 * nValue2, this.M_Round) : SysFunc.Round(result2 * result1, this.M_Round);
                                e.Cell.Record.Cells["tien"].Value = (object)num2;
                                e.Cell.Record.Cells["tien_tt"].Value = (object)num1;
                                e.Cell.Record.Cells["tt"].Value = (object)num2;
                            }
                        }
                        this.Tinh_tien_cltg(e.Cell.Record);
                        this.UpdateTotalHT();
                        break;
                    case "ty_giahtf2":
                        if (!e.Cell.IsDataChanged)
                            break;
                        if (string.IsNullOrEmpty(e.Editor.Text))
                            e.Editor.Value = (object)0;
                        Decimal result3 = new Decimal(0);
                        Decimal.TryParse((e.Cell.Record.DataItem as DataRowView)["ty_gia_ht2"].ToString(), out result3);
                        Decimal result4;
                        if (Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result4))
                        {
                            Decimal num = SysFunc.Round(result4 * result3, this.M_Round);
                            e.Cell.Record.Cells["tien"].Value = (object)num;
                            e.Cell.Record.Cells["tt"].Value = (object)num;
                        }
                        this.Tinh_tien_cltg(e.Cell.Record);
                        this.UpdateTotalHT();
                        break;
                    case "tien_tt":
                        this.Tinh_tien_cltg(e.Cell.Record);
                        this.UpdateTotalHT();
                        break;
                    case "dien_giaii":
                        if (e.Cell.Record.Index != 0 || !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                            break;
                        e.Cell.Value = (object)this.txtDien_giai.Text;
                        break;
                    case "tien":
                        e.Cell.Record.Cells["tt"].Value = e.Cell.Record.Cells["tien"].Value;
                        this.Tinh_tien_cltg(e.Cell.Record);
                        this.UpdateTotalHT();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                Decimal num1;
                Decimal num2;
                Decimal num3;
                Decimal num4;
                if (FrmCACTPC1.IsInEditMode.Value && this.GrdCt.ActiveCell != null && this.CtView.Count > this.GrdCt.ActiveRecord.Index && this.CtData.GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "dien_giaii":
                            if (e.Cell.Record.Index == 0 && string.IsNullOrEmpty(e.Editor.Text.Trim()))
                            {
                                e.Cell.Value = (object)this.txtDien_giai.Text;
                                break;
                            }
                            break;
                        case "tk_i":
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null)
                            {
                                e.Cell.Record.Cells["ten_tk"].Value = autoCompleteControl1.RowResult["ten_tk"];
                                e.Cell.Record.Cells["ten_tk2"].Value = autoCompleteControl1.RowResult["ten_tk2"];
                                if (e.Cell.DataPresenter.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "loai_cl_no")))
                                    e.Cell.Record.Cells["loai_cl_no"].Value = (object)this.GetLoaiTk(autoCompleteControl1);
                                this.CtView[this.GrdCt.ActiveRecord.Index]["tk_cn"] = autoCompleteControl1.RowResult["tk_cn"];
                                if (e.Cell.Record.Index == 0 && string.IsNullOrEmpty(e.Cell.Record.Cells["dien_giaii"].Value.ToString()))
                                    e.Cell.Record.Cells["dien_giaii"].Value = (object)this.PhView[0]["dien_giai"].ToString();
                                break;
                            }
                            break;
                        case "tien_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                Decimal num5 = new Decimal(0);
                                Decimal nValue2 = this.txtTy_gia_ht.nValue;
                                Decimal num6 = SysFunc.Round(nValue1 * nValue2, this.M_Round);
                                switch (e.Cell.Record.Cells["loai_hd"].Value.ToString().Trim())
                                {
                                    case "0":
                                        e.Cell.Record.Cells["ma_thue_i"].Value = (object)"";
                                        e.Cell.Record.Cells["thue_suat"].Value = (object)0;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)0;
                                        e.Cell.Record.Cells["thue"].Value = (object)0;
                                        e.Cell.Record.Cells["tt_nt"].Value = (object)nValue1;
                                        e.Cell.Record.Cells["tt"].Value = e.Cell.Record.Cells["tien"].Value = (object)num6;
                                        this.UpdateTotalHT();
                                        break;
                                    case "1":
                                        if (e.Cell.Record.Cells["thue_suat"].Value != DBNull.Value)
                                        {
                                            Decimal result;
                                            Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result);
                                            Decimal num7 = nValue1 * result / new Decimal(100);
                                            Decimal num8 = this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0) ? SysFunc.Round(num7, this.M_Round) : SysFunc.Round(num7, this.M_Round_nt);
                                            e.Cell.Record.Cells["thue_nt"].Value = (object)num8;
                                            e.Cell.Record.Cells["tt_nt"].Value = (object)(num8 + nValue1);
                                            if (!this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                                            {
                                                Decimal num9 = SysFunc.Round(nValue1 * nValue2, this.M_Round);
                                                Decimal num10 = new Decimal(0);
                                                e.Cell.Record.Cells["tien"].Value = (object)num9;
                                                Decimal num11 = SysFunc.Round(num8 * nValue2, this.M_Round);
                                                e.Cell.Record.Cells["thue"].Value = (object)num11;
                                                e.Cell.Record.Cells["tt"].Value = (object)(num9 + num11);
                                                break;
                                            }
                                            e.Cell.Record.Cells["tien"].Value = (object)nValue1;
                                            e.Cell.Record.Cells["thue"].Value = (object)num8;
                                            e.Cell.Record.Cells["tt"].Value = (object)(num8 + nValue1);
                                            break;
                                        }
                                        Decimal result1 = new Decimal(0);
                                        if (e.Cell.Record.Cells["thue_nt"].Value != DBNull.Value)
                                            Decimal.TryParse(e.Cell.Record.Cells["thue_nt"].Value.ToString(), out result1);
                                        e.Cell.Record.Cells["tt_nt"].Value = (object)(nValue1 + result1);
                                        if (!this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                                        {
                                            Decimal num7 = SysFunc.Round(nValue2 * nValue1, this.M_Round);
                                            Decimal result2 = new Decimal(0);
                                            e.Cell.Record.Cells["tien"].Value = (object)num7;
                                            if (e.Cell.Record.Cells["thue"].Value != DBNull.Value)
                                                Decimal.TryParse(e.Cell.Record.Cells["thue"].Value.ToString(), out result2);
                                            e.Cell.Record.Cells["tt"].Value = (object)(num7 + result2);
                                        }
                                        else
                                        {
                                            Decimal result2 = new Decimal(0);
                                            e.Cell.Record.Cells["tien"].Value = (object)nValue1;
                                            if (e.Cell.Record.Cells["thue"].Value != DBNull.Value)
                                                Decimal.TryParse(e.Cell.Record.Cells["thue"].Value.ToString(), out result2);
                                            e.Cell.Record.Cells["tt"].Value = (object)(nValue1 + result2);
                                        }
                                        break;
                                    case "2":
                                        e.Cell.Record.Cells["tt_nt"].Value = e.Editor.Value;
                                        this.CalculateHd2(this.GrdCt.ActiveRecord as DataRecord);
                                        break;
                                }
                                Decimal result3 = new Decimal(0);
                                num3 = new Decimal(0);
                                Decimal result4 = new Decimal(0);
                                Decimal.TryParse(this.txtTy_gia.nValue.ToString(), out result3);
                                Decimal.TryParse(e.Cell.Record.Cells["tt_nt"].Value.ToString(), out result4);
                                Decimal num12 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(result4 * nValue2, this.M_Round) : SysFunc.Round(result4 * result3, this.M_Round);
                                e.Cell.Record.Cells["tien_tt"].Value = (object)num12;
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tien_tt":
                            if (e.Cell.IsDataChanged)
                            {
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "loai_hd":
                            if (!string.IsNullOrEmpty(e.Editor.Text) && e.Editor is ControlHostEditor && e.Cell.IsDataChanged)
                            {
                                switch (e.Cell.Value.ToString().Trim())
                                {
                                    case "0":
                                        e.Cell.Record.Cells["ma_thue_i"].Value = (object)"";
                                        e.Cell.Record.Cells["thue_suat"].Value = (object)0;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)0;
                                        e.Cell.Record.Cells["thue"].Value = (object)0;
                                        e.Cell.Record.Cells["tt_nt"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                        e.Cell.Record.Cells["tt"].Value = e.Cell.Record.Cells["tien"].Value;
                                        e.Cell.Record.Cells["ma_kh_t"].Value = (object)"";
                                        e.Cell.Record.Cells["ma_ms"].Value = (object)"";
                                        this.UpdateTotalHT();
                                        break;
                                    case "2":
                                        this.CalculateHd2(this.GrdCt.ActiveRecord as DataRecord);
                                        break;
                                }
                                this.UpdateTotalHT();
                            }
                            if (e.Cell.Value.ToString().Trim() != "0")
                            {
                                if (string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kh_t"].Value.ToString()) && string.IsNullOrEmpty(e.Cell.Record.Cells["ten_kh_t"].Value.ToString()))
                                {
                                    if (e.Cell.Record.Index == 0)
                                    {
                                        e.Cell.Record.Cells["ma_kh_t"].Value = (object)this.txtMa_kh.Text;
                                        e.Cell.Record.Cells["ten_kh_t"].Value = (object)this.txtTen_kh.Text;
                                        e.Cell.Record.Cells["dia_chi_t"].Value = (object)this.txtDia_chi.Text;
                                        e.Cell.Record.Cells["mst_t"].Value = (object)this.txtMaSoThue.Text;
                                    }
                                    else
                                    {
                                        e.Cell.Record.Cells["ma_kh_t"].Value = (this.GrdCt.Records[e.Cell.Record.Index - 1] as DataRecord).Cells["ma_kh_t"].Value;
                                        e.Cell.Record.Cells["ten_kh_t"].Value = (this.GrdCt.Records[e.Cell.Record.Index - 1] as DataRecord).Cells["ten_kh_t"].Value;
                                        e.Cell.Record.Cells["dia_chi_t"].Value = (this.GrdCt.Records[e.Cell.Record.Index - 1] as DataRecord).Cells["dia_chi_t"].Value;
                                        e.Cell.Record.Cells["mst_t"].Value = (this.GrdCt.Records[e.Cell.Record.Index - 1] as DataRecord).Cells["mst_t"].Value;
                                    }
                                }
                                if (e.Cell.Record.Cells["ma_ms"].Value.ToString().Trim() == "")
                                    e.Cell.Record.Cells["ma_ms"].Value = (object)StartUp.M_ma_ms;
                                break;
                            }
                            break;
                        case "ma_kh_t":
                            if (e.Cell.IsDataChanged)
                            {
                                AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                                if (autoCompleteControl2.RowResult != null && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    e.Cell.Record.Cells["ten_kh_t"].Value = this.M_LAN.ToUpper().Equals("V") ? autoCompleteControl2.RowResult["ten_kh"] : autoCompleteControl2.RowResult["ten_kh2"];
                                    if (!string.IsNullOrEmpty(autoCompleteControl2.RowResult["dia_chi"].ToString().Trim()))
                                        e.Cell.Record.Cells["dia_chi_t"].Value = autoCompleteControl2.RowResult["dia_chi"];
                                    if (!string.IsNullOrEmpty(autoCompleteControl2.RowResult["ma_so_thue"].ToString().Trim()))
                                        e.Cell.Record.Cells["mst_t"].Value = autoCompleteControl2.RowResult["ma_so_thue"];
                                }
                                break;
                            }
                            break;
                        case "ma_thue_i":
                            AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl3.RowResult != null && autoCompleteControl3.IsDataChanged)
                            {
                                e.Cell.Record.Cells["tk_thue_i"].Value = autoCompleteControl3.RowResult["tk_thue_no"];
                                AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["tk_thue_i"]).Editor as ControlHostEditor);
                                autoCompleteControl2.SearchInit();
                                if (autoCompleteControl2.RowResult != null)
                                    (e.Cell.Record.DataItem as DataRowView)["tk_thue_cn"] = autoCompleteControl2.RowResult["tk_cn"];
                                e.Cell.Record.Cells["thue_suat"].Value = autoCompleteControl3.RowResult["thue_suat"];
                                if (!string.IsNullOrEmpty(e.Cell.Record.Cells["thue_suat"].Value.ToString()))
                                {
                                    Decimal result1 = new Decimal(0);
                                    Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result1);
                                    if (e.Cell.Record.Cells["loai_hd"].Value.ToString().Trim().Equals("2"))
                                    {
                                        ControlHostEditor editor = e.Editor as ControlHostEditor;
                                        this.CalculateHd2(this.GrdCt.ActiveRecord as DataRecord);
                                    }
                                    else
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result2);
                                        num1 = new Decimal(0);
                                        Decimal num5 = result2 * result1 / new Decimal(100);
                                        Decimal num6 = this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0) ? SysFunc.Round(num5, this.M_Round) : SysFunc.Round(num5, this.M_Round_nt);
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)num6;
                                        e.Cell.Record.Cells["tt_nt"].Value = (object)(num6 + result2);
                                        if (!this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                                        {
                                            Decimal result3;
                                            Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result3);
                                            num4 = new Decimal(0);
                                            Decimal num7 = SysFunc.Round(result3 * result1 / new Decimal(100), this.M_Round);
                                            e.Cell.Record.Cells["thue"].Value = (object)num7;
                                            e.Cell.Record.Cells["tt"].Value = (object)(result3 + num7);
                                        }
                                        else
                                        {
                                            e.Cell.Record.Cells["tien"].Value = (object)result2;
                                            e.Cell.Record.Cells["thue"].Value = (object)num6;
                                            e.Cell.Record.Cells["tt"].Value = (object)(num6 + result2);
                                        }
                                    }
                                    num3 = new Decimal(0);
                                    Decimal result4 = new Decimal(0);
                                    Decimal nValue1 = this.txtTy_gia.nValue;
                                    Decimal nValue2 = this.txtTy_gia_ht.nValue;
                                    Decimal.TryParse(e.Cell.Record.Cells["tt_nt"].Value.ToString(), out result4);
                                    Decimal num8 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(result4 * nValue2, this.M_Round) : SysFunc.Round(result4 * nValue1, this.M_Round);
                                    e.Cell.Record.Cells["tien_tt"].Value = (object)num8;
                                    this.Tinh_tien_cltg(e.Cell.Record);
                                    this.UpdateTotalHT();
                                }
                                break;
                            }
                            break;
                        case "thue_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (!string.IsNullOrEmpty(e.Editor.Text))
                                {
                                    Decimal result1 = new Decimal(0);
                                    if (e.Cell.Record.Cells["thue_suat"].Value != DBNull.Value)
                                        Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result1);
                                    if (e.Cell.Record.Cells["loai_hd"].Value.ToString().Trim().Equals("2"))
                                    {
                                        NumericTextBox editor = e.Editor as NumericTextBox;
                                        if (e.Cell.IsDataChanged)
                                        {
                                            num2 = new Decimal(0);
                                            Decimal result2 = new Decimal(0);
                                            Decimal result3 = new Decimal(0);
                                            Decimal.TryParse(e.Cell.Record.Cells["tt_nt"].Value.ToString(), out result2);
                                            Decimal.TryParse(e.Cell.Record.Cells["thue_nt"].Value.ToString(), out result3);
                                            Decimal num5 = result2 - result3;
                                            e.Cell.Record.Cells["tien_nt"].Value = (object)num5;
                                            if (!this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                                            {
                                                Decimal num6 = new Decimal(0);
                                                Decimal result4 = new Decimal(0);
                                                Decimal.TryParse(this.txtTy_gia.nValue.ToString(), out result4);
                                                Decimal num7 = SysFunc.Round(num5 * result4, this.M_Round);
                                                if (num7 > new Decimal(0))
                                                {
                                                    if (e.Cell.Record.Cells["thue_suat"].Value != DBNull.Value)
                                                    {
                                                        num4 = new Decimal(0);
                                                        Decimal num8 = new Decimal(0);
                                                        Decimal num9 = new Decimal(0);
                                                        Decimal num10 = new Decimal(0);
                                                        Decimal num11 = new Decimal(0);
                                                        Decimal num12 = SysFunc.Round(num7 * result1 / new Decimal(100), this.M_Round);
                                                        Decimal num13 = SysFunc.Round(result2 * result4, this.M_Round);
                                                        Decimal num14 = num13 - (num7 + num12);
                                                        Decimal num15 = SysFunc.Round(num14 / Decimal.Add(result1 / new Decimal(100),1), this.M_Round);
                                                        Decimal num16 = num14 - num15;
                                                        e.Cell.Record.Cells["tien"].Value = (object)(num7 + num15);
                                                        e.Cell.Record.Cells["thue"].Value = (object)(num12 + num16);
                                                        e.Cell.Record.Cells["tt"].Value = (object)num13;
                                                    }
                                                    else
                                                    {
                                                        Decimal result5 = new Decimal(0);
                                                        e.Cell.Record.Cells["tien"].Value = (object)num7;
                                                        if (e.Cell.Record.Cells["thue"].Value != DBNull.Value)
                                                            Decimal.TryParse(e.Cell.Record.Cells["thue"].Value.ToString(), out result5);
                                                        e.Cell.Record.Cells["tt"].Value = (object)(num7 + result5);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                e.Cell.Record.Cells["tien"].Value = (object)num5;
                                                e.Cell.Record.Cells["thue"].Value = (object)result3;
                                                e.Cell.Record.Cells["tt"].Value = (object)result2;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal result3 = new Decimal(0);
                                        Decimal.TryParse(e.Cell.Record.Cells["thue_nt"].Value.ToString(), out result2);
                                        Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result3);
                                        e.Cell.Record.Cells["tt_nt"].Value = (object)(result2 + result3);
                                        if (this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                                        {
                                            e.Cell.Record.Cells["thue"].Value = (object)result2;
                                            e.Cell.Record.Cells["tien"].Value = (object)result3;
                                            e.Cell.Record.Cells["tt"].Value = (object)(result2 + result3);
                                        }
                                        else
                                        {
                                            Decimal nValue = this.txtTy_gia_ht.nValue;
                                            Decimal num5 = result2 * nValue;
                                            if (num5 > new Decimal(0))
                                            {
                                                e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num5, this.M_Round);
                                                Decimal num6 = Convert.ToDecimal(e.Cell.Record.Cells["tien"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["tien"].Value);
                                                e.Cell.Record.Cells["tt"].Value = (object)(num6 + SysFunc.Round(num5, this.M_Round));
                                            }
                                        }
                                    }
                                }
                                num3 = new Decimal(0);
                                Decimal result = new Decimal(0);
                                Decimal nValue1 = this.txtTy_gia.nValue;
                                Decimal nValue2 = this.txtTy_gia_ht.nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["tt_nt"].Value.ToString(), out result);
                                Decimal num17 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(result * nValue2, this.M_Round) : SysFunc.Round(result * nValue1, this.M_Round);
                                e.Cell.Record.Cells["tien_tt"].Value = (object)num17;
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tk_thue_i":
                            AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl4.RowResult != null)
                            {
                                (e.Cell.Record.DataItem as DataRowView)["tk_thue_cn"] = autoCompleteControl4.RowResult["tk_cn"];
                                break;
                            }
                            break;
                        case "tien":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    Decimal result1 = new Decimal(0);
                                    Decimal result2 = new Decimal(0);
                                    num4 = new Decimal(0);
                                    Decimal.TryParse(e.Editor.Value.ToString(), out result1);
                                    Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                                    Decimal num5 = SysFunc.Round(result1 * result2 / new Decimal(100), this.M_Round);
                                    e.Cell.Record.Cells["thue"].Value = (object)num5;
                                    e.Cell.Record.Cells["tt"].Value = (object)(result1 + num5);
                                }
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "thue":
                            if (e.Cell.IsDataChanged)
                            {
                                if (!string.IsNullOrEmpty(e.Editor.Text))
                                {
                                    if (e.Cell.Record.Cells["loai_hd"].Value.ToString().Trim().Equals("2"))
                                    {
                                        NumericTextBox editor = e.Editor as NumericTextBox;
                                        if (!e.Cell.IsDataChanged)
                                            break;
                                    }
                                    Decimal result1 = new Decimal(0);
                                    Decimal result2 = new Decimal(0);
                                    Decimal.TryParse(e.Editor.Value.ToString(), out result1);
                                    Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result2);
                                    e.Cell.Record.Cells["tt"].Value = (object)(result1 + result2);
                                }
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
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

        private void txtMa_kh_PreviewGotFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
        }

        private void V_Sua()
        {
            if (this.PhData.Rows.Count == 0)
            {
                int num1 = (int)ExMessageBox.Show(530, StartupBase.SasObj, "Không có dữ liệu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                FormTrans.currActionTask = ActionTask.Edit;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(this.PhView.ToTable());
                this.DsVitual.Tables.Add(this.CtView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                if (this.CtView.Count > 0 && (this.txtMa_gd.Text.Trim() != "1" && this.CtView[0]["tk_i"] != DBNull.Value && !string.IsNullOrEmpty(this.CtView[0]["tk_i"].ToString())))
                    this.txtMa_kh.IsFocus = true;
                FrmCACTPC1.IsInEditMode.Value = true;
                if (!(this.txtMa_gd.Text.Trim() != "1"))
                {
                    if (this.CtView.Cast<DataRowView>().Where<DataRowView>((Func<DataRowView, bool>)(q => q["so_ct0"].ToString().Trim() != "")).Any<DataRowView>())
                    {
                        this.txtMa_kh.IsReadOnly = true;
                        this.txtDia_chi.Focus();
                    }
                    else
                        this.txtMa_kh.IsFocus = true;
                }
                this.Voucher_Ma_nt0.Text = this.PhView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = this.PhView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                if (FrmCACTPC1.Ma_GD_Value.Text.Equals("1"))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString();
                    int num2 = (int)StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                }
                this.TabInfo.SelectedIndex = 0;
                this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            }
        }

        private void ChkSua_tggs_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tggs.Value = this.ChkSua_tggs.IsChecked.Value;
        }

        private void ChkSuaTien_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            if (this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                return;
            bool? isChecked = this.ChkSuaTien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                return;
            Decimal nValue1 = this.txtTy_gia.nValue;
            Decimal nValue2 = this.txtTy_gia_ht.nValue;
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal num5 = new Decimal(0);
            Decimal num6 = new Decimal(0);
            Decimal num7;
            if (FrmCACTPC1.Ma_GD_Value.Text.Trim().IndexOfAny(new char[2]
            {
        '1',
        '2'
            }) >= 0)
            {
                DataGridView dataGridView = this.GrdCtChi;
                if (FrmCACTPC1.Ma_GD_Value.Text.Trim() == "1")
                    dataGridView = this.Grdhd;
                foreach (DataRecord record in (IEnumerable<Record>)dataGridView.Records)
                {
                    DataRowView dataItem = record.DataItem as DataRowView;
                    Decimal dec = FNum.ToDec(dataItem["ty_gia_ht2"]);
                    Decimal num8 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num9 = SysFunc.Round(dec * num8, StartUpTrans.M_ROUND);
                    Decimal num10 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num8 * nValue2, this.M_Round) : SysFunc.Round(num8 * nValue1, this.M_Round);
                    if (!this.PhView[0]["ma_nt"].ToString().ToUpper().Trim().Equals(StartUpTrans.M_ma_nt0.ToUpper().Trim()))
                    {
                        record.Cells["tien"].Value = (object)num9;
                        record.Cells["tien_tt"].Value = (object)num10;
                        if (record.FieldLayout.Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "tt")))
                            record.Cells["tt"].Value = (object)num9;
                        else
                            dataItem["tt"] = (object)num9;
                        this.Tinh_tien_cltg(record);
                    }
                }
            }
            else if (FrmCACTPC1.Ma_GD_Value.Text.Equals("9"))
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtChi.Records)
                {
                    Decimal num8 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num9 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                    Decimal num10 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num8 * nValue2, this.M_Round) : SysFunc.Round(num8 * nValue1, this.M_Round);
                    record.Cells["tien"].Value = (object)num9;
                    record.Cells["tien_tt"].Value = (object)num10;
                    record.Cells["tt"].Value = (object)num9;
                    this.Tinh_tien_cltg(record);
                }
            }
            else if (FrmCACTPC1.Ma_GD_Value.Text.Equals("8"))
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCt.Records)
                {
                    Decimal num8 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num9 = Convert.ToDecimal(record.Cells["thue_nt"].Value);
                    Decimal num10 = Convert.ToDecimal(record.Cells["thue"].Value);
                    if (record.Cells["loai_hd"].Value.ToString().Trim().Equals("2"))
                    {
                        this.CalculateHd2(record);
                    }
                    else
                    {
                        Decimal num11 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                        record.Cells["tien"].Value = (object)num11;
                        if (!string.IsNullOrEmpty(record.Cells["thue_suat"].Value.ToString()))
                        {
                            num7 = new Decimal(0);
                            Decimal num12 = Convert.ToDecimal(record.Cells["thue_suat"].Value);
                            num9 = SysFunc.Round(num8 * num12 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            num10 = SysFunc.Round(nValue2 * num9, StartUpTrans.M_ROUND);
                            record.Cells["thue_nt"].Value = (object)num9;
                            record.Cells["thue"].Value = (object)num10;
                        }
                        record.Cells["tt_nt"].Value = (object)(num8 + num9);
                        record.Cells["tt"].Value = (object)(num11 + num10);
                        Decimal num13 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round((num8 + num9) * nValue2, this.M_Round) : SysFunc.Round((num8 + num9) * nValue1, this.M_Round);
                        record.Cells["tien_tt"].Value = (object)num13;
                        this.Tinh_tien_cltg(record);
                    }
                }
            }
            this.UpdateTotalHT();
            if (this.GrdCtgt.Records.Count > 0)
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtgt.Records)
                {
                    if (record.Cells["t_tien_nt"].Value != DBNull.Value)
                    {
                        Decimal num8 = Convert.ToDecimal(record.Cells["t_tien_nt"].Value);
                        Decimal num9 = Convert.ToDecimal(record.Cells["t_thue_nt"].Value);
                        Decimal num10 = Convert.ToDecimal(record.Cells["t_thue"].Value);
                        Decimal num11 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                        record.Cells["t_tien"].Value = (object)num11;
                        if (!string.IsNullOrEmpty(record.Cells["thue_suat"].Value.ToString()))
                        {
                            num7 = new Decimal(0);
                            Decimal num12 = Convert.ToDecimal(record.Cells["thue_suat"].Value);
                            num9 = SysFunc.Round(num8 * num12 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            num10 = SysFunc.Round(num9 * nValue2, StartUpTrans.M_ROUND);
                            record.Cells["t_thue_nt"].Value = (object)num9;
                            record.Cells["t_thue"].Value = (object)num10;
                        }
                        record.Cells["t_tt_nt"].Value = (object)(num8 + num9);
                        record.Cells["t_tt"].Value = (object)(num11 + num10);
                    }
                }
            }
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
                this.txtMa_gd.IsFocus = true;
                DataRow row1 = this.PhData.NewRow();
                row1["stt_rec"] = (object)str;
                row1["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row1["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                if (this.PhData.Rows.Count == 1)
                {
                    row1["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row1["ma_nt"] = this.PhData.Rows[FrmCACTPC1.iRow]["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, this.PhData.Rows[FrmCACTPC1.iRow]["ma_qs"].ToString().Trim());
                }
                DateTime dateTime;
                if (row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    row1["ty_giaf"] = (object)1;
                }
                else
                {
                    DataRow dataRow = row1;
                    string _ma_nt = row1["ma_nt"].ToString().Trim();
                    dateTime = Convert.ToDateTime(row1["ngay_ct"]);
                    DateTime date = dateTime.Date;
                    dataRow["ty_giaf"] = (object)StartUp.GetRates(_ma_nt, date);
                }
                row1["ma_gd"] = this.PhData.Rows.Count > 1 ? (object)this.PhView[0]["ma_gd"].ToString() : (object)StartUpTrans.DmctInfo["ma_gd"].ToString();
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["tk"] = this.PhView[0]["tk"];
                row1["t_tien_nt"] = (object)0;
                row1["t_tien"] = (object)0;
                row1["t_thue_nt"] = (object)0;
                row1["t_thue"] = (object)0;
                row1["t_tt_nt"] = (object)0;
                row1["t_tt"] = (object)0;
                row1["han_tt"] = (object)0;
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("loai_tg") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                    row1["loai_tg"] = (object)StartUpTrans.Getloai_tg(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                DataRow row2 = this.CtData.NewRow();
                row2["stt_rec"] = (object)str;
                row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)1);
                row2["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row2["loai_hd"] = (object)0;
                DataRow dataRow1 = row2;
                DateTime date1;
                if (this.txtNgay_ct.Value != null)
                {
                    dateTime = this.txtNgay_ct.dValue;
                    date1 = dateTime.Date;
                }
                else
                {
                    dateTime = DateTime.Now;
                    date1 = dateTime.Date;
                }
                dataRow1["ngay_ct"] = (object)date1;
                row2["ma_ms"] = (object)"";
                row2["tien_nt"] = (object)0;
                row2["tien_tt"] = (object)0;
                row2["tien"] = (object)0;
                row2["thue_nt"] = (object)0;
                row2["thue"] = (object)0;
                row2["ty_giahtf2"] = (object)0;
                row2["ty_gia_ht2"] = (object)0;
                this.PhData.Rows.Add(row1);
                this.CtData.Rows.Add(row2);
                this.PhView.RowFilter = "stt_rec= '" + str + "'";
                this.CtView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.iOldRow = FrmCACTPC1.iRow;
                FrmCACTPC1.iRow = this.PhData.Rows.Count - 1;
                FrmCACTPC1.IsInEditMode.Value = true;
                this.txtMa_kh.IsReadOnly = false;
                this.txtTen_kh.Text = "";
                this.txtTenTK.Text = "";
                this.TabInfo.SelectedIndex = 0;
                this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Nhan()
        {
            Catinhtg.Tinh(this.GrdCtChi.Records, (IPhValue)this);
            if (this.txtMa_gd.Text.Trim() == "9" && this.PhView[0]["ap_tggd"].ToString().Trim() == "1")
            {
                this.PhView[0]["ty_gia_htf"] = this.PhView[0]["ty_giaf"];
                this.PhView[0]["ty_gia_ht"] = this.PhView[0]["ty_gia"];
            }
            this.CalculateTyGia();
            this.Update_nt0();
            try
            {
                bool flag1 = false;
                if (!this.IsSequenceSave)
                {
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtChi.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    switch (this.txtMa_gd.Text.Trim())
                    {
                        case "1":
                            if (this.GrdCt.Records.Count == 0)
                            {
                                int num = (int)ExMessageBox.Show(225, StartupBase.SasObj, "Chưa vào chi tiết không lưu được! ", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.Grdhd.Focus();
                                return;
                            }
                            this.XoaThue();
                            break;
                        case "8":
                            if (this.GrdCt.Records.Count == 0)
                            {
                                int num = (int)ExMessageBox.Show(230, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCt.Focus();
                                return;
                            }
                            break;
                        default:
                            if (this.GrdCt.Records.Count == 0)
                            {
                                int num = (int)ExMessageBox.Show(220, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCtChi.Focus();
                                return;
                            }
                            this.XoaThue();
                            break;
                    }
                    if (string.IsNullOrEmpty(this.PhView[0]["ma_kh"].ToString().Trim()) && !this.txtMa_gd.Text.Equals("3"))
                    {
                        int num = (int)ExMessageBox.Show(235, StartupBase.SasObj, "Chưa có mã khách hàng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag1 = true;
                    }
                    else if (string.IsNullOrEmpty(this.PhView[0]["tk"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(240, StartupBase.SasObj, "Chưa vào tài khoản có!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_nx.IsFocus = true;
                        flag1 = true;
                    }
                    else if (this.Ma_gd == "3" && this.Ma_nt != this.M_Ma_nt0)
                    {
                        int num = (int)ExMessageBox.Show(66, StartupBase.SasObj, "Đồng tiền giao dịch với mã gd số 3 phải là đồng tiền hạch toán.", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_nx.IsFocus = true;
                        flag1 = true;
                    }
                    else if (this.txtNgay_ct.dValue == new DateTime())
                    {
                        int num = (int)ExMessageBox.Show(245, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtNgay_ct.Focus();
                        flag1 = true;
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
                                    goto label_29;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_29:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = true;
                            this.txtNgay_ct.Focus();
                        }
                        else if (StartUpTrans.M_ngay_lct.Equals("1") && (this.txtngay_lct.Value == null || this.txtngay_lct.Value == DBNull.Value || this.txtngay_lct.dValue == new DateTime()))
                        {
                            int num2 = (int)ExMessageBox.Show(2012, StartupBase.SasObj, "Chưa vào ngày lập chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtngay_lct.Focus();
                            flag1 = true;
                        }
                        else if (string.IsNullOrEmpty(this.CtView[0]["tk_i"].ToString().Trim()))
                        {
                            this.TabInfo.SelectedIndex = 0;
                            switch (this.txtMa_gd.Text.Trim())
                            {
                                case "1":
                                    int num2 = (int)ExMessageBox.Show(225, StartupBase.SasObj, "Chưa vào chi tiết không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.Grdhd.Focus();
                                    break;
                                case "8":
                                    int num3 = (int)ExMessageBox.Show(250, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.Focus();
                                    break;
                                default:
                                    int num4 = (int)ExMessageBox.Show(250, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCtChi.Focus();
                                    break;
                            }
                            flag1 = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(6101, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = true;
                            this.txtMa_qs.IsFocus = true;
                        }
                        else if (string.IsNullOrEmpty(this.PhView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show((int)byte.MaxValue, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag1 = true;
                        }
                    }
                    if (!flag1)
                    {
                        if (FrmCACTPC1.Ma_GD_Value.Text.Trim().Equals("8"))
                        {
                            DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Select("loai_hd <> 0");
                            if (this.TabInfo.SelectedIndex == 0 && dataRowArray.Length > 0)
                                this.tabItem3.Focus();
                            else if (this.TabInfo.SelectedIndex == 1 && dataRowArray.Length == 0)
                                this.tiHT.Focus();
                        }
                        bool flag2 = false;
                        bool flag3 = false;
                        if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                        {
                            for (int index = 0; index < this.GrdCtgt.Records.Count; ++index)
                            {
                                DataRowView dataItem = (this.GrdCtgt.Records[index] as DataRecord).DataItem as DataRowView;
                                if (string.IsNullOrEmpty(dataItem.Row["ma_ms"].ToString().Trim()))
                                {
                                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(dataItem.Row);
                                    StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                                }
                                else
                                {
                                    if (!StartUpTrans.M_MST_CHECK.Equals("0") && (!SysFunc.CheckSumMaSoThue(dataItem.Row["ma_so_thue"].ToString().Trim()) && !string.IsNullOrEmpty(dataItem.Row["ma_so_thue"].ToString().Trim()) && !flag2))
                                    {
                                        int num = (int)ExMessageBox.Show(270, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag2 = true;
                                        if (StartUpTrans.M_MST_CHECK.Equals("2"))
                                            return;
                                    }
                                    if (!StartUp.M_CHK_HD_VAO.Equals(0))
                                    {
                                        string so_ct0 = dataItem.Row["so_ct0"].ToString().Trim();
                                        string so_seri0 = dataItem.Row["so_seri0"].ToString().Trim();
                                        string str;
                                        if (!string.IsNullOrEmpty(dataItem.Row["ngay_ct0"].ToString().Trim()))
                                        {
                                            DateTime dateTime = Convert.ToDateTime(dataItem.Row["ngay_ct0"].ToString().Trim());
                                            dateTime = dateTime.Date;
                                            str = dateTime.ToShortDateString().Substring(0, 10);
                                        }
                                        else
                                            str = "";
                                        string ngay_ct0 = str;
                                        string ma_so_thue = dataItem.Row["ma_so_thue"].ToString().Trim();
                                        if (StartUp.CheckExistHDVao(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString(), so_ct0, so_seri0, ngay_ct0, ma_so_thue) && !flag3)
                                        {
                                            int num = (int)ExMessageBox.Show(420, StartupBase.SasObj, string.Format("Hoá đơn số [{0}], ký hiệu [{1}], ngày [{2}], MST [{3}] đã tồn tại!", (object)so_ct0, (object)so_seri0, (object)ngay_ct0, (object)ma_so_thue), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag3 = true;
                                            if (StartUpTrans.M_MST_CHECK.Equals((object)2))
                                                return;
                                        }
                                    }
                                    if (!string.IsNullOrEmpty(dataItem.Row["so_ct0"].ToString().Trim()) && string.IsNullOrEmpty(dataItem.Row["ngay_ct0"].ToString().Trim()))
                                    {
                                        int num = (int)ExMessageBox.Show(3000, StartupBase.SasObj, "Chưa vào ngày hóa đơn. Không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.GrdCtgt.ActiveCell = (this.GrdCtgt.Records[index] as DataRecord).Cells["ngay_ct0"];
                                        this.GrdCtgt.Focus();
                                    }
                                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["tk_thue_no"].ToString().Trim()))
                                    {
                                        Decimal result = new Decimal(0);
                                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["t_thue_nt"].ToString(), out result);
                                        if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["so_ct0"].ToString().Trim()) && result > new Decimal(0) && !flag1)
                                        {
                                            int num = (int)ExMessageBox.Show(275, StartupBase.SasObj, "Chưa vào tk thuế, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = true;
                                            this.GrdCtgt.ActiveCell = (this.GrdCtgt.Records[index] as DataRecord).Cells["tk_thue_no"];
                                            this.GrdCtgt.Focus();
                                        }
                                    }
                                }
                            }
                            if (!this.CheckVoucherOutofDate())
                                flag1 = true;
                        }
                        if (this.CtView.Count > 0)
                        {
                            using (IEnumerator<Record> enumerator = ((IEnumerable<Record>)this.GrdCt.Records).GetEnumerator())
                            {
                                while (enumerator.MoveNext())
                                {
                                    DataRecord da = (DataRecord)enumerator.Current;
                                    DataRowView dataItem = da.DataItem as DataRowView;
                                    if (string.IsNullOrEmpty(dataItem.Row["tk_i"].ToString().Trim()))
                                    {
                                        this.CtData.Rows.Remove(dataItem.Row);
                                        this.CtData.AcceptChanges();
                                    }
                                    else if (!string.IsNullOrEmpty(dataItem["ma_thue_i"].ToString().Trim()) && string.IsNullOrEmpty(dataItem["tk_thue_i"].ToString().Trim()) && !flag1)
                                    {
                                        int num = (int)ExMessageBox.Show(285, StartupBase.SasObj, "Chưa vào tk thuế, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.tiHT.Focus();
                                        this.GrdCt.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                                       {
                                           this.GrdCt.ActiveCell = da.Cells["tk_thue_i"];
                                           this.GrdCt.Focus();
                                       }));
                                    }
                                }
                            }
                        }
                    }
                }
                if (!flag1)
                {
                    if (!flag1)
                    {
                        decimal num = 0m;
                        decimal num2 = 0m;
                        decimal d2 = 0m;
                        decimal d3 = 0m;
                        if (!this.IsSequenceSave)
                        {
                            decimal _ty_gia = this.txtTy_gia.nValue;
                            decimal d4 = FNum.ToDec(this.txtTy_gia_ht.nValue);
                            if (!this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0) && this.GrdCt.Records.Count > 0 && _ty_gia != 0m && !this.ChkSuaTien.IsChecked.Value)
                            {
                                if (FrmCACTPC1.Ma_GD_Value.Text.Equals("1") || FrmCACTPC1.Ma_GD_Value.Text.Equals("2"))
                                {
                                    decimal d5 = 0m;
                                    decimal? num3 = (from b in this.CtData.AsEnumerable()
                                                     where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString()
                                                     select b).Sum((DataRow x) => x.Field<decimal?>("tien_nt") ?? 0m);
                                    if (num3 != null)
                                    {
                                        decimal.TryParse(num3.ToString(), out d5);
                                    }
                                    decimal d6;
                                    if (this.PhView[0]["loai_cl_no"].ToString() == "0")
                                    {
                                        d6 = SysFunc.Round(_ty_gia * d5, this.M_Round);
                                    }
                                    else
                                    {
                                        d6 = SysFunc.Round(d4 * d5, this.M_Round);
                                    }
                                    decimal? num4 = (from b in this.CtData.AsEnumerable()
                                                     where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString()
                                                     select b).Sum((DataRow x) => x.Field<decimal?>("tien_tt") ?? 0m);
                                    decimal d7 = FNum.ToDec(num4);
                                    decimal num5 = FNum.ToDec(this.CtView[0]["tien_tt"]);
                                    num5 += d6 - d7;
                                    this.CtView[0]["tien_tt"] = num5;
                                    if (this.CtData.AsEnumerable().All((DataRow x) => this.isEquals(FNum.ToDec(x.Field<decimal?>("ty_gia_ht2")), this.Ty_gia, x["stt_rec"].ToString())))
                                    {
                                        this.CtView[0]["tien"] = num5;
                                    }
                                    this.UpdateTotalHT();
                                }
                                else
                                {
                                    int num6 = (from b in this.CtData.AsEnumerable()
                                                where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString() && b.Field<string>("loai_hd") == "2"
                                                select b).Count<DataRow>();
                                    int num7 = (from b in this.CtData.AsEnumerable()
                                                where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString() && ((b.Field<decimal?>("tien_nt") ?? 0m) == 0m || _ty_gia == 0m) && (b.Field<decimal?>("tien") ?? 0m) != 0m && b.Field<string>("loai_hd") != "2"
                                                select b).Count<DataRow>();
                                    if (num7 + num6 < this.GrdCt.Records.Count)
                                    {
                                        var enumerableRowCollection = from b in this.CtData.AsEnumerable()
                                                                      where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString()
                                                                      select b into x
                                                                      select new
                                                                      {
                                                                          tien_nt = x.Field<decimal?>("tien_nt") ?? 0m,
                                                                          tt_nt = x.Field<decimal?>("tt_nt") ?? 0m,
                                                                          thue_nt = x.Field<decimal?>("thue_nt") ?? 0m,
                                                                          tien = x.Field<decimal?>("tien") ?? 0m,
                                                                          tt = x.Field<decimal?>("tt") ?? 0m,
                                                                          thue = x.Field<decimal?>("thue") ?? 0m,
                                                                          tien_tt = x.Field<decimal?>("tien_tt") ?? 0m,
                                                                          tien_cltg = x.Field<decimal?>("tien_cltg") ?? 0m
                                                                      };
                                        if (enumerableRowCollection != null)
                                        {
                                            decimal.TryParse(enumerableRowCollection.Sum(p => p.tien_nt).ToString(), out num);
                                            decimal.TryParse(enumerableRowCollection.Sum(p => p.tt_nt).ToString(), out d3);
                                            decimal.TryParse(enumerableRowCollection.Sum(p => p.tien).ToString(), out num2);
                                            decimal.TryParse(enumerableRowCollection.Sum(p => p.tien_tt).ToString(), out d2);
                                        }
                                        decimal d8 = SysFunc.Round(d4 * num, this.M_Round);
                                        decimal d9 = SysFunc.Round(_ty_gia * num, this.M_Round);
                                        if (this.Ma_gd == "8")
                                        {
                                            d9 = SysFunc.Round(_ty_gia * d3, this.M_Round);
                                        }
                                        for (int i = 0; i < this.GrdCt.Records.Count; i++)
                                        {
                                            DataRecord dataRecord = this.GrdCt.Records[i] as DataRecord;
                                            decimal d10 = FNum.ToDec(dataRecord.Cells["tien_nt"].Value);
                                            decimal num8 = FNum.ToDec(dataRecord.Cells["tien"].Value);
                                            decimal num5 = FNum.ToDec(dataRecord.Cells["tien_tt"].Value);
                                            decimal d11 = FNum.ToDec(dataRecord.Cells["thue"].Value);
                                            if (!(d10 == 0m) || !(num8 != 0m))
                                            {
                                                if (!string.IsNullOrEmpty(dataRecord.Cells["ngay_ct0"].Value.ToString()) && dataRecord.Cells["loai_hd"].Value.ToString().Equals("1"))
                                                {
                                                    num5 = num5 + d9 - d2;
                                                    num8 = num8 + d8 - num2;
                                                    dataRecord.Cells["tien"].Value = num8;
                                                    dataRecord.Cells["tien_tt"].Value = num5;
                                                    dataRecord.Cells["tt"].Value = num8 + d11;
                                                    for (int j = 0; j < this.GrdCtgt.Records.Count; j++)
                                                    {
                                                        DataRecord dataRecord2 = this.GrdCtgt.Records[j] as DataRecord;
                                                        if (dataRecord2.Cells["so_ct0"].Value.Equals(dataRecord.Cells["so_ct0"].Value))
                                                        {
                                                            dataRecord2.Cells["t_tien"].Value = num8;
                                                            dataRecord2.Cells["t_tt"].Value = num8 + d11;
                                                            break;
                                                        }
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                        decimal d12 = (this.txtT_thue_Nt0.Value == DBNull.Value) ? 0m : Convert.ToDecimal(this.txtT_thue_Nt0.nValue);
                                        this.txtT_tt_Nt0.Value = d12 + d8;
                                    }
                                }
                            }
                            this.PhData.AcceptChanges();
                            this.CtData.AcceptChanges();
                            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                            if (string.IsNullOrEmpty(this.PhView[0]["ma_gd"].ToString()))
                            {
                                this.PhView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                            }
                            if (string.IsNullOrEmpty(this.PhView[0]["ma_dvcs"].ToString()))
                            {
                                this.PhView[0]["ma_dvcs"] = StartupBase.SasObj.GetOption("M_MA_DVCS").ToString();
                            }
                        }
                        d3 = (num = (num2 = (d2 = 0m)));
                        var enumerableRowCollection2 = from b in this.CtData.AsEnumerable()
                                                       where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString()
                                                       select b into x
                                                       select new
                                                       {
                                                           tien_nt = x.Field<decimal?>("tien_nt") ?? 0m,
                                                           tt_nt = x.Field<decimal?>("tt_nt") ?? 0m,
                                                           thue_nt = x.Field<decimal?>("thue_nt") ?? 0m,
                                                           tien = x.Field<decimal?>("tien") ?? 0m,
                                                           tt = x.Field<decimal?>("tt") ?? 0m,
                                                           thue = x.Field<decimal?>("thue") ?? 0m,
                                                           tien_tt = x.Field<decimal?>("tien_tt") ?? 0m,
                                                           tien_cltg = x.Field<decimal?>("tien_cltg") ?? 0m
                                                       };
                        if (enumerableRowCollection2 != null)
                        {
                            decimal.TryParse(enumerableRowCollection2.Sum(p => p.tien_nt).ToString(), out num);
                            decimal.TryParse(enumerableRowCollection2.Sum(p => p.tt_nt).ToString(), out d3);
                            decimal.TryParse(enumerableRowCollection2.Sum(p => p.tien).ToString(), out num2);
                            decimal.TryParse(enumerableRowCollection2.Sum(p => p.tien_tt).ToString(), out d2);
                        }
                        if (num == 0m && num2 == 0m)
                        {
                            if (StartUp.M_CHK_ZERO == 1)
                            {
                                ExMessageBox.Show(365, StartupBase.SasObj, "Hạch toán tiền bằng 0!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else if (StartUp.M_CHK_ZERO == 2)
                            {
                                ExMessageBox.Show(370, StartupBase.SasObj, "Hạch toán tiền bằng 0, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = true;
                            }
                        }
                        this.PhView[0]["so_ct_tt"] = "";
                        if (!flag1)
                        {
                            DataTable dataTable = this.PhData.Clone();
                            dataTable.Rows.Add(this.PhView[0].Row.ItemArray);
                            if (!this.IsSequenceSave)
                            {
                                dataTable.Rows[0]["status"] = 0;
                            }
                            DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", dataTable, "stt_rec;row_id");
                            DataTable dataTable2 = this.CtData.Clone();
                            DataTable dataTable3 = StartUpTrans.DsTrans.Tables[2].Clone();
                            foreach (object obj in this.CtView)
                            {
                                DataRowView dataRowView = (DataRowView)obj;
                                if (!this.IsSequenceSave)
                                {
                                    dataRowView["ngay_ct"] = this.PhView[0]["ngay_ct"];
                                    dataRowView["so_ct"] = this.PhView[0]["so_ct"];
                                    dataRowView["ma_ct"] = StartUpTrans.Ma_ct;
                                }
                                dataTable2.Rows.Add(dataRowView.Row.ItemArray);
                            }
                            foreach (object obj2 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                            {
                                DataRowView dataRowView = (DataRowView)obj2;
                                if (!this.IsSequenceSave)
                                {
                                    dataRowView["ma_nt"] = this.PhView[0]["ma_nt"];
                                    dataRowView["ty_gia"] = this.PhView[0]["ty_gia"];
                                    dataRowView["ty_giaf"] = this.PhView[0]["ty_giaf"];
                                    dataRowView["status"] = this.PhView[0]["status"];
                                    dataRowView["ma_gd"] = this.PhView[0]["ma_gd"];
                                    dataRowView["so_ct"] = this.PhView[0]["so_ct"];
                                }
                                dataTable3.Rows.Add(dataRowView.Row.ItemArray);
                            }
                            if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), dataTable2, this.PhView[0]["stt_rec"].ToString()))
                            {
                                ExMessageBox.Show(375, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                            if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctgtdbf"].ToString(), dataTable3, this.PhView[0]["stt_rec"].ToString()))
                            {
                                ExMessageBox.Show(380, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                        }
                    }
                    if (!this.IsSequenceSave)
                    {
                        if (!flag1)
                        {
                            this.dsCheckData = StartUp.CheckData();
                            if (this.dsCheckData.Tables.Count > 0)
                            {
                                this.dsCheckData.Tables[this.dsCheckData.Tables.Count - 1].AcceptChanges();
                            }
                            if (this.dsCheckData.Tables.Count > 0 && this.dsCheckData.Tables[this.dsCheckData.Tables.Count - 1].Rows.Count > 0)
                            {
                                DataTable dataTable4 = this.dsCheckData.Tables[this.dsCheckData.Tables.Count - 1];
                                foreach (object obj3 in dataTable4.DefaultView)
                                {
                                    DataRowView dataRowView2 = (DataRowView)obj3;
                                    if (flag1)
                                    {
                                        break;
                                    }
                                    string text6 = dataRowView2[0].ToString();
                                    string text = text6;
                                    switch (text)
                                    {
                                        case "PH01":
                                            if (StartUpTrans.M_trung_so.Equals("1"))
                                            {
                                                if (ExMessageBox.Show(385, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + base.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                {
                                                    this.txtSo_ct.SelectAll();
                                                    this.txtSo_ct.Focus();
                                                    flag1 = true;
                                                }
                                            }
                                            else if (StartUpTrans.M_trung_so.Equals("2"))
                                            {
                                                ExMessageBox.Show(390, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.txtSo_ct.SelectAll();
                                                this.txtSo_ct.Focus();
                                                flag1 = true;
                                            }
                                            break;
                                        case "PH02":
                                            ExMessageBox.Show(395, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = true;
                                            this.txtMa_nx.IsFocus = true;
                                            break;
                                        case "PH03":
                                            if (StartUp.M_KT_CHI_TQ != 0)
                                            {
                                                ExMessageBox.Show(400, StartupBase.SasObj, "Chi quá số tiền tại quỹ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                if (StartUp.M_KT_CHI_TQ == 2)
                                                {
                                                    flag1 = true;
                                                }
                                            }
                                            break;
                                        case "PH04":
                                            {
                                                bool flag4 = true;
                                                string text7 = "";
                                                DataTable dataTable5 = this.dsCheckData.Tables[0].Copy();
                                                if (dataTable5 != null && dataTable5.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < dataTable5.Rows.Count; i++)
                                                    {
                                                        if (dataTable5.Rows[i]["warning"].ToString().Trim().Equals("1"))
                                                        {
                                                            text7 += dataTable5.Rows[i]["tk"].ToString().Trim();
                                                            text7 += ",";
                                                        }
                                                        else if (dataTable5.Rows[i]["warning"].ToString().Trim().Equals("2"))
                                                        {
                                                            flag4 = false;
                                                            text7 += dataTable5.Rows[i]["tk"].ToString().Trim();
                                                            text7 += ",";
                                                        }
                                                    }
                                                    if (text7.Trim().Length > 0)
                                                    {
                                                        text7 = text7.Substring(0, text7.Length - 1);
                                                        if (flag4)
                                                        {
                                                            ExMessageBox.Show(900, StartupBase.SasObj, string.Format("Tài khoản [{0}] chi quá số tiền tại quỹ!", text7), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        }
                                                        else
                                                        {
                                                            ExMessageBox.Show(905, StartupBase.SasObj, string.Format("Tài khoản [{0}] chi quá số tiền tại quỹ, không lưu được!", text7), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                            flag1 = true;
                                                        }
                                                    }
                                                }
                                                break;
                                            }
                                        case "CT01":
                                            {
                                                int index = (int)Convert.ToInt16(dataRowView2[1]);
                                                ExMessageBox.Show(405, StartupBase.SasObj, "Tk nợ là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag1 = true;
                                                if (FrmCACTPC1.Ma_GD_Value.Text.Trim().Equals("8"))
                                                {
                                                    this.tiHT.Focus();
                                                    this.GrdCt.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(delegate ()
                                                    {
                                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_i"];
                                                        this.GrdCt.Focus();
                                                    }));
                                                }
                                                else
                                                {
                                                    this.GrdCtChi.ActiveCell = (this.GrdCtChi.Records[index] as DataRecord).Cells["tk_i"];
                                                    this.GrdCtChi.Focus();
                                                }
                                                break;
                                            }
                                        case "CT02":
                                            {
                                                int index = (int)Convert.ToInt16(dataRowView2[1]);
                                                if (!flag1 && FrmCACTPC1.Ma_GD_Value.Text.Trim().Equals("8"))
                                                {
                                                    ExMessageBox.Show(410, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.tiHT.Focus();
                                                    this.GrdCt.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(delegate ()
                                                    {
                                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_thue_i"];
                                                        this.GrdCt.Focus();
                                                    }));
                                                }
                                                break;
                                            }
                                        case "CT03":
                                            {
                                                int index = (int)Convert.ToInt16(dataRowView2[1]);
                                                if (!flag1 && FrmCACTPC1.Ma_GD_Value.Text.Trim().Equals("3"))
                                                {
                                                    ExMessageBox.Show(2410, StartupBase.SasObj, "Chưa vào mã khách, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.tiHT.Focus();
                                                    this.GrdCt.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(delegate ()
                                                    {
                                                        this.GrdCtChi.ActiveCell = (this.GrdCtChi.Records[index] as DataRecord).Cells["ma_kh_i"];
                                                        this.GrdCtChi.Focus();
                                                    }));
                                                }
                                                break;
                                            }
                                        case "GT01":
                                            {
                                                int index2 = (int)Convert.ToInt16(dataRowView2[1]);
                                                ExMessageBox.Show(415, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag1 = true;
                                                this.GrdCtgt.ActiveCell = (this.GrdCtgt.Records[index2] as DataRecord).Cells["tk_thue_no"];
                                                this.GrdCtgt.Focus();
                                                break;
                                            }
                                    }
                                    dataTable4.Rows.Remove(dataRowView2.Row);
                                }
                            }
                        }
                    }
                    if (!flag1)
                    {
                        ThreadStart start = delegate ()
                        {
                            this.Post();
                        };
                        new Thread(start).Start();
                        if (!this.IsSequenceSave)
                        {
                            int num10 = base.GetiRow(this.PhData, this.CtView[0]["stt_rec"].ToString());
                            if (FrmCACTPC1.iRow != num10)
                            {
                                DataRow row = this.PhView[0].Row;
                                DataRow dataRow = this.PhData.NewRow();
                                dataRow.ItemArray = row.ItemArray;
                                if (FrmCACTPC1.iRow > num10)
                                {
                                    this.PhData.Rows.InsertAt(dataRow, num10);
                                }
                                else
                                {
                                    this.PhData.Rows.InsertAt(dataRow, num10 + 1);
                                }
                                this.PhData.AcceptChanges();
                                this.PhData.Rows.Remove(row);
                                this.PhData.AcceptChanges();
                                FrmCACTPC1.iRow = num10;
                            }
                            FrmCACTPC1.IsInEditMode.Value = false;
                            FormTrans.currActionTask = ActionTask.View;
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
            string format = "exec [dbo].{0} @stt_rec, @Ma_ct";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[CACTPC1-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@Ma_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.LoadDataDu13()), DispatcherPriority.Background);
            string str1 = "ph46";
            string str2 = "ct46";
            if (this.Ma_ct != "PC1")
            {
                str1 = "ph56";
                str2 = "ct56";
            }
            if (!(StartUp.M_XL_CL_TGGS == "1"))
                return;
            string str3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            string cmdText = "SELECT stt_rec, t_tien, t_tt FROM " + str1 + " WHERE stt_rec = '" + str3 + "';" + "\nSELECT stt_rec, stt_rec0, tien, tien_tt FROM " + str2 + " WHERE stt_rec = '" + str3 + "';";
            if (this.Ma_ct == "BN1")
                cmdText = "SELECT stt_rec, t_tien, t_tt FROM " + str1 + " WHERE stt_rec = '" + str3 + "';" + "\nSELECT stt_rec, stt_rec0, tien, tien_tt FROM " + str2 + " WHERE stt_rec = '" + str3 + "';";
            DataSet ds = StartupBase.SasObj.ExcuteReader(new SqlCommand(cmdText));
            if (ds.Tables.Count < 2)
                return;
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = ds.Tables[0].Rows[0]["t_tien"];
               StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = ds.Tables[0].Rows[0]["t_tt"];
               DataView defaultView = ds.Tables[1].DefaultView;
               defaultView.Sort = "stt_rec0";
               foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
               {
                   DataRowView[] rows = defaultView.FindRows(dataRowView["stt_rec0"]);
                   if (((IEnumerable<DataRowView>)rows).Any<DataRowView>())
                   {
                       dataRowView["tien"] = rows[0]["tien"];
                       dataRowView["tien_tt"] = rows[0]["tien_tt"];
                   }
               }
           }), DispatcherPriority.Background);
        }

        private void btnNhan_Click(object sender, RoutedEventArgs e)
        {
            this.V_Nhan();
        }

        private void V_Xem()
        {
            FormTrans.currActionTask = ActionTask.View;
            string strBrowse = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[0];
            string strBrowseCt = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[1];
            DataTable dataTable = this.PhData.Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, this.CtView, strBrowse, strBrowseCt, "stt_rec");
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            formView.frmBrw.ShowInTaskbar = false;
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.frmBrw.LanguageID = "CACTPC1_5";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string str = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmCACTPC1.iRow = index + 1;
                this.PhView.RowFilter = "stt_rec= '" + str + "'";
                this.CtView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmTim frmTim = new FrmTim(StartupBase.SasObj, StartUpTrans.filterId, StartUpTrans.filterView);
                SysFunc.LoadIcon((Window)frmTim);
                frmTim.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCtgt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (FrmCACTPC1.IsInEditMode.Value && this.GrdCtgt.ActiveCell != null && StartUpTrans.DsTrans.Tables[2].DefaultView.Count > this.GrdCtgt.ActiveRecord.Index && StartUpTrans.DsTrans.Tables[2].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "so_ct0":
                            if (e.Cell.IsDataChanged && this.GrdCtgt.Records.Count == 1)
                            {
                                DataRecord record = this.GrdCtgt.Records[0] as DataRecord;
                                this.CtData.AcceptChanges();
                                Decimal result1 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                var enumerable = from b in this.CtData.AsEnumerable()
                                                 where b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString()
                                                 select b into x
                                                 select new
                                                 {
                                                     tien_nt = x.Field<Decimal?>("tien_nt"),
                                                     tien = x.Field<Decimal?>("tien")
                                                 };
                                 
                                if (enumerable != null)
                                {
                                    Decimal? nullable = enumerable.Sum(p => p.tien_nt);
                                    Decimal.TryParse(nullable.ToString(), out result1);
                                    nullable = enumerable.Sum(p => p.tien);
                                    Decimal.TryParse(nullable.ToString(), out result2);
                                    e.Cell.Record.Cells["t_tien_nt"].Value = (object)result1;
                                    e.Cell.Record.Cells["t_tien"].Value = (object)result2;
                                }
                                break;
                            }
                            break;
                        case "ma_kh":
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null)
                            {
                                e.Cell.Record.Cells["ten_kh"].Value = autoCompleteControl1.RowResult["ten_kh"];
                                e.Cell.Record.Cells["dia_chi"].Value = autoCompleteControl1.RowResult["dia_chi"];
                                e.Cell.Record.Cells["ma_so_thue"].Value = autoCompleteControl1.RowResult["ma_so_thue"];
                                break;
                            }
                            break;
                        case "t_tien_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Editor.Value != DBNull.Value && this.txtTy_gia_ht.Value != null)
                                {
                                    Decimal result1 = new Decimal(0);
                                    Decimal num1 = new Decimal(0);
                                    Decimal.TryParse(e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), out result1);
                                    if (e.Cell.Record.Cells["thue_suat"].Value != DBNull.Value)
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                                        Decimal num2 = result1 * result2 / new Decimal(100);
                                        Decimal num3 = this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num2, this.M_Round) : SysFunc.Round(num2, this.M_Round_nt);
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)num3;
                                        e.Cell.Record.Cells["t_tt_nt"].Value = (object)(result1 + num3);
                                        if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                        {
                                            Decimal num4 = new Decimal(0);
                                            Decimal num5 = new Decimal(0);
                                            Decimal num6 = SysFunc.Round(this.txtTy_gia_ht.nValue * result1, this.M_Round);
                                            if (num6 > new Decimal(0))
                                            {
                                                e.Cell.Record.Cells["t_tien"].Value = (object)num6;
                                                Decimal num7 = SysFunc.Round(num6 * result2 / new Decimal(100), this.M_Round);
                                                e.Cell.Record.Cells["t_thue"].Value = (object)num7;
                                                e.Cell.Record.Cells["t_tt"].Value = (object)(num6 + num7);
                                            }
                                        }
                                        else
                                        {
                                            e.Cell.Record.Cells["t_tien"].Value = (object)result1;
                                            e.Cell.Record.Cells["t_thue"].Value = (object)num3;
                                            e.Cell.Record.Cells["t_tt"].Value = (object)(result1 + num3);
                                        }
                                    }
                                    else
                                    {
                                        e.Cell.Record.Cells["t_tt_nt"].Value = (object)result1;
                                        Decimal num2 = new Decimal(0);
                                        Decimal num3 = SysFunc.Round(this.txtTy_gia_ht.nValue * result1, this.M_Round);
                                        if (num3 > new Decimal(0))
                                        {
                                            e.Cell.Record.Cells["t_tien"].Value = (object)num3;
                                            e.Cell.Record.Cells["t_tt"].Value = (object)num3;
                                        }
                                    }
                                }
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                        case "t_tien":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Editor.Value == DBNull.Value || e.Editor.Value.ToString().Trim().Equals("0"))
                                {
                                    Decimal num = Convert.ToDecimal(e.Cell.Record.Cells["t_tien_nt"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien_nt"].Value);
                                    Decimal nValue = this.txtTy_gia_ht.nValue;
                                    e.Cell.Record.Cells["t_tien"].Value = (object)SysFunc.Round(nValue * num, this.M_Round);
                                }
                                Decimal result1 = new Decimal(0);
                                Decimal num1 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["t_tien"].Value.ToString(), out result1);
                                if (e.Cell.Record.Cells["thue_suat"].Value != DBNull.Value && this.txtTy_gia_ht.Value != null)
                                {
                                    Decimal result2 = new Decimal(0);
                                    Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                                    num1 = SysFunc.Round(result1 * result2 / new Decimal(100), this.M_Round);
                                    e.Cell.Record.Cells["t_thue"].Value = (object)num1;
                                }
                                e.Cell.Record.Cells["t_tt"].Value = (object)(result1 + num1);
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                        case "ma_thue":
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl2.IsDataChanged)
                            {
                                if (autoCompleteControl2.RowResult != null)
                                {
                                    e.Cell.Record.Cells["tk_thue_no"].Value = autoCompleteControl2.RowResult["tk_thue_no"];
                                    e.Cell.Record.Cells["thue_suat"].Value = autoCompleteControl2.RowResult["thue_suat"];
                                }
                                if (!string.IsNullOrEmpty(e.Cell.Record.Cells["thue_suat"].Value.ToString()))
                                {
                                    Decimal num1 = Convert.ToDecimal(e.Cell.Record.Cells["t_tien_nt"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien_nt"].Value);
                                    Decimal num2 = Convert.ToDecimal(e.Cell.Record.Cells["t_tien"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien"].Value);
                                    Decimal num3 = Convert.ToDecimal(e.Cell.Record.Cells["thue_suat"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["thue_suat"].Value);
                                    Decimal num4 = new Decimal(0);
                                    Decimal num5 = new Decimal(0);
                                    Decimal num6 = this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num1 * num3 / new Decimal(100), this.M_Round) : SysFunc.Round(num1 * num3 / new Decimal(100), this.M_Round_nt);
                                    e.Cell.Record.Cells["t_thue_nt"].Value = (object)num6;
                                    Decimal num7 = SysFunc.Round(num2 * num3 / new Decimal(100), this.M_Round);
                                    e.Cell.Record.Cells["t_thue"].Value = (object)num7;
                                    e.Cell.Record.Cells["t_tt_nt"].Value = (object)(num6 + num1);
                                    e.Cell.Record.Cells["t_tt"].Value = (object)(num7 + num2);
                                }
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                        case "tk_thue_no":
                            AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl3.RowResult != null)
                            {
                                (e.Cell.Record.DataItem as DataRowView)["tk_thue_cn"] = autoCompleteControl3.RowResult["tk_cn"];
                                break;
                            }
                            break;
                        case "t_thue_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Cell.Record.Cells["t_thue_nt"].Value == DBNull.Value || e.Cell.Record.Cells["t_thue_nt"].Value.ToString().Trim().Equals("0"))
                                {
                                    Decimal num1 = Convert.ToDecimal(e.Cell.Record.Cells["t_tien_nt"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien_nt"].Value) * Convert.ToDecimal(e.Cell.Record.Cells["thue_suat"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["thue_suat"].Value) / new Decimal(100);
                                    Decimal num2 = this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num1, this.M_Round) : SysFunc.Round(num1, this.M_Round_nt);
                                    e.Cell.Record.Cells["t_thue_nt"].Value = (object)num2;
                                }
                                if (this.txtTy_gia_ht.Value != null && !string.IsNullOrEmpty(e.Editor.Text))
                                {
                                    Decimal result1 = new Decimal(0);
                                    Decimal result2 = new Decimal(0);
                                    Decimal.TryParse(e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), out result1);
                                    Decimal.TryParse(e.Cell.Record.Cells["t_thue_nt"].Value.ToString(), out result2);
                                    e.Cell.Record.Cells["t_tt_nt"].Value = (object)(result1 + result2);
                                }
                                if (this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    e.Cell.Record.Cells["t_tien"].Value = e.Cell.Record.Cells["t_tien_nt"].Value;
                                    e.Cell.Record.Cells["t_thue"].Value = e.Editor.Value;
                                    e.Cell.Record.Cells["t_tt"].Value = e.Cell.Record.Cells["t_tt_nt"].Value;
                                }
                                else
                                {
                                    Decimal nValue = this.txtTy_gia_ht.nValue;
                                    Decimal num1 = Convert.ToDecimal(e.Cell.Record.Cells["t_thue_nt"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_thue_nt"].Value) * nValue;
                                    if (num1 > new Decimal(0))
                                    {
                                        e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(num1, this.M_Round);
                                        Decimal num2 = Convert.ToDecimal(e.Cell.Record.Cells["t_tien"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien"].Value);
                                        e.Cell.Record.Cells["t_tt"].Value = (object)(num2 + SysFunc.Round(num1, this.M_Round));
                                    }
                                }
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                        case "t_thue":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Cell.Record.Cells["t_thue"].Value == DBNull.Value || e.Cell.Record.Cells["t_thue"].Value.ToString().Trim().Equals("0"))
                                {
                                    Decimal num1 = Convert.ToDecimal(e.Cell.Record.Cells["t_tien"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["t_tien"].Value);
                                    Decimal num2 = Convert.ToDecimal(e.Cell.Record.Cells["thue_suat"].Value.Equals((object)DBNull.Value) ? (object)0 : e.Cell.Record.Cells["thue_suat"].Value);
                                    e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(num1 * num2 / new Decimal(100), this.M_Round);
                                }
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                    }
                }
                if (e.Cell.Field.Index != this.GrdCtgt.FieldLayouts[0].Fields.Count - 1)
                    return;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                   this.GrdCtgt.ActiveCell = (this.GrdCtgt.ActiveRecord as DataRecord).Cells["ma_ms"];
               }));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private bool GrdCtgt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            return this.NewRowCtGt();
        }

        private void cbMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null)
                return;
            if (this.cbMa_nt.IsDataChanged)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.cbMa_nt.RowResult["loai_tg"];
                this.Loai_tg.Text = this.cbMa_nt.RowResult["loai_tg"].ToString();
                this.Voucher_Ma_nt0.Text = this.PhView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = this.PhView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.SetStatusVisibleField();
                if (this.cbMa_nt.RowResult["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    this.txtTy_gia.Value = (object)1;
                else
                {
                    if (this.txtTy_gia.Value.ToString().Trim() == "1" || this.txtTy_gia.Value.ToString().Trim() == "0")
                        this.txtTy_gia.Value = (object)StartUp.GetRates(this.cbMa_nt.RowResult["ma_nt"].ToString().Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
                }
                this.CalculateTyGia();
            }
            this.SetRound();
        }

        private void SetStatusVisibleField()
        {
            this.ChangeLanguage();
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value || (this.txtMa_kh.RowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim())))
                return;
            this.txtTen_kh.Text = !this.M_LAN.ToUpper().Equals("V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
            if (string.IsNullOrEmpty(this.txtOng_ba.Text.Trim()))
                this.txtOng_ba.Text = this.txtMa_kh.RowResult["doi_tac"].ToString();
            this.txtMaSoThue.Text = this.txtMa_kh.RowResult["ma_so_thue"].ToString();
            if (!string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
                this.txtDia_chi.Text = this.txtMa_kh.RowResult["dia_chi"].ToString();
            if (this.CtView.Count == 1 && (this.CtView[0]["tk_i"] == DBNull.Value || string.IsNullOrEmpty(this.CtView[0]["tk_i"].ToString())))
            {
                this.CtView[0]["tk_i"] = (object)this.txtMa_kh.RowResult["tk"].ToString();
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("SELECT ten_tk, ten_tk2 FROM dmtk WHERE tk = '{0}'", (object)this.txtMa_kh.RowResult["tk"].ToString().Trim())));
                if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    if (StartUpTrans.DsTrans.Tables[1].Columns.Contains("ten_tk"))
                        StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ten_tk"] = dataSet.Tables[0].Rows[0]["ten_tk"];
                    if (StartUpTrans.DsTrans.Tables[1].Columns.Contains("ten_tk2"))
                        StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ten_tk2"] = dataSet.Tables[0].Rows[0]["ten_tk2"];
                }
            }
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            DataRow rowResult = this.txtMa_kh.RowResult;
            bool? isChecked = this.ChkTheo_doi_pt.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy) && this.ParseInt((object)dataRowView["han_tt"].ToString(), 0) == 0)
                dataRowView["han_tt"] = (object)this.ParseInt(rowResult["han_tt"], 0);
            this.txtDiaChiFocusable = this.txtMa_kh.RowResult["dia_chi"].ToString().Trim().Equals("");
            this.LoadDataDu13();
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
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

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) && !this.txtMa_nx.IsReadOnly)
                this.txtTenTK.Text = !this.M_LAN.ToUpper().Equals("V") ? this.txtMa_nx.RowResult["ten_tk2"].ToString() : this.txtMa_nx.RowResult["ten_tk"].ToString();
            if (this.Ma_gd == "3" && this.txtMa_nx.RowResult["ma_nt"].ToString().Trim() != this.M_Ma_nt0)
            {
                int num = (int)ExMessageBox.Show(66, StartupBase.SasObj, "Đồng tiền giao dịch với mã gd số 3 phải là đồng tiền hạch toán.", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_nx.IsFocus = true;
            }
            else
            {
                if (this.txtMa_nx.RowResult != null && FrmCACTPC1.IsInEditMode.Value)
                {
                    if (StartUp.M_SUA_MANT_PTC.ToString().Trim() == "1")
                    {
                        this.cbMa_nt.Text = !string.IsNullOrEmpty(this.cbMa_nt.Text.ToString()) ? this.cbMa_nt.Text.ToString() : this.txtMa_nx.RowResult["ma_nt"].ToString();
                    }
                    else
                    {
                        this.cbMa_nt.Text = this.txtMa_nx.RowResult["ma_nt"].ToString();
                    }
                    //this.cbMa_nt.Text = this.txtMa_nx.RowResult["ma_nt"].ToString();
                    this.cbMa_nt.IsDataChanged = true;
                    this.cbMa_nt.SearchInit();
                    this.cbMa_nt_PreviewLostFocus((object)this.cbMa_nt, (KeyboardFocusChangedEventArgs)null);
                    this.PhView[0]["loai_cl_no"] = this.txtMa_nx.RowResult["loai_cl_no"];
                    this.PhView[0]["loai_cl_co"] = this.txtMa_nx.RowResult["loai_cl_co"];
                    string str = this.PhView[0]["loai_cl_no"].ToString();
                    if (this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                    {
                        this.txtTy_gia.Value = (object)1;
                    }
                    else
                    {
                        switch (this.Ma_gd)
                        {
                            case "1":
                            case "2":
                            case "3":
                                if (str == "0")
                                {
                                    this.txtTy_gia_ht.Value = (object)0;
                                    break;
                                }
                                this.txtTy_gia.Value = (object)0;
                                if (this.txtTy_gia_ht.nValue == new Decimal(0))
                                    this.txtTy_gia_ht.Value = (object)this.GetTggd(this.Ma_nt, this.Ngay_ct);
                                break;
                            case "8":
                            case "9":
                                if (str != "0")
                                    this.txtTy_gia.Value = (object)0;
                                if (this.txtTy_gia_ht.nValue == new Decimal(0))
                                {
                                    this.txtTy_gia_ht.Value = (object)this.GetTggd(this.Ma_nt, this.Ngay_ct);
                                    break;
                                }
                                break;
                        }
                    }
                    this.txtTy_gia_ht_LostFocus((object)null, (RoutedEventArgs)null);
                }
                this.LoadDataDu13();
            }
        }

        private void LoadDataDu13()
        {
            this.txtSoDuKH.Value = (object)ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk"].ToString());
        }

        private Decimal GetTggd(string ma_nt, DateTime ngay_ct)
        {
            Decimal num = new Decimal(0);
            SqlCommand sqlcmd = new SqlCommand("SELECT ty_gia FROM dmtgnt WHERE ma_nt = @Ma_nt AND ngay_ct <= @Ngay_ct ORDER by ngay_ct DESC");
            sqlcmd.Parameters.Add("@Ma_nt", SqlDbType.VarChar).Value = (object)ma_nt;
            sqlcmd.Parameters.Add("@Ngay_ct", SqlDbType.SmallDateTime).Value = (object)ngay_ct;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            if (dataSet.Tables[0].Rows.Count != 0)
                num = (Decimal)dataSet.Tables[0].Rows[0][0];
            return num;
        }

        private void GrdCtChi_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                this.NewRowCtChi();
                this.GrdCtChi.ActiveRecord = this.GrdCtChi.Records[this.GrdCtChi.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                return;
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                this.NewRowCt();
                this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                return;
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
        }

        private void GrdCtChi_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtChi.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete") && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
                        break;
                    DataRecord activeRecord1 = this.GrdCt.ActiveRecord as DataRecord;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.GrdCtChi.ActiveRecord = this.GrdCtChi.Records[this.GrdCtChi.Records.Count - 1];
                            this.GrdCtChi.ActiveCell = (this.GrdCtChi.ActiveRecord as DataRecord).Cells["tk_i"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCtChi, "tk_i");
                            break;
                    }
                    break;
                case Key.F5:
                    if (StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Equals("FA") || StartUp.dtRegInfo != null && !StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Equals("FK"))
                    {
                        Catinhtg.Tinh(this.GrdCtChi.Records, (IPhValue)this);
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(755, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtChi.ActiveRecord is DataRecord activeRecord2))
                        break;
                    Cell activeCell = this.GrdCtChi.ActiveCell;
                    int num1 = activeRecord2.Index;
                    if (activeRecord2.Index == 0)
                    {
                        if (this.GrdCtChi.Records.Count == 1)
                        {
                            this.GrdCtChi_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                            (this.GrdCtChi.Records[1] as DataRecord).Cells["dien_giaii"].Value = (object)"";
                        }
                    }
                    else if (activeRecord2.Index == this.GrdCtChi.Records.Count - 1)
                        num1 = activeRecord2.Index - 1;
                    int num2 = this.GrdCtChi.ActiveCell == null ? 0 : this.GrdCtChi.ActiveCell.Field.Index;
                    this.GrdCtChi.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        this.CtData.Rows.Remove(this.CtView[activeRecord2.Index].Row);
                        this.CtData.AcceptChanges();
                        if (this.GrdCtChi.Records.Count > 0)
                            this.GrdCtChi.ActiveRecord = this.GrdCtChi.Records[num1 > this.GrdCtChi.Records.Count - 1 ? this.GrdCtChi.Records.Count - 1 : num1];
                        this.UpdateTotalHT();
                    }
                    break;
            }
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete") && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
                        break;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.StartEditMode);
                            this.NewRowCt();
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["tk_i"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCt, "tk_i");
                            break;
                    }
                    break;
                case Key.F5:
                    if (StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Equals("FA") || StartUp.dtRegInfo != null && !StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Equals("FK"))
                    {
                        Catinhtg.Tinh(this.GrdCtChi.Records, (IPhValue)this);
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(760, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord))
                        break;
                    Cell activeCell = this.GrdCt.ActiveCell;
                    int num1 = activeRecord.Index;
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
                        this.CtData.Rows.Remove(this.CtView[activeRecord.Index].Row);
                        this.CtData.AcceptChanges();
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[num1 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : num1];
                        this.UpdateTotalHT();
                    }
                    break;
            }
        }

        private void GrdCtgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && this.PhView[0]["ma_gd"].ToString().Equals("8"))
            {
                this.NewRowCtGt();
                this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
            }
            if (Keyboard.IsKeyDown(Key.Tab) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Return && e.Key != Key.Tab)
                return;
            Cell activeCell = this.GrdCtgt.ActiveCell;
            if (activeCell != null && activeCell.Field.Index == this.GrdCtgt.FieldLayouts[0].Fields.Count - 1)
                this.IsLastGtCell = true;
        }

        private void GrdCtgt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value || !this.PhView[0]["ma_gd"].ToString().Equals("8"))
                return;
            if (!FrmCACTPC1.IsInEditModeThue.Value)
            {
                this.GrdCtgtMoveToSave();
            }
            else
            {
                switch (e.Key)
                {
                    case Key.Tab:
                    case Key.Return:
                        if (this.GrdCtgt.ActiveCell != null && this.IsLastGtCell)
                        {
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                           {
                               this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                               this.GrdCtgt.ActiveCell = (this.GrdCtgt.ActiveRecord as DataRecord).Cells["ma_ms"];
                           }));
                            this.IsLastGtCell = false;
                            break;
                        }
                        break;
                    case Key.F4:
                        this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                        this.GrdCtChi.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                        this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                        this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                        if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete"))
                        {
                            if (!(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
                                break;
                        }
                        else if (this.GrdCtgt.ActiveCell != null && this.GrdCtgt.ActiveCell.Field.GetType().Equals(typeof(NotEmptyField)) && (this.GrdCtgt.ActiveCell.Value == DBNull.Value || string.IsNullOrEmpty(this.GrdCtgt.ActiveCell.Value.ToString().Trim())))
                            break;
                        this.NewRowCtGt();
                        this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
                        this.GrdCtgt.ActiveCell = (this.GrdCtgt.ActiveRecord as DataRecord).Cells["so_ct0"];
                        break;
                    case Key.F5:
                        if (StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Equals("FA") || StartUp.dtRegInfo != null && !StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Equals("FK"))
                        {
                            Catinhtg.Tinh(this.GrdCtChi.Records, (IPhValue)this);
                            break;
                        }
                        break;
                    case Key.F8:
                        if (this.GrdCtgt.ActiveRecord == null || ((this.GrdCtgt.ActiveRecord as DataRecord).DataItem as DataRowView)["thue_ct"].ToString() == "1" || (ExMessageBox.Show(765, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtgt.ActiveRecord is DataRecord activeRecord)))
                            break;
                        int num = 0;
                        Cell activeCell = this.GrdCtgt.ActiveCell;
                        if (activeRecord.Index == 0)
                        {
                            if (this.GrdCtgt.Records.Count == 1)
                                this.GrdCtgt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                            num = activeRecord.Index;
                        }
                        else if (activeRecord.Index > 0)
                            num = activeRecord.Index - 1;
                        if ((this.GrdCtgt.ActiveCell == null ? 0 : this.GrdCtgt.ActiveCell.Field.Index) >= 0)
                        {
                            this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                            StartUpTrans.DsTrans.Tables[2].Rows.Remove(StartUpTrans.DsTrans.Tables[2].DefaultView[activeRecord.Index].Row);
                            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                            if (this.GrdCtgt.Records.Count > 0)
                                this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[num > this.GrdCtgt.Records.Count - 1 ? this.GrdCtgt.Records.Count - 1 : num];
                        }
                        break;
                }
            }
        }

        private int GetLastIndexVisibleField(DataGridView grd)
        {
            int num = -1;
            for (int index = grd.FieldLayouts[0].Fields.Count - 1; num == -1 && index >= 0; --index)
            {
                if (grd.FieldLayouts[0].Fields[index].Visibility == Visibility.Visible && grd.FieldLayouts[0].Fields[index].CellMaxWidthResolved != 0.0)
                    num = index;
            }
            return num;
        }

        private void GrdCtgtMoveToSave()
        {
            if (FrmCACTPC1.IsInEditModeThue.Value || Keyboard.Modifiers != ModifierKeys.None)
                return;
            Cell activeCell = this.GrdCtgt.ActiveCell;
            if (activeCell.Field.Index != this.GetLastIndexVisibleField(this.GrdCtgt) || activeCell.Record.Index != activeCell.DataPresenter.Records.Count - 1)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private bool NewRowCtGt()
        {
            try
            {
                if (!FrmCACTPC1.IsInEditModeThue.Value)
                {
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
                    return false;
                }
                DataRow dataRow = StartUpTrans.DsTrans.Tables[2].NewRow();
                dataRow["stt_rec"] = this.PhView[0]["stt_rec"];
                int result1 = 0;
                int result2 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = this.CtData.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result1);
                }
                if (this.GrdCtgt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == this.PhView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result2);
                }
                int num = (result1 >= result2 ? result1 : result2) + 1;
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ngay_ct"] = this.txtNgay_ct.Value;
                dataRow["ten_vt"] = (object)this.txtDien_giai.Text;
                dataRow["t_tien_nt"] = (object)0;
                dataRow["t_tien"] = (object)0;
                dataRow["thue_suat"] = (object)0;
                dataRow["t_thue_nt"] = (object)0;
                dataRow["t_thue"] = (object)0;
                dataRow["thue_ct"] = (object)0;
                if (this.GrdCtgt.Records.Count > 0 && this.GrdCtgt.ActiveRecord != null)
                {
                    DataRecord activeRecord = this.GrdCtgt.ActiveRecord as DataRecord;
                    dataRow["ma_ms"] = activeRecord.Cells["ma_ms"].Value;
                }
                else
                    dataRow["ma_ms"] = StartupBase.SasObj.GetOption("M_MA_MS");
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[2].DefaultView, dataRow, 2);
                StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow);
                return true;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
                return false;
            }
        }

        private void UpdateTotalHT()
        {
            try
            {
                if (FormTrans.currActionTask == ActionTask.View)
                    return;
                this.CtData.AcceptChanges();
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal result3 = new Decimal(0);
                Decimal result4 = new Decimal(0);
                Decimal result5 = new Decimal(0);
                Decimal result6 = new Decimal(0);
                Decimal result7 = new Decimal(0);
                Decimal result8 = new Decimal(0);
                string stt_rec = this.PhView[0]["stt_rec"].ToString();
                var source = from o in this.CtData.AsEnumerable()
                             where o.Field<string>("stt_rec") == stt_rec
                             select o into x
                             select new
                             {
                                 tien_nt = x.Field<Decimal?>("tien_nt"),
                                 tien = x.Field<Decimal?>("tien"),
                                 tien_tt = x.Field<Decimal?>("tien_tt"),
                                 tien_cltg = x.Field<Decimal?>("tien_cltg"),
                                 thue_nt = x.Field<Decimal?>("thue_nt"),
                                 thue = x.Field<Decimal?>("thue"),
                                 tt_nt = x.Field<Decimal?>("tt_nt"),
                                 tt = x.Field<Decimal?>("tt")
                             };
                if (source != null)
                {
                    Decimal? nullable = source.Sum(p => p.tien_nt);
                    Decimal.TryParse(nullable.ToString(), out result1);
                    nullable = source.Sum(p => p.tien);
                    Decimal.TryParse(nullable.ToString(), out result2);
                    nullable = source.Sum(p => p.tien_tt);
                    Decimal.TryParse(nullable.ToString(), out result3);
                    nullable = source.Sum(p => p.tien_cltg);
                    Decimal.TryParse(nullable.ToString(), out result4);
                    nullable = source.Sum(p => p.thue_nt);
                    Decimal.TryParse(nullable.ToString(), out result5);
                    nullable = source.Sum(p => p.thue);
                    Decimal.TryParse(nullable.ToString(), out result6);
                    nullable = source.Sum(p => p.tt_nt);
                    Decimal.TryParse(nullable.ToString(), out result7);
                    nullable = source.Sum(p => p.tt);
                    Decimal.TryParse(nullable.ToString(), out result8);
                }
                Decimal num1 = result1 + result5;
                Decimal num2 = result2 + result6;
                this.PhView[0]["t_tien_nt"] = (object)result1;
                this.PhView[0]["t_tien"] = (object)result2;
                this.PhView[0]["t_tien_tt"] = (object)result3;
                this.PhView[0]["t_tien_cltg"] = (object)result4;
                this.PhView[0]["t_thue_nt"] = (object)result5;
                this.PhView[0]["t_thue"] = (object)result6;
                this.PhView[0]["t_tt_nt"] = (object)num1;
                this.PhView[0]["t_tt"] = (object)num2;
                if (result5 == new Decimal(0))
                    this.UpdateTotalThue();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void UpdateTotalThue()
        {
            try
            {
                if (FormTrans.currActionTask == ActionTask.View)
                    return;
                this.CtData.AcceptChanges();
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                string stt_rec = this.PhView[0]["stt_rec"].ToString();
                var source = from o in StartUpTrans.DsTrans.Tables[2].AsEnumerable()
                             where o.Field<string>("stt_rec") == stt_rec
                             select o into x
                             select new
                             {
                                 thue_nt = x.Field<Decimal?>("t_thue_nt"),
                                 thue = x.Field<Decimal?>("t_thue")
                             };
                if (source != null)
                {
                    Decimal? nullable = source.Sum(p => p.thue_nt);
                    Decimal.TryParse(nullable.ToString(), out result1);
                    nullable = source.Sum(p => p.thue);
                    Decimal.TryParse(nullable.ToString(), out result2);
                }
                Decimal num3 = FNum.ToDec(this.PhView[0]["t_tien_nt"]) + result1;
                Decimal num4 = FNum.ToDec(this.PhView[0]["t_tien"]) + result2;
                this.PhView[0]["t_thue_nt"] = (object)result1;
                this.PhView[0]["t_thue"] = (object)result2;
                this.PhView[0]["t_tt_nt"] = (object)num3;
                this.PhView[0]["t_tt"] = (object)num4;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtTy_gia_ht_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtTy_gia_ht.Value == DBNull.Value)
                this.txtTy_gia_ht.Value = (object)0;
            Decimal nValue1 = this.txtTy_gia.nValue;
            Decimal nValue2 = this.txtTy_gia_ht.nValue;
            Decimal num1;
            Decimal num2;
            Decimal num3;
            switch (FrmCACTPC1.Ma_GD_Value.Text)
            {
                case "8":
                    using (IEnumerator<Record> enumerator = ((IEnumerable<Record>)this.GrdCt.Records).GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            DataRecord current = (DataRecord)enumerator.Current;
                            num1 = new Decimal(0);
                            num2 = new Decimal(0);
                            Decimal num4 = new Decimal(0);
                            Decimal num5 = new Decimal(0);
                            num3 = new Decimal(0);
                            Decimal num6 = Convert.ToDecimal(current.Cells["tien_nt"].Value);
                            Decimal num7 = SysFunc.Round(nValue2 * num6, StartUpTrans.M_ROUND);
                            Decimal num8 = Convert.ToDecimal(current.Cells["thue_nt"].Value);
                            Decimal num9 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                            Decimal num10 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round((num6 + num8) * nValue2, this.M_Round) : SysFunc.Round((num6 + num8) * nValue1, this.M_Round);
                            current.Cells["tien"].Value = (object)num7;
                            current.Cells["tien_tt"].Value = (object)num10;
                            current.Cells["thue"].Value = (object)num9;
                            current.Cells["tt"].Value = (object)(num7 + num9);
                            this.Tinh_tien_cltg(current);
                        }
                        break;
                    }
                case "9":
                    using (IEnumerator<Record> enumerator = ((IEnumerable<Record>)this.GrdCtChi.Records).GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            DataRecord current = (DataRecord)enumerator.Current;
                            num1 = new Decimal(0);
                            num2 = new Decimal(0);
                            num3 = new Decimal(0);
                            Decimal num4 = Convert.ToDecimal(current.Cells["tien_nt"].Value);
                            Decimal num5 = SysFunc.Round(nValue2 * num4, StartUpTrans.M_ROUND);
                            Decimal num6 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num4 * nValue2, this.M_Round) : SysFunc.Round(num4 * nValue1, this.M_Round);
                            current.Cells["tien"].Value = (object)num5;
                            current.Cells["tien_tt"].Value = (object)num6;
                            current.Cells["tt"].Value = (object)num5;
                            this.Tinh_tien_cltg(current);
                        }
                        break;
                    }
                default:
                    DataGridView dataGridView = this.GrdCtChi;
                    if (this.Ma_gd == "1")
                        dataGridView = this.Grdhd;
                    using (IEnumerator<Record> enumerator = ((IEnumerable<Record>)dataGridView.Records).GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            DataRecord current = (DataRecord)enumerator.Current;
                            num1 = new Decimal(0);
                            num3 = new Decimal(0);
                            Decimal num4 = Convert.ToDecimal(current.Cells["tien_nt"].Value);
                            Decimal num5 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num4 * nValue2, this.M_Round) : SysFunc.Round(num4 * nValue1, this.M_Round);
                            current.Cells["tien_tt"].Value = (object)num5;
                            this.Tinh_tien_cltg(current);
                        }
                        break;
                    }
            }
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtTy_gia.Value == DBNull.Value)
                this.txtTy_gia.Value = (object)0;
            try
            {
                if (FormTrans.currActionTask == ActionTask.View || !(this.txtTy_gia.OldValue != this.txtTy_gia.nValue))
                    return;
                this.CalculateTyGia();
                this.Ty_Gia_ValueChange.Value = !this.Ty_Gia_ValueChange.Value;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void CalculateTyGia()
        {
            Decimal nValue1 = this.txtTy_gia.nValue;
            if (this.cbMa_nt.Text.Trim().Equals(this.M_Ma_nt0))
                this.txtTy_gia_ht.Value = this.txtTy_gia.Value;
            Decimal nValue2 = this.txtTy_gia_ht.nValue;
            if (this.Sua_tien == new Decimal(1))
                return;
            Decimal num1;
            Decimal num2;
            Decimal num3;
            Decimal num4;
            if (FrmCACTPC1.Ma_GD_Value.Text.Trim().ToString().IndexOfAny(new char[2]
            {
        '2',
        '3'
            }) >= 0)
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtChi.Records)
                {
                    num1 = new Decimal(0);
                    num2 = new Decimal(0);
                    Decimal num5 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num6 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num5 * nValue2, this.M_Round) : SysFunc.Round(num5 * nValue1, this.M_Round);
                    if (!this.PhView[0]["ma_nt"].ToString().ToUpper().Trim().Equals(StartUpTrans.M_ma_nt0.ToUpper().Trim()))
                    {
                        record.Cells["tien_tt"].Value = (object)num6;
                        this.Tinh_tien_cltg(record);
                    }
                }
            }
            else if (FrmCACTPC1.Ma_GD_Value.Text.Equals("9"))
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtChi.Records)
                {
                    num1 = new Decimal(0);
                    num2 = new Decimal(0);
                    num3 = new Decimal(0);
                    Decimal num5 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num6 = SysFunc.Round(num5 * nValue2, this.M_Round);
                    Decimal num7 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num5 * nValue2, this.M_Round) : SysFunc.Round(num5 * nValue1, this.M_Round);
                    record.Cells["tien"].Value = (object)num6;
                    record.Cells["tien_tt"].Value = (object)num7;
                    record.Cells["tt"].Value = (object)num7;
                    this.Tinh_tien_cltg(record);
                }
            }
            else if (FrmCACTPC1.Ma_GD_Value.Text.Equals("8"))
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCt.Records)
                {
                    if (record.Cells["loai_hd"].Value.ToString().Trim().Equals("2"))
                    {
                        this.CalculateHd2(record);
                    }
                    else
                    {
                        num1 = new Decimal(0);
                        num3 = new Decimal(0);
                        num2 = new Decimal(0);
                        Decimal num5 = new Decimal(0);
                        Decimal num6 = new Decimal(0);
                        Decimal num7 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                        Decimal num8 = Convert.ToDecimal(record.Cells["thue_nt"].Value);
                        Decimal num9 = Convert.ToDecimal(record.Cells["thue"].Value);
                        Decimal num10 = SysFunc.Round(nValue2 * num7, StartUpTrans.M_ROUND);
                        record.Cells["tien"].Value = (object)num10;
                        if (FNum.ToDec(this.GetPhValue("sua_thue")) == new Decimal(0) && !string.IsNullOrEmpty(record.Cells["thue_suat"].Value.ToString()))
                        {
                            num4 = new Decimal(0);
                            Decimal num11 = Convert.ToDecimal(record.Cells["thue_suat"].Value);
                            num8 = SysFunc.Round(num7 * num11 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            num9 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                            record.Cells["thue_nt"].Value = (object)num8;
                            record.Cells["thue"].Value = (object)num9;
                        }
                        Decimal num12 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round((num7 + num8) * nValue2, this.M_Round) : SysFunc.Round((num7 + num8) * nValue1, this.M_Round);
                        record.Cells["tt_nt"].Value = (object)(num7 + num8);
                        record.Cells["tt"].Value = (object)(num10 + num9);
                        record.Cells["tien_tt"].Value = (object)num12;
                        this.Tinh_tien_cltg(record);
                    }
                }
            }
            else if (FrmCACTPC1.Ma_GD_Value.Text.Equals("1"))
            {
                num1 = new Decimal(0);
                num2 = new Decimal(0);
                num3 = new Decimal(0);
                Decimal num5 = new Decimal(0);
                foreach (DataRecord record in (IEnumerable<Record>)this.Grdhd.Records)
                {
                    Decimal num6 = Convert.ToDecimal(record.Cells["tien_nt"].Value);
                    Decimal num7 = Convert.ToDecimal((record.DataItem as DataRowView)["ty_gia_ht2"]);
                    Decimal num8 = SysFunc.Round(num6 * num7, this.M_Round);
                    Decimal num9 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num6 * nValue2, this.M_Round) : SysFunc.Round(num6 * nValue1, this.M_Round);
                    record.Cells["tien"].Value = (object)num8;
                    record.Cells["tien_tt"].Value = (object)num9;
                    this.Tinh_tien_cltg(record);
                }
            }
            this.UpdateTotalHT();
            if (this.GrdCtgt.Records.Count > 0)
            {
                foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtgt.Records)
                {
                    if (record.Cells["t_tien_nt"].Value != DBNull.Value)
                    {
                        num1 = new Decimal(0);
                        Decimal num5 = new Decimal(0);
                        Decimal num6 = new Decimal(0);
                        Decimal num7 = new Decimal(0);
                        Decimal num8 = Convert.ToDecimal(record.Cells["t_tien_nt"].Value);
                        Decimal num9 = Convert.ToDecimal(record.Cells["t_thue_nt"].Value);
                        Decimal num10 = Convert.ToDecimal(record.Cells["t_thue"].Value);
                        Decimal num11 = SysFunc.Round(nValue2 * num8, StartUpTrans.M_ROUND);
                        record.Cells["t_tien"].Value = (object)num11;
                        if (FNum.ToDec(this.GetPhValue("sua_thue")) == new Decimal(0) && !string.IsNullOrEmpty(record.Cells["thue_suat"].Value.ToString()))
                        {
                            num4 = new Decimal(0);
                            Decimal num12 = Convert.ToDecimal(record.Cells["thue_suat"].Value);
                            num9 = SysFunc.Round(num8 * num12 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            num10 = SysFunc.Round(num9 * nValue2, StartUpTrans.M_ROUND);
                            record.Cells["t_thue_nt"].Value = (object)num9;
                            record.Cells["t_thue"].Value = (object)num10;
                        }
                        record.Cells["t_tt_nt"].Value = (object)(num8 + num9);
                        record.Cells["t_tt"].Value = (object)(num11 + num10);
                    }
                }
            }
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value == DBNull.Value)
                this.txtNgay_ct.Value = (object)DateTime.Now;
            if (this.txtNgay_ct.IsFocusWithin || FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy || (!StartUpTrans.M_ngay_lct.Equals("0") && !(this.txtngay_lct.dValue == new DateTime()) || !(this.txtNgay_ct.dValue != new DateTime())))
                return;
            this.txtngay_lct.Value = (object)this.txtNgay_ct.dValue.Date;
        }

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
        }

        private void GrdCtChi_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            if (FrmCACTPC1.Ma_GD_Value.Text.Equals("4"))
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
            else
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void FormMain_Closed(object sender, EventArgs e)
        {
            if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                return;
            Application.Current.Shutdown();
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void TabInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.PhView[0]["ma_gd"] != DBNull.Value && !string.IsNullOrEmpty(this.PhView[0]["ma_gd"].ToString()) && this.PhView[0]["ma_gd"].ToString().Equals("8"))
            {
                if (this.TabInfo.SelectedIndex == 1)
                {
                    if (FrmCACTPC1.IsInEditMode != null && FrmCACTPC1.IsInEditMode.Value)
                    {
                        FrmCACTPC1.IsInEditModeThue.Value = true;
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec= '" + this.PhData.Rows[FrmCACTPC1.iRow]["stt_rec"].ToString() + "' AND thue_ct = '1'"))
                            StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                        IEnumerable<Record> source = this.GrdCt.Records.Where<Record>((Func<Record, bool>)(x =>
                       {
                           DataRowView dataItem = (x as DataRecord).DataItem as DataRowView;
                           return !dataItem["ngay_ct0"].Equals((object)DBNull.Value) && !string.IsNullOrEmpty(dataItem["ngay_ct0"].ToString()) && !dataItem["loai_hd"].ToString().Trim().Equals("0");
                       }));
                        if (source.Any<Record>())
                        {
                            foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec= '" + this.PhData.Rows[FrmCACTPC1.iRow]["stt_rec"].ToString() + "'"))
                                StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                        }
                        foreach (Record record in source)
                        {
                            DataRowView dataItem = (record as DataRecord).DataItem as DataRowView;
                            if (!dataItem["ngay_ct0"].Equals((object)DBNull.Value) && !string.IsNullOrEmpty(dataItem["ngay_ct0"].ToString()) && !dataItem["loai_hd"].ToString().Trim().Equals("0"))
                            {
                                FrmCACTPC1.IsInEditModeThue.Value = false;
                                DataRow row = StartUpTrans.DsTrans.Tables[2].NewRow();
                                row["stt_rec"] = this.PhView[0]["stt_rec"];
                                row["thue_ct"] = (object)"1";
                                row["stt_rec0"] = dataItem["stt_rec0"];
                                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                                row["ngay_ct"] = this.txtNgay_ct.Value;
                                row["ma_ms"] = dataItem["ma_ms"];
                                row["so_ct0"] = dataItem["so_ct0"];
                                row["so_seri0"] = dataItem["so_seri0"];
                                row["kh_mau_hd"] = dataItem["kh_mau_hd"];
                                row["ngay_ct0"] = dataItem["ngay_ct0"];
                                row["ma_kh"] = dataItem["ma_kh_t"];
                                row["ten_kh"] = dataItem["ten_kh_t"];
                                row["dia_chi"] = dataItem["dia_chi_t"];
                                row["ma_so_thue"] = dataItem["mst_t"];
                                row["ten_vt"] = dataItem["ten_vt_t"];
                                row["t_tien_nt"] = dataItem["tien_nt"];
                                row["t_tien"] = dataItem["tien"];
                                row["ma_thue"] = dataItem["ma_thue_i"];
                                row["thue_suat"] = dataItem["thue_suat"];
                                row["t_thue_nt"] = dataItem["thue_nt"];
                                row["t_thue"] = dataItem["thue"];
                                row["t_tt_nt"] = dataItem["tt_nt"];
                                row["t_tt"] = dataItem["tt"];
                                row["tk_thue_no"] = dataItem["tk_thue_i"];
                                row["tk_thue_cn"] = dataItem["tk_thue_cn"];
                                row["ghi_chu"] = dataItem["ghi_chu_t"];
                                row["ma_vv"] = dataItem["ma_vv_i"];
                                row["ma_phi"] = dataItem["ma_phi_i"];
                                if (this.GrdCtgt.DefaultFieldLayout.Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_td4")))
                                    row["ma_td4"] = dataItem["ma_td4_i"];
                                StartUpTrans.DsTrans.Tables[2].Rows.Add(row);
                            }
                        }
                        if (this.GrdCtgt.Records.Count == 0)
                            this.GrdCtgt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    this.UpdateTotalThue();
                }
                this.GrdCtgt.FieldSettings.AllowEdit = new bool?(true);
            }
            else
            {
                if (FrmCACTPC1.IsInEditModeThue != null)
                    FrmCACTPC1.IsInEditModeThue.Value = false;
                this.GrdCtgt.FieldSettings.AllowEdit = new bool?(false);
            }
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0.Text = this.PhView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = this.PhView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Loai_tg.Text = this.PhView[0]["loai_tg"].ToString();
            FrmCACTPC1.Ma_GD_Value.Text = this.txtMa_gd.Text;
            FrmCACTPC1.Ma_GD_Value.Value = this.txtMa_gd.Text.Equals("8");
            if (this.txtMa_gd.Text.Equals("3"))
            {
                this.txtMa_kh.AllowEmty = true;
            }
            else
                this.txtMa_kh.AllowEmty = false;
            this.SetStatusVisibleField();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.TabInfo_SelectionChanged((object)null, (SelectionChangedEventArgs)null)));
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            ExRateTextBox exRateTextBox = sender as ExRateTextBox;
            if (exRateTextBox.IsReadOnly)
            {
                KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
            }
            else
            {
                if (sender == this.txtTy_gia_ht && this.txtTy_gia_ht.nValue == new Decimal(0))
                    this.txtTy_gia_ht.Value = (object)this.GetTggd(this.Ma_nt, this.Ngay_ct);
                exRateTextBox.SelectAll();
            }
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(this.PhView[0]["ma_qs"].ToString()) || !string.IsNullOrEmpty(this.PhView[0]["so_ct"].ToString().Trim()))
                return;
            if (string.IsNullOrEmpty(this.PhView[0]["so_cttmp"].ToString().Trim()) || !this.PhView[0]["ma_qs"].ToString().Trim().Equals(this.PhView[0]["ma_qstmp"].ToString().Trim()))
            {
                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                this.PhView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                this.PhView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
            }
            else
                this.txtSo_ct.Text = this.PhView[0]["so_cttmp"].ToString().Trim();
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

        private void txtngay_lct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtngay_lct.IsFocusWithin || (FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy || this.txtNgay_ct.dValue.Date.Equals(this.txtngay_lct.dValue.Date)))
                return;
            int num = (int)ExMessageBox.Show(770, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtMa_gd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_gd.RowResult != null && !string.IsNullOrEmpty(this.txtMa_gd.Text.Trim()))
                this.txtTen_gd.Text = !this.M_LAN.ToUpper().Equals("V") ? this.txtMa_gd.RowResult["ten_gd2"].ToString() : this.txtMa_gd.RowResult["ten_gd"].ToString();
            FrmCACTPC1.Ma_GD_Value.Text = this.txtMa_gd.Text;
            if (this.txtMa_gd.Text.Equals("8"))
            {
                FrmCACTPC1.Ma_GD_Value.Value = true;
            }
            else
            {
                FrmCACTPC1.Ma_GD_Value.Value = false;
                this.PhView[0]["sua_thue"] = (object)0;
            }
            this.SetStatusVisibleField();
            if (FrmCACTPC1.IsInEditMode != null && !FrmCACTPC1.IsInEditMode.Value)
                return;
            FrmCACTPC1.IsInEditModeThue.Value = this.txtMa_gd.Text.Equals("8");
            this.GrdCtgt.FieldSettings.AllowEdit = new bool?(FrmCACTPC1.IsInEditModeThue.Value);
            if (this.txtMa_gd.Text.Equals("3"))
            {
                this.txtMa_kh.AllowEmty = true;                
            }
            else
                this.txtMa_kh.AllowEmty = false;
        }

        private void txtHan_tt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.Return))
                return;
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            e.Handled = true;
        }

        private void GrdCt_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Add && (this.CtView.Count == 1 && this.CtView[0]["dien_giaii"].ToString() == string.Empty))
                this.CtView[0]["dien_giaii"] = this.PhView[0].Row["dien_giai"];
            if (FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit || !(this.txtMa_gd.Text == "3"))
                return;
            foreach (DataRecord record in (IEnumerable<Record>)this.GrdCtChi.Records)
            {
                record.Cells["ty_giahtf2"].Value = this.txtTy_gia.Value;
                Decimal nValue = this.txtTy_gia.nValue;
                Decimal num = this.mLoai_tg == new Decimal(1) ? nValue : (nValue == new Decimal(0) ? new Decimal(0) : SysFunc.Round(new Decimal(1) / nValue, this.M_ROUND_TY_GIA));
                (record.DataItem as DataRowView)["ty_gia_ht2"] = (object)num;
            }
        }

        private void btnSoHD_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsEditMode)
            {
                int num1 = (int)ExMessageBox.Show(775, StartupBase.SasObj, "Phải lưu chứng từ rồi mới phân bổ cho các hóa đơn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (this.PhView[0]["status"].ToString() != "2")
            {
                int num2 = (int)ExMessageBox.Show(780, StartupBase.SasObj, "Phải ghi vào sổ cái rồi mới phân bổ cho các hđ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                string sttRec = this.Stt_rec;
                object obj = this.PhData.Rows[FrmCACTPC1.iRow]["ngay_ct"];
                PbInfo pbInfo = new PbInfo(obj, obj, (object)"", "", "");
                pbInfo.TitleView = "Phan bo";
                Apttpb.StartUp.Procedure = StartUpTrans.CommandInfo["parameter"].ToString().Split(';');
                new Apttpb.StartUp().Pb_tt(sttRec, (IPhanbo)pbInfo);
                this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
               {
                   SqlCommand sqlcmd = new SqlCommand("SELECT so_ct_tt FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec = @Stt_rec");
                   sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = (object)this.Stt_rec;
                   DataSet dataSet = this.SasObj.ExcuteReader(sqlcmd);
                   if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                       this.lblSo_ct_tt.Text = dataSet.Tables[0].Rows[0][0].ToString();
                   else
                       this.lblSo_ct_tt.Text = "";
               }), DispatcherPriority.Background);
            }
        }

        private void Grdhd_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                DataRecord rec = e.Cell.Record;
                if (FrmCACTPC1.IsInEditMode.Value && this.Grdhd.ActiveCell != null && this.CtView.Count > this.Grdhd.ActiveRecord.Index && this.CtData.GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "tk_i":
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl.RowResult != null)
                            {
                                if (this.M_LAN.ToUpper().Equals("V"))
                                    e.Cell.Record.Cells["ten_tk"].Value = autoCompleteControl.RowResult["ten_tk"];
                                else
                                    e.Cell.Record.Cells["ten_tk2"].Value = autoCompleteControl.RowResult["ten_tk2"];
                                break;
                            }
                            break;
                        case "tien_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result = new Decimal(0);
                                Decimal num1 = new Decimal(0);
                                Decimal num2 = new Decimal(0);
                                Decimal num3 = new Decimal(0);
                                Decimal num4 = new Decimal(0);
                                if (!string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                    num4 = (e.Editor as NumericTextBox).nValue;
                                Decimal.TryParse((e.Cell.Record.DataItem as DataRowView)["ty_gia_ht2"].ToString(), out result);
                                Decimal num5 = SysFunc.Round(num4 * result, this.M_Round);
                                Decimal nValue1 = this.txtTy_gia.nValue;
                                Decimal nValue2 = this.txtTy_gia_ht.nValue;
                                Decimal num6 = !(this.PhView[0]["loai_cl_no"].ToString() == "0") ? SysFunc.Round(num4 * nValue2, this.M_Round) : SysFunc.Round(num4 * nValue1, this.M_Round);
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                    e.Cell.Record.Cells["tien_tt"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                }
                                else
                                {
                                    e.Cell.Record.Cells["tien"].Value = (object)num5;
                                    e.Cell.Record.Cells["tien_tt"].Value = (object)num6;
                                }
                                DataRowView dataItem = e.Cell.Record.DataItem as DataRowView;
                                dataItem["tt_nt"] = e.Editor.Value;
                                dataItem["tt"] = e.Cell.Record.Cells["tien"].Value;
                                this.Tinh_tien_cltg(e.Cell.Record);
                                e.Cell.Record.Cells["tien_con_pt"].Value = (object)(this.Parsedecimal(e.Cell.Record.Cells["tien_con_pt0"].Value, new Decimal(0)) - (e.Editor as NumericTextBox).nValue);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tien":
                            if (e.Cell.IsDataChanged)
                            {
                                (e.Cell.Record.DataItem as DataRowView)["tt"] = e.Cell.Record.Cells["tien"].Value;
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tien_tt":
                            if (e.Cell.IsDataChanged)
                            {
                                this.Tinh_tien_cltg(e.Cell.Record);
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "dien_giaii":
                            if (e.Cell.Record.Index == 0 && string.IsNullOrEmpty(e.Editor.Text.Trim()))
                            {
                                e.Cell.Value = (object)this.txtDien_giai.Text;
                                break;
                            }
                            break;
                        case "so_ct0":
                            DataRecord activeRecord1 = this.Grdhd.ActiveRecord as DataRecord;
                            if ((activeRecord1.DataItem as DataRowView)["state"].ToString().Trim() == "1")
                                break;
                            if ((e.Editor.Value is DBNull || string.IsNullOrEmpty(e.Editor.Value.ToString().Trim())) && ExMessageBox.Show(795, StartupBase.SasObj, "Có nhập tiếp không?", StartupBase.SasObj.GetSysvar("M_FAST_VER").ToString(), MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                            {
                                if (activeRecord1 != null)
                                {
                                    int num1 = 0;
                                    Cell activeCell = this.Grdhd.ActiveCell;
                                    num1 = activeRecord1.Index;
                                    if (activeRecord1.Index == 0)
                                    {
                                        if (this.Grdhd.Records.Count == 1)
                                            this.Grdhd_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                                    }
                                    else if (activeRecord1.Index == this.Grdhd.Records.Count - 1)
                                        num1 = activeRecord1.Index - 1;
                                    int num2 = this.Grdhd.ActiveCell == null ? 0 : this.Grdhd.ActiveCell.Field.Index;
                                    this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                                    if (num2 >= 0)
                                    {
                                        this.CtData.Rows.Remove(this.CtView[activeRecord1.Index].Row);
                                        this.CtData.AcceptChanges();
                                        if (this.Grdhd.Records.Count > 0)
                                            this.Grdhd.ActiveRecord = this.Grdhd.Records[this.Grdhd.Records.Count - 1];
                                        this.UpdateTotalHT();
                                    }
                                }
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
                                break;
                            }
                            if (this.isCurrentFormActive)
                                break;
                            FrmCACTPC1DSHD frmCactpC1Dshd = new FrmCACTPC1DSHD();
                            frmCactpC1Dshd.ShowHd(e.Editor.Value.ToString(), this.stt_rec_magd1);
                            if (frmCactpC1Dshd.grdDSHD.ActiveRecord is DataRecord activeRecord && activeRecord.RecordType == RecordType.DataRecord)
                            {
                                Decimal num = new Decimal(0);
                                string empty = string.Empty;
                                Decimal dec = FNum.ToDec(activeRecord.Cells["t_tien_dt"].Value);
                                string str = activeRecord.Cells["stt_rec"].Value.ToString();
                                this.stt_rec_magd1 = str;
                                DataRow rateCl = StartUp.GetRateCL(this.cbMa_nt.Text.Trim(), (DateTime)activeRecord.Cells["ngay_ct0"].Value, this.txtNgay_ct.dValue);
                                rec.Cells["so_ct0"].Value = (object)activeRecord.Cells["so_ct0"].Value.ToString();
                                rec.Cells["ngay_ct0"].Value = activeRecord.Cells["ngay_ct0"].Value;
                                rec.Cells["tk_i"].Value = (object)activeRecord.Cells["tk"].Value.ToString();
                                rec.Cells["ma_nt_i"].Value = (object)activeRecord.Cells["ma_nt"].Value.ToString();
                                rec.Cells["tien_hd"].Value = activeRecord.Cells["tien_hd"].Value;
                                rec.Cells["t_tien_dt"].Value = (object)dec;
                                if (rateCl == null)
                                {
                                    (rec.DataItem as DataRowView)["ty_giahtf2"] = (activeRecord.DataItem as DataRowView)["ty_giaf"];
                                    (rec.DataItem as DataRowView)["ty_gia_ht2"] = (activeRecord.DataItem as DataRowView)["ty_gia"];
                                }
                                else
                                {
                                    (rec.DataItem as DataRowView)["ty_giahtf2"] = rateCl["ty_giaf"];
                                    (rec.DataItem as DataRowView)["ty_gia_ht2"] = rateCl["ty_gia"];
                                }
                                rec.Cells["tien_con_pt0"].Value = (object)(FNum.ToDec(rec.Cells["tien_hd"].Value) - FNum.ToDec(rec.Cells["t_tien_dt"].Value));
                                rec.Cells["tien_con_pt"].Value = (object)(FNum.ToDec(rec.Cells["tien_con_pt0"].Value) - FNum.ToDec(rec.Cells["tien_nt"].Value));
                                rec.Cells["stt_rec_tt"].Value = (object)str;
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_vv_i")))
                                    rec.Cells["ma_vv_i"].Value = (object)activeRecord.Cells["ma_vv"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_td_i")))
                                    rec.Cells["ma_td_i"].Value = (object)activeRecord.Cells["ma_td"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_td2_i")))
                                    rec.Cells["ma_td2_i"].Value = (object)activeRecord.Cells["ma_td2"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_td3_i")))
                                    rec.Cells["ma_td3_i"].Value = (object)activeRecord.Cells["ma_td3"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "so_lsx_i")))
                                    rec.Cells["so_lsx_i"].Value = (object)activeRecord.Cells["so_lsx"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "so_dh_i")))
                                    rec.Cells["so_dh_i"].Value = (object)activeRecord.Cells["so_dh"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_bpht_i")))
                                    rec.Cells["ma_bpht_i"].Value = (object)activeRecord.Cells["ma_bpht"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_hd_i")))
                                    rec.Cells["ma_hd_i"].Value = (object)activeRecord.Cells["ma_hd"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_ku_i")))
                                    rec.Cells["ma_ku_i"].Value = (object)activeRecord.Cells["ma_ku"].Value.ToString();
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_phi_i")))
                                    rec.Cells["ma_phi_i"].Value = activeRecord.Cells["ma_phi"].Value;
                                if (rec.Cells.Any<Cell>((Func<Cell, bool>)(x => x.Field.Name == "ma_sp")))
                                    rec.Cells["ma_sp"].Value = activeRecord.Cells["ma_sp"].Value;
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.Grdhd.ActiveCell = rec.Cells["tien_nt"]));
                                break;
                            }
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                           {
                               this.Grdhd.ActiveCell = e.Cell;
                               this.Grdhd.ExecuteCommand(DataPresenterCommands.StartEditMode);
                           }));
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void btnPO_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsEditMode)
            {

                DateTime endDate = DateTime.Today;
                DateTime startDate = endDate.AddMonths(-1);
                SqlCommand sqlcmd = new SqlCommand("Exec [POBK1_PC] @StartDate, @EndDate, @Condition, @LoaiPhieuNhap");
                sqlcmd.Parameters.Add("@StartDate", SqlDbType.VarChar, 8).Value =
                    startDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);

                sqlcmd.Parameters.Add("@EndDate", SqlDbType.VarChar, 8).Value =
                    endDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
                sqlcmd.Parameters.Add("@Condition", SqlDbType.NVarChar).Value = (object)("and nxt = 1  and ma_dvcs like '" + this.Ma_dvcs + "%'");
                sqlcmd.Parameters.Add("@LoaiPhieuNhap", SqlDbType.VarChar).Value = (object)"331,341,111,141";
                StartUp.HDBData = StartupBase.SasObj.ExcuteReader(sqlcmd);

                DataTable dataTable = StartUp.HDBData.Tables[0];

                COTKTH2Dvcs Frmlocdon = new COTKTH2Dvcs();
                Frmlocdon.GrdCt.DataSource = (IEnumerable)dataTable.DefaultView;
                if (StartupBase.M_LAN != "V")
                    Frmlocdon.Title = "Supplier list";
                bool? nullable = Frmlocdon.ShowDialog();

                if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                    return;
                StartUp.PO = dataTable.Select("tag = 1");
                if (StartUp.PO != null && StartUp.PO.Length > 0)
                {
                    foreach (DataRow dataRow in StartUp.PO)
                    {
                        this.NewRowCtPO(dataRow);
                    }
                    foreach (DataRow blankRow in StartUpTrans.DsTrans.Tables[1].Select("so_ct0 IS NULL"))
                    {
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(blankRow);
                    }
                }
            }
        }
        private bool NewRowCtPO(DataRow data)
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                int result = 1;
                string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                {
                    int.TryParse(str.ToString(), out result);
                    ++result;
                }
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)result);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ngay_ct"] = this.txtNgay_ct.Value == null ? (object)DateTime.Now.Date : this.txtNgay_ct.Value;
                //if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                //    dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[StartUpTrans.DsTrans.Tables[1].DefaultView.Count - 1]["ma_kho_i"];
                dataRow["gia_nt"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["tien_con_pt0"] = (object)0;
                dataRow["tien_hd"] = (object)0;
                dataRow["t_tien_dt"] = (object)0;

                dataRow["tien_nt"] = data["t_tt"] == DBNull.Value
                    ? (object)0m
                    : data["t_tt"];
                dataRow["tien_tt"]=(object)0;
                dataRow["tt_nt"]=(object)0;
                dataRow["thue"] = (object)0;

                dataRow["thue_nt"] = (object)0;

                dataRow["ma_hdm_i"] = (object)data["so_ct"];
                dataRow["so_ct0"] = (object)data["so_ct"];
                dataRow["ty_gia_ht2"] =(object)0;

                dataRow["tien"] = (object)0;

                txtMa_kh.Text = data["ma_kh"] == DBNull.Value ? "" : data["ma_kh"].ToString();

                DataSet dsKh = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("SELECT dia_chi, ma_so_thue FROM dmkh WHERE ma_kh = '{0}'", txtMa_kh.Text.Trim())));
                if (dsKh != null && dsKh.Tables.Count > 0 && dsKh.Tables[0].Rows.Count > 0)
                {
                    txtDia_chi.Text = dsKh.Tables[0].Rows[0]["dia_chi"] == DBNull.Value ? "" : dsKh.Tables[0].Rows[0]["dia_chi"].ToString();
                    txtMaSoThue.Text = dsKh.Tables[0].Rows[0]["ma_so_thue"] == DBNull.Value ? "" : dsKh.Tables[0].Rows[0]["ma_so_thue"].ToString();
                }



                //dataRow["ton13"] = (object)0;
                //dataRow["ma_vt"] = (object)data["ma_vt"];
                //dataRow["ten_vt"] = (object)data["ten_vt"];
                //dataRow["dvt"] = (object)data["dvt"];
                //dataRow["ma_kho"] = (object)data["ma_kho"];
                //dataRow["so_luong"] = (object)data["ton_cuoi"];



                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
                return true;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
                return false;
            }
        }

        private void Grdhd_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Add || (this.CtView.Count != 1 || !(this.CtView[0]["dien_giaii"].ToString() == string.Empty)))
                return;
            this.CtView[0]["dien_giaii"] = this.PhView[0].Row["dien_giai"];
        }

        private bool Grdhd_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCt();
            return true;
        }

        private void Grdhd_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()), DispatcherPriority.Background);
        }

        private void Grdhd_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                this.NewRowCt();
                this.Grdhd.ActiveRecord = this.Grdhd.Records[this.Grdhd.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                return;
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
        }

        private void Grdhd_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtChi.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete") && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
                        break;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.Grdhd.ActiveRecord = this.Grdhd.Records[this.Grdhd.Records.Count - 1];
                            this.Grdhd.ActiveCell = (this.Grdhd.ActiveRecord as DataRecord).Cells["so_ct0"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.Grdhd, "so_ct0");
                            break;
                    }
                    break;
                case Key.F5:
                    if (StartupBase.SasObj.VersionInfo.Rows[0]["product_code"].ToString().Equals("FA") || StartUp.dtRegInfo != null && !StartUp.dtRegInfo.Rows[18]["content"].ToString().Trim().Equals("FK"))
                    {
                        Catinhtg.Tinh(this.GrdCtChi.Records, (IPhValue)this);
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(800, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.Grdhd.ActiveRecord is DataRecord activeRecord))
                        break;
                    (activeRecord.DataItem as DataRowView)["state"] = (object)1;
                    Cell activeCell = this.Grdhd.ActiveCell;
                    int num1 = activeRecord.Index;
                    if (activeRecord.Index == 0)
                    {
                        if (this.Grdhd.Records.Count == 1)
                            this.Grdhd_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord.Index == this.Grdhd.Records.Count - 1)
                        num1 = activeRecord.Index - 1;
                    int num2 = this.Grdhd.ActiveCell == null ? 0 : this.Grdhd.ActiveCell.Field.Index;
                    this.Grdhd.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        this.CtData.Rows.Remove(this.CtView[activeRecord.Index].Row);
                        this.CtData.AcceptChanges();
                        if (this.Grdhd.Records.Count > 0)
                            this.Grdhd.ActiveRecord = this.Grdhd.Records[num1 > this.Grdhd.Records.Count - 1 ? this.Grdhd.Records.Count - 1 : num1];
                        this.UpdateTotalHT();
                    }
                    break;
            }
        }

        private bool isEquals(Decimal a, Decimal b, string stt_rec)
        {
            return stt_rec != this.Stt_rec || a == b;
        }

        public Decimal Sua_thue
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("sua_thue"));
            }
            set
            {
                this.SetPhValue("sua_thue", (object)value);
            }
        }

        public SasObject SasObj
        {
            get
            {
                return StartupBase.SasObj;
            }
        }

        public int M_Round
        {
            get
            {
                return StartUpTrans.M_ROUND;
            }
        }

        public int M_Round_nt
        {
            get
            {
                return StartUpTrans.M_ROUND_NT;
            }
        }

        public int M_ROUND_TY_GIA
        {
            get
            {
                return (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar(nameof(M_ROUND_TY_GIA)));
            }
        }

        public string M_Ma_nt0
        {
            get
            {
                return StartUpTrans.M_ma_nt0;
            }
        }

        public Decimal Sua_tien
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("sua_tien"));
            }
            set
            {
                this.SetPhValue("sua_tien", (object)value);
            }
        }

        public Decimal Sua_tggs
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("sua_tggs"));
            }
            set
            {
                this.SetPhValue("sua_tggs", (object)value);
            }
        }

        public Decimal Ty_giaf
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("ty_giaf"));
            }
            set
            {
                this.SetPhValue("ty_giaf", (object)value);
            }
        }

        public Decimal Ty_gia_htf
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("ty_gia_htf"));
            }
            set
            {
                this.SetPhValue("ty_gia_htf", (object)value);
            }
        }

        public Decimal Ty_gia
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("ty_gia"));
            }
            set
            {
                this.SetPhValue("ty_gia", (object)value);
            }
        }

        public Decimal Ty_gia_ht
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("ty_gia_ht"));
            }
            set
            {
                this.SetPhValue("ty_gia_ht", (object)value);
            }
        }

        public Decimal mLoai_tg
        {
            get
            {
                return FNum.ToDec(this.GetPhValue("loai_tg"));
            }
            set
            {
                this.SetPhValue("loai_tg", (object)value);
            }
        }

        public DateTime Ngay_ct
        {
            get
            {
                return FDate.ToDate(this.GetPhValue("ngay_ct"));
            }
            set
            {
                this.SetPhValue("ngay_ct", (object)value);
            }
        }

        public string Stt_rec
        {
            get
            {
                return this.GetPhValue("stt_rec").ToString();
            }
            set
            {
                this.SetPhValue("stt_rec", (object)value);
            }
        }

        public string Ma_ct
        {
            get
            {
                return this.GetPhValue("ma_ct").ToString();
            }
            set
            {
                this.SetPhValue("ma_ct", (object)value);
            }
        }

        public string Ma_gd
        {
            get
            {
                return this.GetPhValue("ma_gd").ToString();
            }
            set
            {
                this.SetPhValue("ma_gd", (object)value);
            }
        }

        public string Ma_nt
        {
            get
            {
                return this.GetPhValue("ma_nt").ToString();
            }
            set
            {
                this.SetPhValue("ma_nt", (object)value);
            }
        }

        public string Ma_kh
        {
            get
            {
                return this.GetPhValue("ma_kh").ToString();
            }
            set
            {
                this.SetPhValue("ma_kh", (object)value);
            }
        }

        public string Ma_dvcs
        {
            get
            {
                return this.GetPhValue("ma_dvcs").ToString();
            }
            set
            {
                this.SetPhValue("ma_dvcs", (object)value);
            }
        }

        public string Tk
        {
            get
            {
                return this.GetPhValue("tk").ToString();
            }
            set
            {
                this.SetPhValue("tk", (object)value);
            }
        }

        public string So_ct
        {
            get
            {
                return this.GetPhValue("so_ct").ToString();
            }
            set
            {
                this.SetPhValue("so_ct", (object)value);
            }
        }

        public object GetPhValue(string columnName)
        {
            try
            {
                return this.PhData.Rows[FrmCACTPC1.iRow][columnName];
            }
            catch (Exception ex)
            {
                Type dataType = this.PhData.Columns[columnName].DataType;
                if (dataType.Equals(typeof(string)))
                    return (object)"";
                return dataType.Equals(typeof(DateTime)) ? (object)DateTime.Now.Date : (object)null;
            }
        }

        public void SetPhValue(string columnName, object value)
        {
            try
            {
                this.PhData.Rows[FrmCACTPC1.iRow][columnName] = value;
            }
            catch (Exception ex)
            {
            }
        }

        public void UpdateChanged()
        {
            this.PhData.AcceptChanges();
        }

        private DataView PhView
        {
            get
            {
                return this.PhData.DefaultView;
            }
        }

        private DataTable PhData
        {
            get
            {
                return StartUpTrans.DsTrans.Tables[0];
            }
        }

        private DataView CtView
        {
            get
            {
                return this.CtData.DefaultView;
            }
        }

        private DataTable CtData
        {
            get
            {
                return StartUpTrans.DsTrans.Tables[1];
            }
        }

        public Decimal Parsedecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void Tinh_tien_cltg(DataRecord rec)
        {
            Decimal dec1 = FNum.ToDec(rec.Cells["tien_tt"].Value);
            Decimal dec2 = FNum.ToDec(rec.Cells["tien"].Value);
            if (this.Ma_gd.Trim() == "8")
                dec2 = FNum.ToDec(rec.Cells["tt"].Value);
            rec.Cells["tien_cltg"].Value = (object)(dec2 - dec1);
        }

        private void Update_nt0()
        {
            if (this.Ma_nt != StartupBase.M_MA_NT0)
                return;
            this.PhView[0]["t_tien"] = this.PhView[0]["t_tien_tt"] = this.PhView[0]["t_tien_nt"];
            this.PhView[0]["t_tt"] = this.PhView[0]["t_tt_nt"];
            this.PhView[0]["t_tien_cltg"] = (object)0;
            this.PhView[0]["t_thue"] = this.PhView[0]["t_thue_nt"];
            foreach (DataRowView dataRowView in this.CtView)
            {
                dataRowView["tien"] = dataRowView["tien_tt"] = dataRowView["tien_nt"];
                dataRowView["tt"] = dataRowView["tt_nt"];
                dataRowView["tien_cltg"] = (object)0;
                dataRowView["thue"] = dataRowView["thue_nt"];
            }
            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
            {
                dataRowView["t_thue"] = dataRowView["t_thue_nt"];
                dataRowView["t_tien"] = dataRowView["t_tien_nt"];
            }
        }

        private void GrdCtgt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCtChi_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.Up:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveUp((DataRecord)this.GrdCtChi.ActiveRecord);
                    break;
                case Key.Down:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveDown((DataRecord)this.GrdCtChi.ActiveRecord);
                    break;
            }
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
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

        private void Grdhd_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.Up:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveUp((DataRecord)this.Grdhd.ActiveRecord);
                    break;
                case Key.Down:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    this.MoveDown((DataRecord)this.Grdhd.ActiveRecord);
                    break;
            }
        }

        private void XoaThue()
        {
            if (!FrmCACTPC1.IsInEditMode.Value || this.txtMa_gd.Text.Equals("8"))
                return;
            foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec= '" + this.PhData.Rows[FrmCACTPC1.iRow]["stt_rec"].ToString() + "'"))
                StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
        }

        private void ChkTheo_doi_pt_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!FrmCACTPC1.IsInEditMode.Value || this.txthan_tt == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"] = (object)0;
        }

    }
}
