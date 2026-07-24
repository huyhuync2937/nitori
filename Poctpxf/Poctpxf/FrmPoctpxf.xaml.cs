using Apttpb;
using ArapLib;
using Infragistics.Windows.Controls;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using InvtLib;
using SasControls;
using SasControls.ControlLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Linq;
using System.Collections;
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

namespace Poctpxf
{
    public partial class FrmPoctpxf : FormTrans
    {
        public static int iRow = 0;
        private int iOldRow = 0;
        private bool txtDiaChiFocusable = true;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject IsCheckedPx_gia_dd;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject Voucher_Lan0;
        public DataSet DsVitual;
        private DataSet dsCheckData;

        public FrmPoctpxf()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FrmPoctpxf_Loaded);
            this.BindingSasObj = StartupBase.SasObj;
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.txtma_nt;
            this.C_So_ct = this.txtSo_ct;
        }

        private void FrmPoctpxf_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmPoctpxf.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                this.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource("IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource("Voucher_Ma_nt0");
                this.IsCheckedPx_gia_dd = (CodeValueBindingObject)FormMain.FindResource("IsCheckedPx_gia_dd");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource("IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource("Ty_Gia_ValueChange");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource("Voucher_Lan0");
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)this.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                this.GrdCt.Lan = this.M_LAN;
                if (StartUp.M_SD_HDDT.Equals("1"))
                    this.tabHDDT.Visibility = Visibility.Visible;
                else
                    this.tabHDDT.Visibility = Visibility.Hidden;
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtSoDuKH.Visibility = this.tblSoDuKH.Visibility = Visibility.Hidden;
                FreeCodeFieldLib.InitFreeCodeField(this.BindingSasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
                    this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.IsCheckedPx_gia_dd.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["px_gia_dd"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = true;
                    this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                    StartUpTrans.DsTrans.Tables[1].DefaultView.ListChanged += new ListChangedEventHandler(this.DefaultView_ListChanged);
                }
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.LoadDataDu13();
                this.UpdateTonKho();
                this.TabInfo.SelectedIndex = 0;
                this.SetFocusToolbar();
                if (StartupBase.SasObj.GetOption("M_TON_KHO13").ToString() != "1" && this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ton13")))
                    this.GrdCt.FieldLayouts[0].Fields["ton13"].Visibility = Visibility.Collapsed;
                string str = StartupBase.SasObj.GetOption("M_DC_THUE_CK").ToString();
                if (str == "2" || str == "3")
                    return;
                this.lbNhom_hh.Visibility = Visibility.Collapsed;
                this.txtNhom_hh.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void DefaultView_ListChanged(object sender, ListChangedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void LoadData()
        {
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
        }

        private void V_Next()
        {
            if (FrmPoctpxf.iRow < StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                ++FrmPoctpxf.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
        }

        private void V_Previous()
        {
            if (FrmPoctpxf.iRow > 1)
                --FrmPoctpxf.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
        }

        private void V_Top()
        {
            FrmPoctpxf.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count >= 2 ? 1 : 0;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
        }

        private void V_Bottom()
        {
            FrmPoctpxf.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
        }

        private void V_Moi()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.Add;
                string stt_rec = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
                if (string.IsNullOrEmpty(stt_rec))
                    return;
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.txtMa_kh.IsFocus = true;
                this.txtsd_hddt_yn.IsReadOnly = false;
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1["stt_rec"] = (object)stt_rec;
                row1["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row1["ma_gd"] = StartUpTrans.DsTrans.Tables[0].Rows.Count > 1 ? (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString() : (object)StartUpTrans.DmctInfo["ma_gd"].ToString();
                row1["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row1["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row1["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["ma_qs"].ToString().Trim());
                }
                if (row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    row1["ty_giaf"] = 1;
                }
                else
                {
                    row1["ty_giaf"] = StartUp.GetRates(row1["ma_nt"].ToString().Trim(), Convert.ToDateTime(row1["ngay_ct"]).Date);
                }
                row1["sd_hddt_yn"] = (object)0;
                row1["tinh_trang_hddt"] = (object)0;
                row1["mau_hddt"] = (object)DBNull.Value;
                row1["so_seri_hddt"] = (object)DBNull.Value;
                row1["so_ct_hddt"] = (object)DBNull.Value;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["sl_in"] = (object)0;
                row1["sua_tien"] = (object)0;
                row1["px_gia_dd"] = (object)1;
                row1["t_tien"] = (object)0;
                row1["t_tien_nt"] = (object)0;
                row1["t_thue_nt"] = (object)0;
                row1["t_thue"] = (object)0;
                row1["t_tt_nt"] = (object)0;
                row1["t_tt"] = (object)0;
                row1["ma_ms"] = (object)StartUp.M_ma_ms;
                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                row2["stt_rec"] = (object)stt_rec;
                row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)1);
                row2["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row2["ngay_ct"] = ((this.txtNgay_ct.Value == null) ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                row2["gia_nt"] = (object)0;
                row2["tien_nt"] = (object)0;
                row2["gia"] = (object)0;
                row2["tien"] = (object)0;
                row2["ton13"] = (object)0;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                this.iOldRow = FrmPoctpxf.iRow;
                FrmPoctpxf.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                this.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
                this.txtMa_kh.IsFocus = true;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num = (int)ExMessageBox.Show(1015, StartupBase.SasObj, "Không có dữ liệu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                this.txtMa_kh.IsFocus = true;
                FormTrans.currActionTask = ActionTask.Edit;
                if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("0"))
                    this.txtsd_hddt_yn.IsReadOnly = true;
                else
                    this.txtsd_hddt_yn.IsReadOnly = false;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.IsVisibilityFieldsXamDataGridBySua_Tien();
                this.TabInfo.SelectedIndex = 0;
                this.IsInEditMode.Value = true;
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            }
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            FrmCopy frmCopy = new FrmCopy();
            frmCopy.Closed += new EventHandler(this._formcopy_Closed);
            if (this.M_LAN != "V")
                frmCopy.Title = "Copy";
            frmCopy.ShowDialog();
            this.txtsd_hddt_yn.IsReadOnly = false;
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!(sender as FrmCopy).isCopy)
                return;
            string stt_rec = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(stt_rec))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow].ItemArray;
                if (StartUp.M_SD_HDDT.Equals("0"))
                    row1["sd_hddt_yn"] = (object)0;
                row1["tinh_trang_hddt"] = (object)0;
                row1["mau_hddt"] = (object)DBNull.Value;
                row1["so_seri_hddt"] = (object)DBNull.Value;
                row1["so_ct_hddt"] = (object)DBNull.Value;
                row1["stt_rec"] = (object)stt_rec;
                row1["ngay_ct"] = (object)FrmCopy.ngay_ct;
                row1["ngay_lct"] = (object)FrmCopy.ngay_ct;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["sl_in"] = (object)0;
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
                        row2["stt_rec"] = (object)stt_rec;
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                    }
                }
                this.iOldRow = FrmPoctpxf.iRow;
                FrmPoctpxf.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                this.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
            }
        }

        private void V_In()
        {
            FrmIn frmIn = new FrmIn(this.IsNd51);
            if (this.M_LAN != "V")
                frmIn.Title = "Print";
            frmIn.ShowDialog();
        }

        public string LastSo_ct(string ma_qs)
        {
            string str = "SELECT MAX(so_ct) as last_so_ct  FROM ph86 WHERE ma_qs='" + ma_qs + "'";
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = str;
            return StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows[0][0].ToString().Trim();
        }

        private void V_Nhan()
        {
            try
            {
                bool flag1 = true;
                if (!this.IsSequenceSave)
                {
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(1020, StartupBase.SasObj, "Chưa có mã khách hàng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag1 = false;
                        this.txtMa_kh.IsFocus = true;
                    }
                    else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(1025, StartupBase.SasObj, "Chưa vào tài khoản có!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag1 = false;
                        this.txtMa_nx.IsFocus = true;
                    }
                    else if (this.txtNgay_ct.dValue == new DateTime())
                    {
                        int num = (int)ExMessageBox.Show(1030, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtNgay_ct.Focus();
                        flag1 = false;
                    }
                    else
                    {
                        int num1;
                        if (flag1 && StartUp.M_NGAY_BAT_DAU.HasValue)
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
                                    goto label_16;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_16:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = false;
                            this.txtNgay_ct.Focus();
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(1542, StartupBase.SasObj, "Chưa vào quyển c.từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = false;
                            this.txtMa_qs.IsFocus = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(275, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            this.txtSo_ct.Focus();
                            flag1 = false;
                        }
                        else if (StartUp.M_SD_HDDT.Equals("1") && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sd_hddt_yn"].ToString() == "1" && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() != string.Empty)
                        {
                            SqlCommand sqlcmd = new SqlCommand();
                            sqlcmd.CommandText = "SELECT * FROM khhddt WHERE ma_kh='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() + "' AND status in (1,2)";
                            if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count == 0)
                            {
                                int num2 = (int)ExMessageBox.Show(999, StartupBase.SasObj, "Khách hàng [" + (StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kh.RowResult["ten_kh"].ToString().Trim() : this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim()) + "] chưa cập nhật khách hàng hóa đơn điện tử!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = false;
                                this.TabInfo.SelectedIndex = 0;
                                this.txtMa_kh.IsFocus = true;
                            }
                        }
                        else if (this.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtSo_ct.Text, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            string str = this.LastSo_ct(this.txtMa_qs.Text.Trim());
                            if (StartUpTrans.M_trung_so.Equals("1"))
                            {
                                if (ExMessageBox.Show(10410, StartupBase.SasObj, "có chứng từ trùng số. số cuối cùng là: [" + str + "]. có lưu chứng từ này không?", "xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                {
                                    this.txtSo_ct.SelectAll();
                                    this.txtSo_ct.Focus();
                                    flag1 = false;
                                }
                            }
                            else if (StartUpTrans.M_trung_so.Equals("1"))
                            {
                                int num2 = (int)ExMessageBox.Show(1045, StartupBase.SasObj, "số chứng từ đã tồn tại!", "xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtSo_ct.SelectAll();
                                this.txtSo_ct.Focus();
                                flag1 = false;
                            }
                        }
                        else if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1050, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = false;
                            this.NewRowCt();
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[0] as DataRecord).Cells["ma_vt"];
                            this.GrdCt.Focus();
                        }
                        else if (!this.CheckVoucherOutofDate())
                            flag1 = false;
                    }
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    {
                        int index = 0;
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            if (string.IsNullOrEmpty(dataRowView.Row["ma_vt"].ToString().Trim()))
                            {
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt"];
                                this.GrdCt.Focus();
                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.F2);
                                return;
                            }
                            if (string.IsNullOrEmpty(dataRowView.Row["tk_vt"].ToString().Trim()))
                            {
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_vt"];
                                this.GrdCt.Focus();
                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.F2);
                                return;
                            }
                            if (string.IsNullOrEmpty(dataRowView.Row["ma_kho_i"].ToString().Trim()))
                            {
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_kho_i"];
                                this.GrdCt.Focus();
                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.F2);
                                return;
                            }
                            if (int.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_ton"].ToString()) == 3 && Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString()) == new Decimal(0))
                            {
                                int num = (int)ExMessageBox.Show(1130, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = false;
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["so_luong"];
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GrdCt.Focus()));
                            }
                            ++index;
                        }
                    }
                }
                if (flag1)
                {
                    if (!this.IsSequenceSave)
                    {
                        this.UpdateTotal();
                        Decimal nValue = this.txtTy_gia.nValue;
                        Decimal result1 = new Decimal(0);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b =>
                       {
                           int num;
                           if (b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())
                           {
                               Decimal? nullable = b.Field<Decimal?>("tien_nt");
                               if ((!(nullable.GetValueOrDefault() == new Decimal(0)) ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                               {
                                   nullable = b.Field<Decimal?>("tien");
                                   num = nullable.GetValueOrDefault() != new Decimal(0) ? 1 : (!nullable.HasValue ? 1 : 0);
                                   goto label_4;
                               }
                           }
                           num = 0;
                       label_4:
                           return num != 0;
                       })).Count<DataRow>().ToString(), out result1);
                        if (nValue != new Decimal(0) && (!this.Chksua_tien.IsChecked.Value && StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0) && result1 < (Decimal)StartUpTrans.DsTrans.Tables[1].DefaultView.Count)
                        {
                            Decimal num1 = SysFunc.Round(nValue * this.txtt_tien_nt.nValue, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)num1;
                            Decimal result2 = new Decimal(0);
                            Decimal? nullable = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien")));
                            if (nullable.HasValue)
                                Decimal.TryParse(nullable.ToString(), out result2);
                            Decimal result3 = new Decimal(0);
                            Decimal result4 = new Decimal(0);
                            Decimal result5 = new Decimal(0);
                            if (!this.ChkSuaThue.IsChecked.Value)
                            {
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"].ToString(), out result3);
                                result4 = SysFunc.Round(num1 * result3 / new Decimal(100), StartUpTrans.M_ROUND);
                                result5 = SysFunc.Round(this.txtt_tien_nt.nValue * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)result5;
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)result4;
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)SysFunc.Round(this.txtt_tien_nt.nValue + result5, StartUpTrans.M_ROUND_NT);
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)SysFunc.Round(num1 + result4, StartUpTrans.M_ROUND);
                            }
                            bool flag2 = false;
                            Decimal result6 = new Decimal(0);
                            Decimal result7 = new Decimal(0);
                            Decimal num2 = new Decimal(0);
                            Decimal num3 = new Decimal(0);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString(), out result7);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"].ToString(), out result6);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result5);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result4);
                            for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                            {
                                DataRecord record = this.GrdCt.Records[index] as DataRecord;
                                Decimal result8 = new Decimal(0);
                                Decimal result9 = new Decimal(0);
                                Decimal num4 = new Decimal(0);
                                Decimal num5 = new Decimal(0);
                                Decimal.TryParse(record.Cells["tien_nt"].Value.ToString(), out result8);
                                Decimal.TryParse(record.Cells["tien"].Value.ToString(), out result9);
                                if (!this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    if (result7 != new Decimal(0))
                                        num5 = SysFunc.Round(result8 / result7 * result5, StartUpTrans.M_ROUND_NT);
                                }
                                else if (result7 != new Decimal(0))
                                    num5 = SysFunc.Round(result8 / result7 * result5, StartUpTrans.M_ROUND);
                                if (result6 != new Decimal(0))
                                    num4 = SysFunc.Round(result9 / result6 * result4, StartUpTrans.M_ROUND);
                                num2 += num4;
                                num3 += num5;
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue_nt"] = (object)num5;
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue"] = (object)num4;
                                if (!flag2 && (!(result8 == new Decimal(0)) || !(result9 != new Decimal(0))))
                                {
                                    record.Cells["tien"].Value = (object)(result9 + (num1 - result2));
                                    flag2 = true;
                                }
                            }
                            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"].ToString()) + result5 - num3);
                            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"].ToString()) + result4 - num2);
                        }
                        else
                        {
                            Decimal result2 = new Decimal(0);
                            Decimal result3 = new Decimal(0);
                            Decimal result4 = new Decimal(0);
                            Decimal result5 = new Decimal(0);
                            Decimal num1 = new Decimal(0);
                            Decimal num2 = new Decimal(0);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString(), out result5);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"].ToString(), out result4);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result3);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result2);
                            for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                            {
                                DataRecord record = this.GrdCt.Records[index] as DataRecord;
                                Decimal result6 = new Decimal(0);
                                Decimal result7 = new Decimal(0);
                                Decimal num3 = new Decimal(0);
                                Decimal num4 = new Decimal(0);
                                Decimal.TryParse(record.Cells["tien_nt"].Value.ToString(), out result6);
                                Decimal.TryParse(record.Cells["tien"].Value.ToString(), out result7);
                                if (!this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    if (result5 != new Decimal(0))
                                        num4 = SysFunc.Round(result6 / result5 * result3, StartUpTrans.M_ROUND_NT);
                                }
                                else if (result5 != new Decimal(0))
                                    num4 = SysFunc.Round(result6 / result5 * result3, StartUpTrans.M_ROUND);
                                if (result4 != new Decimal(0))
                                    num3 = SysFunc.Round(result7 / result4 * result2, StartUpTrans.M_ROUND);
                                num1 += num3;
                                num2 += num4;
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue_nt"] = (object)num4;
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue"] = (object)num3;
                            }
                            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"].ToString()) + result3 - num2);
                            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"].ToString()) + result2 - num1);
                        }
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[1].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                        {
                            dataRowView.Row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            dataRowView.Row["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            dataRowView.Row["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                        }

                        if (dataRowView["dvt1"].ToString().Trim() == dataRowView["dvt"].ToString().Trim())
                        {
                            dataRowView["he_so1"] = 0;
                            dataRowView["so_luong1"] = 0;
                        }
                        if ((Decimal)dataRowView["he_so1"] != new Decimal(0))
                        {
                            dataRowView["so_luong1"] = (Decimal)(Convert.ToDecimal(dataRowView["so_luong"]) * Convert.ToDecimal(dataRowView["he_so1"]));
                        }

                        LocalTable1.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[0].Clone();
                    LocalTable2.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable2.Rows[0]["status"] = (object)0;

                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable2, "stt_rec;row_id");
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable1, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(1090, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return;
                    }
                    string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    string str2 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString();
                    string str3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString();
                    SqlCommand sqlcmd = new SqlCommand("EXEC AddDmqssoct @Stt_rec, @Ma_qs, @So_ct");
                    sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.Char, 11).Value = (object)str1;
                    sqlcmd.Parameters.Add("@Ma_qs", SqlDbType.Char, 50).Value = (object)str2;
                    sqlcmd.Parameters.Add("@So_ct", SqlDbType.Char, 16).Value = (object)str3;
                    StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                }
                if (!this.IsSequenceSave)
                {
                    if (flag1)
                    {
                        this.dsCheckData = StartUp.CheckData();
                        this.dsCheckData.Tables[0].AcceptChanges();
                        if (this.dsCheckData.Tables.Count > 0)
                        {
                            string str = "";
                            foreach (DataRowView dataRowView in this.dsCheckData.Tables[0].DefaultView)
                            {
                                if (flag1)
                                {
                                    switch (dataRowView[0].ToString())
                                    {
                                        case "PH01":
                                            if (StartUpTrans.M_trung_so.Equals("1"))
                                            {
                                                if (ExMessageBox.Show(1095, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                {
                                                    this.txtSo_ct.SelectAll();
                                                    this.txtSo_ct.Focus();
                                                    flag1 = false;
                                                    break;
                                                }
                                                break;
                                            }
                                            if (StartUpTrans.M_trung_so.Equals("2"))
                                            {
                                                int num = (int)ExMessageBox.Show(1100, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.txtSo_ct.SelectAll();
                                                this.txtSo_ct.Focus();
                                                flag1 = false;
                                                break;
                                            }
                                            break;
                                        case "PH02":
                                            int num1 = (int)ExMessageBox.Show(1105, StartupBase.SasObj, "Tk nợ là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = false;
                                            this.txtMa_nx.IsFocus = true;
                                            break;
                                        case "PH03":
                                            int num2 = (int)ExMessageBox.Show(1110, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = false;
                                            this.txtTk_thue_co.IsFocus = true;
                                            break;
                                        case "CT01":
                                            int int16 = (int)Convert.ToInt16(dataRowView[1]);
                                            int num3 = (int)ExMessageBox.Show(1115, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = false;
                                            this.GrdCt.ActiveCell = (this.GrdCt.Records[int16] as DataRecord).Cells["tk_vt"];
                                            this.GrdCt.Focus();
                                            break;
                                        case "CT02":
                                            StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ton13"] = InFuncLib.GetTon13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_kho_i"].ToString(), StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString(), StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vv_i"].ToString());
                                            if (!str.Contains(StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString().Trim()))
                                            {
                                                str = str + StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString().Trim() + ", ";
                                                break;
                                            }
                                            break;
                                    }
                                    this.dsCheckData.Tables[0].Rows.Remove(dataRowView.Row);
                                }
                                else
                                    break;
                            }
                            if (!string.IsNullOrEmpty(str))
                            {
                                if (StartUp.M_CHK_TON_VT.Equals("2"))
                                {
                                    int num = (int)ExMessageBox.Show(1120, StartupBase.SasObj, "Có vật tư [" + str + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    flag1 = false;
                                }
                                else if (StartUp.M_CHK_TON_VT.Equals("1"))
                                {
                                    int num4 = (int)ExMessageBox.Show(1125, StartupBase.SasObj, "Có vật tư [" + str + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                            }
                        }
                    }
                    if (flag1 && this.IsNd51)
                    {
                        if (this.txtMa_qs.RowResult == null)
                            this.txtMa_qs.SearchInit();
                        this.UpdateNewSoCt(FormTrans.SasO, this.txtMa_qs.Text);
                        if (FormTrans.currActionTask == ActionTask.Edit && this.txtMa_qs.RowResult != null)
                            this.UpdateNewNgayCt(FormTrans.SasO, this.txtMa_qs.Text, this.GetCurrentSo_ct(this.txtMa_qs.RowResult["transform"].ToString(), this.txtSo_ct.Text.Trim()));
                    }
                }
                if (flag1)
                {
                    string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                    new Thread((ThreadStart)(() =>
                   {
                       this.Post();
                       if (this.IsSequenceSave)
                           return;
                       this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate) new Action(() =>
             {
                         if (!StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(_stt_rec1))
                             return;
                         this.UpdateTonKho();
                         this.LoadDataDu13();
                     }));
                   })).Start();
                    if (!this.IsSequenceSave)
                    {
                        int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                        if (FrmPoctpxf.iRow != pos)
                        {
                            DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                            DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                            row2.ItemArray = row1.ItemArray;
                            if (FrmPoctpxf.iRow > pos)
                                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                            else
                                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                            StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                            FrmPoctpxf.iRow = pos;
                        }
                        this.IsInEditMode.Value = false;
                        FormTrans.currActionTask = ActionTask.View;
                    }
                    this.txtsd_hddt_yn.IsReadOnly = true;
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[Poctpxf-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private void UpdateTonKho()
        {
            string str1 = "";
            string str2 = "";
            string str3 = "";
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseInt(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["vt_ton_kho"], 0) == 1)
                {
                    str1 = str1 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString().Trim();
                    str2 = str2 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString().Trim();
                    str3 = str3 + ";" + StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vv_i"].ToString().Trim();
                }
                else
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = (object)DBNull.Value;
            }
            if (!string.IsNullOrEmpty(str1) && !string.IsNullOrEmpty(str2))
            {
                string ma_kho = str1.Substring(1);
                string ma_vt = str2.Substring(1);
                string ma_vv = str3.Substring(1);
                DataTable listTon13 = InFuncLib.GetListTon13(StartupBase.SasObj, ma_kho, ma_vt, ma_vv);
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
                        SqlCommand sqlcmd = new SqlCommand("exec [CheckTonvv] @Ma_kho");
                        sqlcmd.Parameters.Add("@Ma_kho", SqlDbType.VarChar).Value = (object)str4;
                        if (this.ParseInt(StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows[0][0], 0) == 0)
                            str6 = "";
                        DataRow[] dataRowArray = listTon13.Select("ma_kho LIKE '" + str4 + "' AND ma_vt LIKE '" + str5 + "' AND ma_vv LIKE '" + str6 + "'");
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = dataRowArray.Length <= 0 ? (object)0 : dataRowArray[0]["ton13"];
                    }
                }
            }
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmTim frmTim = new FrmTim(StartupBase.SasObj, StartUpTrans.filterId, StartUpTrans.filterView);
                SysFunc.LoadIcon((Window)frmTim);
                if (this.M_LAN != "V")
                    frmTim.Title = "Search";
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
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString());
            if (this.M_LAN != "V")
                formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString());
            formView.ListFieldSum = "t_tt_nt;t_tt";
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "PoctpxfXemView";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                FrmPoctpxf.iRow = index + 1;
                string str = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_Huy()
        {
            this.IsInEditMode.Value = false;
            this.txtsd_hddt_yn.IsReadOnly = true;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmPoctpxf.iRow = this.iOldRow;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
                        break;
                    }
                    break;
                case ActionTask.Edit:
                    string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                    }
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow].ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    break;
            }
            FormTrans.currActionTask = ActionTask.None;
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUpTrans.UpdateTkSd13(1, 0);
                string format1 = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd1 = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 12 ? string.Format(format1, (object)"[DeleteVoucher]") : string.Format(format1, (object)StartUpTrans.Process_Store[12]));
                sqlcmd1.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd1.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)str;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd1);
                if (this.IsNd51)
                {
                    string format2 = "exec [dbo].{0} @stt_rec;";
                    SqlCommand sqlcmd2 = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 3 ? string.Format(format2, (object)"[POCTPXF-PostCTHHD]") : string.Format(format2, (object)StartUpTrans.Post_store[3]));
                    if (StartUp.currActionTask == ActionTask.Delete)
                    {
                        sqlcmd2.CommandText += " DECLARE @ngay_ct smalldatetime;";
                        sqlcmd2.CommandText += " DECLARE @so_ct numeric(16, 0), @so_ct1 numeric(16, 0);";
                        SqlCommand sqlCommand1 = sqlcmd2;
                        sqlCommand1.CommandText = sqlCommand1.CommandText + " SELECT @so_ct = so_ct - 1, @so_ct1 = so_ct1 - 1 FROM dmqs WHERE ma_qs='" + this.txtMa_qs.Text.Trim() + "';";
                        sqlcmd2.CommandText += " IF @so_ct = @so_ct1";
                        SqlCommand sqlCommand2 = sqlcmd2;
                        sqlCommand2.CommandText = sqlCommand2.CommandText + "    SELECT @ngay_ct = ngay_qs1 FROM dmqs WHERE ma_qs='" + this.txtMa_qs.Text.Trim() + "';";
                        sqlcmd2.CommandText += " ELSE";
                        SqlCommand sqlCommand3 = sqlcmd2;
                        sqlCommand3.CommandText = sqlCommand3.CommandText + "    SELECT TOP 1 @ngay_ct = ngay_ct FROM cthhd WHERE ma_qs='" + this.txtMa_qs.Text.Trim() + "' AND so_ct IS NOT NULL ORDER BY ngay_ct DESC;";
                        SqlCommand sqlCommand4 = sqlcmd2;
                        sqlCommand4.CommandText = sqlCommand4.CommandText + " UPDATE dmqs SET so_ct = @so_ct, ngay_ct = ISNULL(@ngay_ct,ngay_qs1) WHERE ma_qs='" + this.txtMa_qs.Text.Trim() + "';";
                    }
                    sqlcmd2.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)str;
                    StartupBase.SasObj.ExcuteNonQuery(sqlcmd2);
                }
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmPoctpxf.iRow);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    FrmPoctpxf.iRow = FrmPoctpxf.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmPoctpxf.iRow - 1 : FrmPoctpxf.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["stt_rec"].ToString());
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            FormTrans.currActionTask = ActionTask.None;
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            this.IsVisibilityFieldsXamDataGridByMa_NT(ma_nt);
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
        }

        private void IsVisibilityFieldsXamDataGridByMa_NT(string ma_nt)
        {
            this.ChangeLanguage();
            if (ma_nt == StartUpTrans.M_ma_nt0)
            {
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Visible;
                FieldSettings settings1 = this.GrdCt.FieldLayouts[0].Fields["gia"].Settings;
                FieldLength fieldLength = this.GrdCt.FieldLayouts[0].Fields["gia"].Width.Value;
                double num1 = fieldLength.Value;
                settings1.CellMaxWidth = num1;
                FieldSettings settings2 = this.GrdCt.FieldLayouts[0].Fields["tien"].Settings;
                fieldLength = this.GrdCt.FieldLayouts[0].Fields["tien"].Width.Value;
                double num2 = fieldLength.Value;
                settings2.CellMaxWidth = num2;
            }
            this.ChangeLanguage();
        }

        private void IsVisibilityFieldsXamDataGridBySua_Tien()
        {
            this.IsCheckedSua_tien.Value = this.Chksua_tien.IsChecked.Value;
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.IsInEditMode.Value && (this.txtMa_kh.RowResult != null && !string.IsNullOrEmpty(this.txtMa_kh.Text.Trim())))
            {
                if (this.txtMa_kh.IsDataChanged)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = this.txtMa_kh.RowResult["ten_kh"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = this.txtMa_kh.RowResult["ten_kh2"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = this.txtMa_kh.RowResult["ma_so_thue"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"] = this.txtMa_kh.RowResult["han_tt"];
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
                    if (this.txtMa_kh.RowResult["tk"].ToString().Trim() != "")
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = this.txtMa_kh.RowResult["tk"];
                    if (this.txtMa_kh.RowResult["dia_chi"] == DBNull.Value || string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)"";
                        this.txtDiaChiFocusable = true;
                        this.txtDia_chi.IsTabStop = true;
                    }
                    else
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = this.txtMa_kh.RowResult["dia_chi"];
                        this.txtDiaChiFocusable = false;
                        this.txtDia_chi.IsTabStop = false;
                    }
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtDia_chi.Focus()));
                    this.LoadDataDu13();
                }
                else if (this.txtMa_kh.RowResult != null)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = this.txtMa_kh.RowResult["ten_kh"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = this.txtMa_kh.RowResult["ten_kh2"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = this.txtMa_kh.RowResult["ma_so_thue"];
                    if (this.txtMa_kh.RowResult["dia_chi"] == DBNull.Value || string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
                    {
                        this.txtDiaChiFocusable = true;
                        this.txtDia_chi.IsTabStop = true;
                    }
                    else
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = this.txtMa_kh.RowResult["dia_chi"];
                        this.txtDiaChiFocusable = false;
                        this.txtDia_chi.IsTabStop = false;
                    }
                }
                if (StartUp.M_SD_HDDT.Equals("1") && this.txtMa_kh.IsDataChanged)
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sd_hddt_yn"] = (object)SysFunc.suDungHDDT(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString());
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

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()))
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
            }
            this.LoadDataDu13();
        }

        private void txtma_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.txtma_nt.IsDataChanged)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.txtma_nt.RowResult["loai_tg"];
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.IsVisibilityFieldsXamDataGridByMa_NT(this.txtma_nt.Text.Trim());
            if (this.txtma_nt.Text == StartUpTrans.M_ma_nt0.Trim())
                this.txtTy_gia.Value = (object)1;
            else
                this.txtTy_gia.Value = (object)StartUp.GetRates(this.txtma_nt.RowResult["ma_nt"].ToString().Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
            this.CalculateTyGia();
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            if (this.txtTy_gia.Value == DBNull.Value)
                this.txtTy_gia.Value = (object)0;
            if (!(this.txtTy_gia.OldValue != this.txtTy_gia.nValue))
                return;
            this.CalculateTyGia();
            this.Ty_Gia_ValueChange.Value = !this.Ty_Gia_ValueChange.Value;
        }

        private void CalculateTyGia()
        {
            Decimal num1 = new Decimal(0);
            Decimal nValue = this.txtTy_gia.nValue;
            if (!(nValue != new Decimal(0)))
                return;
            for (int index = 0; index < this.GrdCt.Records.Count; ++index)
            {
                Decimal result1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal.TryParse((this.GrdCt.Records[index] as DataRecord).Cells["so_luong"].Value.ToString(), out result2);
                Decimal.TryParse((this.GrdCt.Records[index] as DataRecord).Cells["gia_nt"].Value.ToString(), out result1);
                if (result2 * result1 != new Decimal(0))
                {
                    num2 = SysFunc.Round(result2 * result1, StartUpTrans.M_ROUND_NT);
                    (this.GrdCt.Records[index] as DataRecord).Cells["tien_nt"].Value = (object)num2;
                }
                if (num2 * nValue != new Decimal(0))
                    (this.GrdCt.Records[index] as DataRecord).Cells["tien"].Value = (object)SysFunc.Round(num2 * nValue, StartUpTrans.M_ROUND);
                if (result1 * nValue != new Decimal(0))
                    (this.GrdCt.Records[index] as DataRecord).Cells["gia"].Value = (object)SysFunc.Round(result1 * nValue, StartUpTrans.M_ROUND_GIA);
            }
            this.UpdateTotal();
        }

        private void txtT_thue_nt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = !this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)SysFunc.Round(SysFunc.Parsedecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"], new Decimal(0)) * this.txtTy_gia.nValue, StartUpTrans.M_ROUND) : StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"];
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result2);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["t_tt_nt"] = (object)(result1 + result2);
        }

        private void txtT_thue_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result2);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["t_tt"] = (object)(result1 + result2);
        }

        private void txtThue_suat_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            this.UpdateTotal();
        }

        private void Chksua_tien_Click(object sender, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            bool? isChecked = this.Chksua_tien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.CalculateTyGia();
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
                if (!this.IsInEditMode.Value)
                    return;
                Decimal result1 = new Decimal(0);
                int result2 = 0;
                int.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString(), out result2);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result1);
                if (this.GrdCt.ActiveCell != null && StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "ma_vt":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                            CellCollection cells = e.Cell.Record.Cells;
                            if (e.Cell.IsDataChanged)
                            {
                                if (autoCompleteControl1.RowResult != null)
                                {
                                    e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl1.RowResult["ten_vt"];
                                    e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl1.RowResult["ten_vt2"];
                                    e.Cell.Record.Cells["dvt"].Value = autoCompleteControl1.RowResult["dvt"];
                                    e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl1.RowResult["tk_vt"];
                                    (e.Cell.Record.DataItem as DataRowView)["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                                    e.Cell.Record.Cells["gia_ton"].Value = autoCompleteControl1.RowResult["gia_ton"];
                                    AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                                    if (autoCompleteControl2 != null)
                                    {
                                        autoCompleteControl2.SearchInit();
                                        if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                                            e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl2.RowResult["tk_dl"];
                                    }
                                    DataRowView dataItem = e.Cell.Record.DataItem as DataRowView;
                                    dataItem["sua_tk_vt"] = autoCompleteControl1.RowResult["sua_tk_vt"];
                                    dataItem["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                                    ControlFunction.RefreshSingleBinding((DependencyObject)CellValuePresenter.FromCell(e.Cell.Record.Cells["tk_vt"]), AutoCompleteTextBox.IsReadOnlyProperty);
                                }
                                if (this.ParseInt(autoCompleteControl1.RowResult["vt_ton_kho"], 0) == 1)
                                {
                                    if (!string.IsNullOrEmpty(e.Cell.Record.Cells["ma_vt"].Value.ToString()) && !string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kho_i"].Value.ToString()))
                                        e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                }
                                else
                                    e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;

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
                        case "ma_kho_i":
                            if (e.Editor.Value == null || !e.Cell.IsDataChanged)
                                break;
                            AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                            if (autoCompleteControl3.RowResult == null)
                                autoCompleteControl3.SearchInit();
                            if (autoCompleteControl3.RowResult != null && (autoCompleteControl3.RowResult["sua_tk_vt"] != DBNull.Value && Convert.ToDecimal(autoCompleteControl3.RowResult["sua_tk_vt"]) == new Decimal(0)))
                                e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl3.RowResult["tk_vt"];
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
                                    if (!string.IsNullOrEmpty(e.Cell.Record.Cells["ma_vt"].Value.ToString()) && !string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kho_i"].Value.ToString()))
                                        e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                }
                                else
                                    e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;
                            }
                            break;
                        case "so_luong":
                            if (e.Editor.Value == null || !e.Cell.IsDataChanged)
                                break;
                            Decimal result3 = new Decimal(0);
                            Decimal result4 = new Decimal(0);
                            Decimal result5 = new Decimal(0);
                            Decimal result6 = new Decimal(0);
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result3);
                            AutoCompleteTextBox autoCompleteControl5 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                            if (autoCompleteControl5.RowResult != null && autoCompleteControl5.RowResult["gia_ton"] != DBNull.Value && int.Parse(autoCompleteControl5.RowResult["gia_ton"].ToString()) == 3 && result3 == new Decimal(0))
                            {
                                int num = (int)ExMessageBox.Show(1130, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() =>
                               {
                                   this.GrdCt.ActiveCell = e.Cell.Record.Cells["so_luong"];
                                   this.GrdCt.Focus();
                               }));
                                break;
                            }
                            if (result3 == new Decimal(0))
                            {
                                e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                e.Cell.Record.Cells["gia"].Value = (object)0;
                            }
                            Decimal.TryParse(e.Cell.Record.Cells["gia_nt"].Value.ToString(), out result4);
                            if (result3 * result4 != new Decimal(0))
                            {
                                if (!this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)SysFunc.Round(result3 * result4, StartUpTrans.M_ROUND_NT);
                                else
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)SysFunc.Round(result3 * result4, StartUpTrans.M_ROUND);
                            }
                            Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result5);
                            if (result3 * result5 != new Decimal(0))
                                e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result3 * result5, StartUpTrans.M_ROUND);
                            Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result6);
                            if (this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result6, StartUpTrans.M_ROUND);
                            this.UpdateTotal();
                            break;
                        case "gia_nt":
                            if (e.Editor.Value == null || !e.Cell.IsDataChanged)
                                break;
                            Decimal result7 = new Decimal(0);
                            Decimal result8 = new Decimal(0);
                            Decimal result9 = new Decimal(0);
                            Decimal num1 = new Decimal(0);
                            Decimal num2 = new Decimal(0);
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result7);
                            Decimal.TryParse(e.Editor.Value.ToString(), out result8);
                            if (result7 * result8 != new Decimal(0))
                            {
                                bool? isChecked = this.Chksua_tien.IsChecked;
                                if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                                {
                                    Decimal num3 = result7 * result8;
                                    Decimal num4 = this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num3, StartUpTrans.M_ROUND) : SysFunc.Round(num3, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)num4;
                                }
                            }
                            if (result1 * result8 != new Decimal(0))
                                e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(result1 * result8, StartUpTrans.M_ROUND_GIA);
                            if (this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(result8, StartUpTrans.M_ROUND_GIA);
                            Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result9);
                            if (result7 * result9 != new Decimal(0))
                                e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result7 * result9, StartUpTrans.M_ROUND);
                            this.UpdateTotal();
                            break;
                        case "tien_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result10 = new Decimal(0);
                                Decimal result11 = new Decimal(0);
                                Decimal result12 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result11);
                                Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result12);
                                Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result10);
                                if (result11 * result1 != new Decimal(0))
                                    e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result11 * result1, StartUpTrans.M_ROUND);
                                if (this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                    e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result11, StartUpTrans.M_ROUND);
                                Decimal result13 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["gia_nt"].Value.ToString(), out result13);
                                if (result13 == new Decimal(0))
                                {
                                    if (result10 != new Decimal(0))
                                        result13 = SysFunc.Round(result11 / result10, StartUpTrans.M_ROUND_GIA);
                                    e.Cell.Record.Cells["gia_nt"].Value = (object)result13;
                                }
                                this.UpdateTotal();
                                break;
                            }
                            break;
                        case "gia":
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result10 = new Decimal(0);
                                Decimal result11 = new Decimal(0);
                                Decimal result12 = new Decimal(0);
                                Decimal result13 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result10);
                                Decimal.TryParse(e.Editor.Value.ToString(), out result11);
                                if (result11 * result10 != new Decimal(0))
                                    e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(result10 * result11, StartUpTrans.M_ROUND);
                                if (result10 != new Decimal(0) && result11 == new Decimal(0))
                                {
                                    Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result12);
                                    Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result13);
                                    if (!this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                        e.Cell.Record.Cells["gia_nt"].Value = (object)SysFunc.Round(result12 / result10, StartUpTrans.M_ROUND_GIA_NT);
                                    else
                                        e.Cell.Record.Cells["gia_nt"].Value = (object)SysFunc.Round(result12 / result10, StartUpTrans.M_ROUND_GIA);
                                    e.Cell.Record.Cells["gia"].Value = (object)SysFunc.Round(result13 / result10, StartUpTrans.M_ROUND_GIA);
                                }
                                this.UpdateTotal();
                                break;
                            }
                            break;
                        case "tien":
                            Decimal result14 = new Decimal(0);
                            Decimal result15 = new Decimal(0);
                            Decimal result16 = new Decimal(0);
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result14);
                            Decimal.TryParse(e.Cell.Record.Cells["tien_nt"].Value.ToString(), out result15);
                            Decimal.TryParse(e.Cell.Record.Cells["tien"].Value.ToString(), out result16);
                            this.UpdateTotal();
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
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void UpdateTotal()
        {
            try
            {
                if (FormTrans.currActionTask == ActionTask.View)
                    return;
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result2);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)result1;
                Decimal result3 = new Decimal(0);
                Decimal result4 = new Decimal(0);
                Decimal? nullable1 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien_nt")));
                if (nullable1.HasValue)
                    Decimal.TryParse(nullable1.ToString(), out result3);
                Decimal num1 = this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(result3, StartUpTrans.M_ROUND) : SysFunc.Round(result3, StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = (object)num1;
                Decimal? nullable2 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien")));
                if (nullable2.HasValue)
                    Decimal.TryParse(nullable2.ToString(), out result4);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);
                if (!string.IsNullOrEmpty(this.txtThue_suat.Text))
                {
                    Decimal result5 = new Decimal(0);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"].ToString(), out result5);
                    num2 = this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num1 * result5 / new Decimal(100), StartUpTrans.M_ROUND) : SysFunc.Round(num1 * result5 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                    num3 = SysFunc.Round(result4 * result5 / new Decimal(100), StartUpTrans.M_ROUND);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num2;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)num3;
                }
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(num2 + num1);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(num3 + result4);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtMa_thue_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_thue.RowResult == null || string.IsNullOrEmpty(this.txtMa_thue.Text))
                return;
            if (this.txtMa_thue.RowResult["thue_suat"] != DBNull.Value)
                this.txtThue_suat.Text = this.txtMa_thue.RowResult["thue_suat"].ToString();
            if (this.txtMa_thue.RowResult["tk_thue_co"] != DBNull.Value)
                this.txtTk_thue_co.Text = this.txtMa_thue.RowResult["tk_thue_co"].ToString();
        }

        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            return this.NewRowCt();
        }

        private bool NewRowCt()
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
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[StartUpTrans.DsTrans.Tables[1].DefaultView.Count - 1]["ma_kho_i"];
                dataRow["gia_nt"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["tien_nt"] = (object)0;
                dataRow["tien"] = (object)0;
                dataRow["ton13"] = (object)0;
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

        private bool NewRowCtNG(DataRow data)
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
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[StartUpTrans.DsTrans.Tables[1].DefaultView.Count - 1]["ma_kho_i"];
                dataRow["gia_nt"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["tien_nt"] = (object)0;
                dataRow["tien"] = (object)0;
                dataRow["ton13"] = (object)0;
                dataRow["ma_vt"] = (object)data["ma_vt"];
                dataRow["ten_vt"] = (object)data["ten_vt"];
                dataRow["dvt"] = (object)data["dvt"];
                dataRow["ma_kho"] = (object)data["ma_kho"];
                dataRow["so_luong"] = (object)data["ton_cuoi"];



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


        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.txtMa_tc.IsFocus = true));
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"];
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.StartEditMode);
                    if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete") && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
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
                        CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell((this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"]);
                        if (cellValuePresenter != null && cellValuePresenter.Editor is ControlHostEditor editor)
                        {
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(editor);
                            if (string.IsNullOrEmpty(autoCompleteControl.Text.Trim()))
                            {
                                int num1 = (int)ExMessageBox.Show(1135, StartupBase.SasObj, "Chưa nhập mã vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else if (autoCompleteControl != null)
                            {
                                if (autoCompleteControl.CheckLostFocus())
                                {
                                    DataTable giaPn = StartUp.GetGiaPN(this.txtNgay_ct.Value == null ? "" : string.Format("{0:yyyyMMdd}", (object)this.txtNgay_ct.dValue), (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"].Value.ToString(), (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_kho_i"].Value.ToString(), this.txtMa_kh.Text);
                                    if (giaPn.Rows.Count > 0)
                                    {
                                        new FrmPoctpxf_PN(giaPn, this.GrdCt.ActiveRecord, this.Voucher_Ma_nt0).ShowDialog();
                                        this.UpdateTotal();
                                    }
                                    else
                                    {
                                        int num2 = (int)ExMessageBox.Show(1140, StartupBase.SasObj, "Không tìm thấy phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    }
                                }
                                else
                                {
                                    int num3 = (int)ExMessageBox.Show(1140, StartupBase.SasObj, "Không tìm thấy phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                            }
                        }
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(1150, StartupBase.SasObj, "Có xoá dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord))
                        break;
                    Cell activeCell = this.GrdCt.ActiveCell;
                    int num4 = activeRecord.Index;
                    if (activeRecord.Index == 0)
                    {
                        if (this.GrdCt.Records.Count == 1)
                            this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord.Index == this.GrdCt.Records.Count - 1)
                        num4 = activeRecord.Index - 1;
                    int num5 = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num5 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[activeRecord.Index].Row);
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[num4 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : num4];
                        this.UpdateTotal();
                    }
                    break;
            }
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void txtTk_thue_co_ValueChanged()
        {
            if (this.txtTk_thue_co.RowResult == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_cn"] = (object)this.txtTk_thue_co.RowResult["tk_cn"].ToString();
            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_cn"].ToString().Trim().Equals("0"))
                this.txtMa_kh2.Text = "";
            ControlFunction.RefreshSingleBinding((DependencyObject)this.txtMa_kh2, Control.IsTabStopProperty);
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            if (menuItemName.Equals("btnSave"))
                return;
            this.LoadDataDu13();
            this.UpdateTonKho();
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void ChkSuaThue_Click(object sender, RoutedEventArgs e)
        {
            if (this.ChkSuaThue.IsChecked.Value)
            {
                this.txtT_thue_nt.Focus();
            }
            else
            {
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal result3 = new Decimal(0);
                Decimal num1 = new Decimal(0);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString(), out result2);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"].ToString(), out result1);
                Decimal num2 = this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(result2 * result1 / new Decimal(100), StartUpTrans.M_ROUND) : SysFunc.Round(result2 * result1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num2;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = !this.txtma_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)SysFunc.Round(SysFunc.Parsedecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"], new Decimal(0)) * this.txtTy_gia.nValue, StartUpTrans.M_ROUND) : StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"];
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result3);
                StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpxf.iRow]["t_tt_nt"] = (object)(result2 + result3);
                this.txtHan_tt.Focus();
            }
        }

        private void LoadDataDu13()
        {
            this.txtSoDuKH.Value = (object)ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString());
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.IsInEditMode.Value)
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri"] = (object)this.txtMa_qs.RowResult["so_seri"].ToString().Trim();
            if (!this.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                return;
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri"].ToString().Trim()) && this.txtMa_qs.RowResult != null)
                this.txtSo_seri.Text = this.txtMa_qs.RowResult["so_seri"].ToString();
            int num;
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["kh_mau_hd"].ToString().Trim()))
            {
                bool? isChecked = this.ChkThueDauRa.IsChecked;
                if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                {
                    num = 1;
                    goto label_9;
                }
            }
            num = this.txtMa_qs.RowResult == null ? 1 : 0;
        label_9:
            if (num == 0)
            {
                this.txtkh_mau_hd.Text = this.txtMa_qs.RowResult["kh_mau_hd"].ToString();
                if (string.IsNullOrEmpty(this.txtkh_mau_hd.Text.Trim()))
                    this.txtkh_mau_hd.Text = this.txtMa_qs.RowResult["mau_hd"].ToString();
            }
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()) || this.IsNd51 && this.txtMa_qs.IsDataChanged)
            {
                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim()) || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"].ToString().Trim()) || this.IsNd51)
                {
                    this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
                }
                else
                    this.txtSo_ct.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim();
            }
        }

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.Return) && (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt)))
            {
                TextBox textBox = sender as TextBox;
                textBox.SelectedText = Environment.NewLine;
                ++textBox.SelectionStart;
                textBox.SelectionLength = 0;
                e.Handled = true;
            }
            else
            {
                if (!Keyboard.IsKeyDown(Key.Return) && !Keyboard.IsKeyDown(Key.Tab))
                    return;
                (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
                e.Handled = true;
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

        private void txtngay_lct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtngay_lct.IsFocusWithin || (FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy || this.txtNgay_ct.dValue.Date.Equals(this.txtngay_lct.dValue.Date)))
                return;
            int num = (int)ExMessageBox.Show(1155, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtHan_tt_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtHan_tt.SelectAll();
        }

        private void txtSo_seri_GotFocus(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() =>
           {
               if (this.txtSo_seri.IsReadOnly)
                   return;
               this.txtSo_seri.Text = this.txtSo_seri.Text.Trim();
               this.txtSo_seri.CaretIndex = this.txtSo_seri.Text.Length;
           }));
        }

        private void V_HuyHD()
        {
            try
            {
                bool flag = false;
                if (this.C_QS != null && this.C_NgayHT != null)
                {
                    switch (StartUpTrans.CheckQS(this.C_QS.Text, this.C_NgayHT.dValue.ToString("yyyyMMdd"), (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString())))
                    {
                        case 1:
                            int num1 = (int)ExMessageBox.Show(235, FormTrans.SasO, "Ngày bắt đầu sử dụng quyển sổ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            break;
                        case 2:
                            int num2 = (int)ExMessageBox.Show(240, FormTrans.SasO, "Quyền sử dụng quyển sổ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            break;
                    }
                }
                if (!flag)
                {
                    FrmLogin frmLogin = new FrmLogin();
                    frmLogin.ShowDialog();
                    if (frmLogin.IsLogined)
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"] = (object)3;
                        DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                        LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                        DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                        DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            dataRowView["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            dataRowView["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            dataRowView["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                        }
                        if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            int num3 = (int)ExMessageBox.Show(245, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            return;
                        }
                        new Thread((ThreadStart)(() => this.Post())).Start();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            this.txtsd_hddt_yn.IsReadOnly = true;
        }

        private void txtkh_mau_hd_PreviewKeyDown(object sender, KeyEventArgs e)
        {
        }

        private void txtHTTT_PreviewKeyDown(object sender, KeyEventArgs e)
        {
        }

        private void txtGhi_chu_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void txtGhi_chu_PreviewKeyDown_1(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value || e.Key != Key.Return && e.Key != Key.Tab)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }
        private void txtMa_ncc_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.F11)
                return;
            this.OnF11();
        }
        private void OnF11()
        {
            //string text = this.txtMa_ncc.Text;
            //char[] chArray = new char[1] { ',' };
            //foreach (object obj in text.Split(chArray))
            //{
            //    foreach (DataRow dataRow in this.dtMa_ncc.Select(string.Format("ma_kh = '{0}'", obj)))
            //        dataRow["tag"] = (object)true;
            //}
            //DataTable dataTable = this.dtMa_ncc.Copy();
            //COTKTH2Dvcs cotktH2Dvcs = new COTKTH2Dvcs();
            ////dataTable.DefaultView.RowFilter = !string.IsNullOrEmpty(this.txtMa_dvcs.Text.Trim()) ? "ma_dvcs LIKE '" + this.txtMa_dvcs.Text + "'" : "1=1";
            //cotktH2Dvcs.GrdCt.DataSource = (IEnumerable)dataTable.DefaultView;
            //if (StartupBase.M_LAN != "V")
            //    cotktH2Dvcs.Title = "Supplier list";
            //bool? nullable = cotktH2Dvcs.ShowDialog();
            //if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
            //    return;
            //DataRow[] dataRowArray = dataTable.Select("tag = 1");
            //string str = "";
            //foreach (DataRow dataRow in dataRowArray)
            //    str = str + (str == "" ? "" : ",") + dataRow["ma_kh"].ToString().Trim();
            //this.txtMa_ncc.Text = str;
        }
        private void btnNG_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsEditMode)
            {
                FrmLoc Locdonhang = new FrmLoc();
                Locdonhang.ShowDialog();

                if(StartUp.KhoNG != null && StartUp.KhoNG.Length > 0)
                {
                    foreach (DataRow dataRow in StartUp.KhoNG)
                    {
                        this.NewRowCtNG(dataRow);
                        //DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                        //row2["stt_rec"] = (object)StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"];
                        //row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)(StartUpTrans.DsTrans.Tables[1].DefaultView.Count + 1));
                        //row2["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].Rows[0]["ngay_ct"];
                        //row2["ma_vt"] = dataRow["ma_vt"];
                        //row2["ten_vt"] = dataRow["ten_vt"];
                        //row2["ma_kho_i"] = dataRow["ma_kho"];
                        //StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                    }  
                }
            }
               
        }

        private void btnSoHD_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsEditMode)
            {
                int num1 = (int)ExMessageBox.Show(775, StartupBase.SasObj, "Phải lưu chứng từ rồi mới phân bổ cho các hóa đơn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString() != "2")
            {
                int num2 = (int)ExMessageBox.Show(780, StartupBase.SasObj, "Phải ghi vào sổ cái rồi mới phân bổ cho các hđ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                string stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                object obj = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                PbInfo pbInfo = new PbInfo(obj, obj, (object)"", "", "");
                pbInfo.TitleView = "Phan bo";
                Apttpb.StartUp.Procedure = StartUpTrans.CommandInfo["parameter"].ToString().Split(';');
                new Apttpb.StartUp().Pb_tt(stt_rec, (IPhanbo)pbInfo);
                this.Dispatcher.BeginInvoke((Delegate) new Action(() =>
               {
                   SqlCommand sqlcmd = new SqlCommand("SELECT so_ct_tt FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec = @Stt_rec");
                   sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = (object)stt_rec;
                   DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                   if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                       this.lblSo_ct_tt.Text = dataSet.Tables[0].Rows[0][0].ToString();
                   else
                       this.lblSo_ct_tt.Text = "";
               }), DispatcherPriority.Background);
            }
        }

        private void txtsd_hddt_yn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
        }

        private void txtsd_hddt_yn_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtsd_hddt_yn.Text))
                return;
            this.txtsd_hddt_yn.Text = "0";
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value)
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

        private void ChkPx_gia_dd_Click(object sender, RoutedEventArgs e)
        {
            IsCheckedPx_gia_dd.Value = ChkPx_gia_dd.IsChecked.Value;
        }
    }
}
