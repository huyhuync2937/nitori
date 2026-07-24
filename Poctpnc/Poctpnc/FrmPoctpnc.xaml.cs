using ArapLib;
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
using System.Linq;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace Poctpnc
{
    public partial class FrmPoctpnc : FormTrans
    {
        public static int iRow = 0;
        public static int OldiRow = 0;
        public string Old_ma_kho = string.Empty;
        private bool txtDiaChiFocusable = true;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject Voucher_Lan0;
        private DataSet DsVitual;
        private DataSet dsCheckData;


        public FrmPoctpnc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.cbMa_nt;
            this.C_So_ct = this.txtSo_ct;
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.BindingSasObj = StartupBase.SasObj;
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmPoctpnc.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpnc.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmPoctpnc.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtso_du_kh.Visibility = this.tblso_du_kh.Visibility = Visibility.Hidden;
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCp.Lan = StartUpTrans.M_LAN;
                this.GrdCtgt.Lan = StartUpTrans.M_LAN;
                this.LanguageProvider.Language = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCp, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCtgt, StartUpTrans.Ma_ct, 2);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                }
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.loaddataDu13();
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
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

                if (this.GrdCp.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_vv_i")))
                {
                    this.GrdCp.FieldLayouts[0].Fields["ma_vv_i"].Settings.AllowEdit = new bool?(false);
                    this.GrdCp.FieldLayouts[0].Fields["ma_vv_i"].Settings.EditorStyle = (Style)null;
                }
                if (this.GrdCp.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_hdm_i")))
                {
                    this.GrdCp.FieldLayouts[0].Fields["ma_hdm_i"].Settings.AllowEdit = new bool?(false);
                    this.GrdCp.FieldLayouts[0].Fields["ma_hdm_i"].Settings.EditorStyle = (Style)null;
                }
                if (this.GrdCp.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_td_i")))
                {
                    this.GrdCp.FieldLayouts[0].Fields["ma_td_i"].Settings.AllowEdit = new bool?(false);
                    this.GrdCp.FieldLayouts[0].Fields["ma_td_i"].Settings.EditorStyle = (Style)null;
                }
                if (this.GrdCp.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_td2_i")))
                {
                    this.GrdCp.FieldLayouts[0].Fields["ma_td2_i"].Settings.AllowEdit = new bool?(false);
                    this.GrdCp.FieldLayouts[0].Fields["ma_td2_i"].Settings.EditorStyle = (Style)null;
                }
                if (this.GrdCp.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ma_td3_i")))
                {
                    this.GrdCp.FieldLayouts[0].Fields["ma_td3_i"].Settings.AllowEdit = new bool?(false);
                    this.GrdCp.FieldLayouts[0].Fields["ma_td3_i"].Settings.EditorStyle = (Style)null;
                }
                string str = StartupBase.SasObj.GetOption("M_DC_THUE_CK").ToString();
                if (!(str == "2") && !(str == "3") && this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ten_vt")))
                {
                    this.GrdCtgt.FieldLayouts[0].Fields["ten_vt"].Visibility = Visibility.Hidden;
                    this.GrdCtgt.FieldLayouts[0].Fields["ten_vt"].Width = new FieldLength?(new FieldLength(0.0));
                    this.GrdCtgt.FieldLayouts[0].Fields["dvt"].Visibility = Visibility.Collapsed;
                    this.GrdCtgt.FieldLayouts[0].Fields["dvt"].Width = new FieldLength?(new FieldLength(0.0));
                    this.GrdCtgt.FieldLayouts[0].Fields["so_luong"].Visibility = Visibility.Collapsed;
                    this.GrdCtgt.FieldLayouts[0].Fields["so_luong"].Width = new FieldLength?(new FieldLength(0.0));
                    this.GrdCtgt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Collapsed;
                    this.GrdCtgt.FieldLayouts[0].Fields["gia"].Width = new FieldLength?(new FieldLength(0.0));
                    this.GrdCtgt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Collapsed;
                    this.GrdCtgt.FieldLayouts[0].Fields["gia_nt"].Width = new FieldLength?(new FieldLength(0.0));
                }
                if (str == "1" || str == "3" || !this.GrdCtgt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "dia_chi")))
                    return;
                this.GrdCtgt.FieldLayouts[0].Fields["dia_chi"].Visibility = Visibility.Hidden;
                this.GrdCtgt.FieldLayouts[0].Fields["dia_chi"].Width = new FieldLength?(new FieldLength(0.0));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void LoadData()
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout10.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout20.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout21.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdTongChiPhi.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCp.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.GrdCtgt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[2].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtStatus.IsEnabled = false;
        }

        private void V_Dau()
        {
            FrmPoctpnc.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Truoc()
        {
            if (FrmPoctpnc.iRow <= 1)
                return;
            --FrmPoctpnc.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Sau()
        {
            if (FrmPoctpnc.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmPoctpnc.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Cuoi()
        {
            FrmPoctpnc.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString() + "'";
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)str;
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["loai_pb"] = (object)1;
                row["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                row["ma_nx"] = (object)DBNull.Value;
                row["so_pn"] = (object)DBNull.Value;
                row["ngay_pn"] = (object)DBNull.Value;
                row["status"] = StartUpTrans.DmctInfo["ma_post"];
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["ma_qs"].ToString().Trim());
                }
                row["sua_tien"] = (object)0;
                row["ma_gd"] = StartUpTrans.DsTrans.Tables[0].Rows.Count > 1 ? (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString() : (object)StartUpTrans.DmctInfo["ma_gd"].ToString();
                row["ty_giaf"] = !row["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)StartUp.GetRates(row["ma_nt"].ToString().Trim(), Convert.ToDateTime(row["ngay_ct"]).Date) : (object)1;
                row["status"] = StartUpTrans.DmctInfo["ma_post"];
                row["t_cp_nt"] = (object)0;
                row["t_cp"] = (object)0;
                row["t_thue_nt"] = (object)0;
                row["t_thue"] = (object)0;
                row["t_tt_nt"] = (object)0;
                row["t_tt"] = (object)0;
                row["t_so_luong"] = (object)0;
                row["tao_pc"] = (object)0;
                row["stt_rec_pc"] = "";
                row["so_ct_pc"] = "";
                row["ma_ct_pc"] = "";
                row["ma_qs_pc"] = "";
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.NewRowCt();
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.txtngay_lct.Text = "";
                FrmPoctpnc.OldiRow = FrmPoctpnc.iRow;
                FrmPoctpnc.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpnc.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
                this.ChkSuaTien.IsChecked = new bool?(false);
                this.ChkTaoPc.IsEnabled = false;
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
            FrmPoctpncCopy frmPoctpncCopy = new FrmPoctpncCopy();
            frmPoctpncCopy.Closed += new EventHandler(this._formcopy_Closed);
            if (this.M_LAN != "V")
                frmPoctpncCopy.Title = "Copy";
            frmPoctpncCopy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!FrmPoctpncCopy.isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow].ItemArray;
                row1["stt_rec"] = (object)str;
                row1["ngay_ct"] = (object)FrmPoctpncCopy.ngay_ct;
                if (StartUpTrans.M_ngay_lct.Equals("0") && !string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                    row1["ngay_lct"] = (object)FrmPoctpncCopy.ngay_ct;
                row1["t_thue_nt"] = (object)0;
                row1["t_thue"] = (object)0;
                row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];
                row1["stt_rec_pc"] = "";
                row1["so_ct_pc"] = "";
                row1["ma_ct_pc"] = "";
                row1["ma_qs_pc"] = "";
                row1["tao_pc"] = (object)0;
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
                FrmPoctpnc.OldiRow = FrmPoctpnc.iRow;
                FrmPoctpnc.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmPoctpnc.IsInEditMode.Value = true;
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
        }

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num1 = (int)ExMessageBox.Show(815, StartupBase.SasObj, "Không có dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)))
            {
                int num2 = (int)ExMessageBox.Show(820, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                FormTrans.currActionTask = ActionTask.Edit;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                FrmPoctpnc.IsInEditMode.Value = true;
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                refreshTao_pc();
            }
        }

        private void V_Huy()
        {
            FrmPoctpnc.IsInEditMode.Value = false;
            if (this.DsVitual == null || StartUpTrans.DsTrans.Tables[0].Rows.Count <= 0)
                return;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmPoctpnc.iRow = FrmPoctpnc.OldiRow;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString());
                        break;
                    }
                    break;
                case ActionTask.Edit:
                    FormTrans.currActionTask = ActionTask.View;
                    string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                    }
                    if (StartUpTrans.DsTrans.Tables[2].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                    }
                    StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmPoctpnc.iRow);
                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row1, FrmPoctpnc.iRow);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    StartUpTrans.DsTrans.Tables[2].Merge(this.DsVitual.Tables[2]);
                    if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim() == "")
                        this.ChkTaoPc.IsEnabled = false;
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
                FrmPoctpnc.iRow = FrmPoctpnc.iRow > 0 ? FrmPoctpnc.iRow - 1 : 0;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec='" + _stt_rec + "'");
                StartUpTrans.DsTrans.Tables[0].Rows.Remove(dataRowArray[0]);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[2].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    FrmPoctpnc.iRow = FrmPoctpnc.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmPoctpnc.iRow - 1 : FrmPoctpnc.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpnc.iRow]["stt_rec"].ToString());
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
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString()).ToString();
            if (this.M_LAN != "V")
                formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString()).ToString();
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "PoctpncXemView";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string str = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmPoctpnc.iRow = index + 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_In()
        {
            FrmIn frmIn = new FrmIn();
            if (StartUpTrans.M_LAN != "V")
                frmIn.Title = "Report form list";
            frmIn.ShowDialog();
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (StartUpTrans.DsTrans.Tables[0].DefaultView.Count > 0)
            {
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                if (!menuItemName.Equals("btnSave"))
                    this.loaddataDu13();
            }
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count <= 0)
                return;
            this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
        }

        public DataRow NewRowCt()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                dataRow["ma_kho_i"] = (object)DBNull.Value;
                dataRow["ten_vt"] = (object)DBNull.Value;
                dataRow["dvt"] = (object)DBNull.Value;
                dataRow["so_luong"] = (object)0;
                dataRow["tien_nt0"] = (object)0;
                dataRow["tien0"] = (object)0;
                dataRow["cp_nt"] = (object)0;
                dataRow["cp"] = (object)0;
                dataRow["tk_vt"] = (object)DBNull.Value;
                return dataRow;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return (DataRow)null;
        }

        private bool NewRowCtGt()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[2].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                int num1 = 0;
                int result = 0;
                if (this.GrdCtgt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (num1 >= result ? num1 : result) + 1;
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num2);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ma_ms"] = (object)StartUp.M_MA_MS;
                dataRow["ngay_ct"] = this.txtNgay_ct.Value;
                dataRow["so_luong"] = (object)0;
                dataRow["gia_nt"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["t_tien_nt"] = (object)0;
                dataRow["t_tien"] = (object)0;
                dataRow["thue_suat"] = (object)0;
                dataRow["t_thue_nt"] = (object)0;
                dataRow["t_thue"] = (object)0;
                dataRow["ma_vv"] = (object)"";
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

        private bool GrdCtgt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            return this.NewRowCtGt();
        }

        private void GrdCtgt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                Decimal num1;
                Decimal num2;
                Decimal num3;
                if (this.IsEditMode && this.GrdCtgt.ActiveCell != null && StartUpTrans.DsTrans.Tables[2].DefaultView.Count > this.GrdCtgt.ActiveRecord.Index && StartUpTrans.DsTrans.Tables[2].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "so_ct0":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Cell.Record.Cells["ma_kh"].Value == DBNull.Value || e.Cell.Record.Cells["ma_kh"].Value != null && string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kh"].Value.ToString().Trim()))
                                {
                                    e.Cell.Record.Cells["ma_kh"].Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"];
                                    if (this.txtMa_kh.RowResult == null)
                                        this.txtMa_kh.SearchInit();
                                    if (this.txtMa_kh.RowResult != null)
                                    {
                                        e.Cell.Record.Cells["dia_chi_dmkh"].Value = this.txtMa_kh.RowResult["dia_chi"];
                                        e.Cell.Record.Cells["ma_so_thue_dmkh"].Value = this.txtMa_kh.RowResult["ma_so_thue"];
                                        DataRowView dataItem = e.Cell.Record.DataItem as DataRowView;
                                        if (string.IsNullOrEmpty(dataItem["ma_thck"].ToString().Trim()))
                                            dataItem["ma_thck"] = this.txtMa_kh.RowResult["ma_thck"];
                                        if (this.ParseInt((object)dataItem["han_tt"].ToString(), 0) == 0)
                                            dataItem["han_tt"] = (object)this.ParseInt(this.txtMa_kh.RowResult["han_tt"], 0);
                                    }
                                    e.Cell.Record.Cells["ten_kh"].Value = StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] : StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"];
                                    e.Cell.Record.Cells["dia_chi"].Value = string.IsNullOrEmpty(this.txtDia_chi.Text.Trim()) ? StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] : (object)this.txtDia_chi.Text;
                                    e.Cell.Record.Cells["ma_so_thue"].Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"];
                                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                                    {
                                        e.Cell.Record.Cells["ten_vt"].Value = (object)StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ten_vt"].ToString();
                                        StartUpTrans.DsTrans.Tables[2].DefaultView[e.Cell.Record.Index]["ma_vv"] = (object)StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vv_i"].ToString();
                                        StartUpTrans.DsTrans.Tables[2].DefaultView[e.Cell.Record.Index]["ma_phi"] = (object)StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_phi_i"].ToString();
                                    }
                                }
                                Decimal result1 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[2].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[2].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result2);
                                if (!e.Cell.Record.Cells["t_tien_nt"].Value.ToString().Equals(""))
                                {
                                    result1 = SysFunc.Round(result1 - Decimal.Parse(e.Cell.Record.Cells["t_tien_nt"].Value.ToString()), StartUpTrans.M_ROUND_NT);
                                    result2 = SysFunc.Round(result2 - Decimal.Parse(e.Cell.Record.Cells["t_tien"].Value.ToString()), StartUpTrans.M_ROUND);
                                }
                                Decimal num4 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString());
                                Decimal num5 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString());
                                if (this.ParseDecimal((object)e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), new Decimal(0)) == new Decimal(0))
                                    e.Cell.Record.Cells["t_tien_nt"].Value = (object)SysFunc.Round(num4 - result1, StartUpTrans.M_ROUND_NT);
                                if (this.ParseDecimal((object)e.Cell.Record.Cells["t_tien"].Value.ToString(), new Decimal(0)) == new Decimal(0))
                                    e.Cell.Record.Cells["t_tien"].Value = (object)SysFunc.Round(num5 - result2, StartUpTrans.M_ROUND);
                                Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                Decimal num7 = SysFunc.Round(num6 * (num4 - result1) / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                Decimal num8 = SysFunc.Round(num6 * (num5 - result2) / new Decimal(100), StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["t_thue_nt"].Value = (object)num7;
                                e.Cell.Record.Cells["t_thue"].Value = (object)num8;
                                Decimal result3 = new Decimal(0);
                                Decimal result4 = new Decimal(0);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[2].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result3);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[2].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result4);
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)result3;
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)result4;
                                e.Cell.Record.Cells["t_tt_nt"].Value = (object)(num7 + (num4 - result1));
                                e.Cell.Record.Cells["t_tt"].Value = (object)(num8 + (num5 - result2));
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "ma_kh":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null && e.Editor.Value.ToString().Trim() != "")
                            {
                                if (this.M_LAN == "V")
                                    e.Cell.Record.Cells["ten_kh"].Value = autoCompleteControl1.RowResult["ten_kh"];
                                else
                                    e.Cell.Record.Cells["ten_kh"].Value = autoCompleteControl1.RowResult["ten_kh2"];
                                if (e.Cell.Record.Cells["ma_thck"].Value == DBNull.Value || string.IsNullOrEmpty(e.Cell.Record.Cells["ma_thck"].Value.ToString().Trim()))
                                {
                                    e.Cell.Record.Cells["ma_thck"].Value = autoCompleteControl1.RowResult["ma_thck"];
                                    AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_thck"]).Editor as ControlHostEditor);
                                    if (autoCompleteControl2 != null)
                                    {
                                        autoCompleteControl2.SearchInit();
                                        if (autoCompleteControl2.RowResult != null)
                                            e.Cell.Record.Cells["han_tt"].Value = autoCompleteControl2.RowResult["han_tt"];
                                    }
                                }
                                if (!string.IsNullOrEmpty(autoCompleteControl1.RowResult["dia_chi"].ToString()))
                                    e.Cell.Record.Cells["dia_chi"].Value = autoCompleteControl1.RowResult["dia_chi"];
                                if (!string.IsNullOrEmpty(autoCompleteControl1.RowResult["ma_so_thue"].ToString()))
                                    e.Cell.Record.Cells["ma_so_thue"].Value = autoCompleteControl1.RowResult["ma_so_thue"];
                                e.Cell.Record.Cells["dia_chi_dmkh"].Value = autoCompleteControl1.RowResult["dia_chi"];
                                e.Cell.Record.Cells["ma_so_thue_dmkh"].Value = autoCompleteControl1.RowResult["ma_so_thue"];
                                break;
                            }
                            break;
                        case "so_luong":
                        case "gia_nt":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["so_luong"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                num3 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                                Decimal result = new Decimal(0);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result);
                                if (this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) == new Decimal(0))
                                {
                                    e.Cell.Record.Cells["t_tien_nt"].Value = (object)SysFunc.Round(num4 * num5, StartUpTrans.M_ROUND);
                                }

                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "t_tien_nt":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["t_tien_nt"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0));
                                Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(num4 * num5 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["t_thue"].Value = e.Cell.Record.Cells["t_thue_nt"].Value;
                                    e.Cell.Record.Cells["t_tien"].Value = (object)num4;
                                    e.Cell.Record.Cells["t_tt_nt"].Value = (object)(this.ParseDecimal((object)e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), new Decimal(0)) + this.ParseDecimal((object)e.Cell.Record.Cells["t_thue_nt"].Value.ToString(), new Decimal(0)));
                                    e.Cell.Record.Cells["t_tt"].Value = e.Cell.Record.Cells["t_tt_nt"].Value;
                                }
                                else
                                {
                                    bool? isChecked = this.ChkSuaTien.IsChecked;
                                    if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                                    {
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(num4 * num5 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_tt_nt"].Value = (object)(this.ParseDecimal((object)e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), new Decimal(0)) + this.ParseDecimal((object)e.Cell.Record.Cells["t_thue_nt"].Value.ToString(), new Decimal(0)));
                                    }
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "t_tien":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["t_tien"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                num3 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0));
                                Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                Decimal result = new Decimal(0);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result);
                                if (this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0)) == new Decimal(0))
                                {
                                    e.Cell.Record.Cells["t_tien"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) * result, StartUpTrans.M_ROUND);
                                    num4 = this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0));
                                }
                                e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(num4 * num5 / new Decimal(100), StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["t_tt"].Value = (object)(this.ParseDecimal((object)e.Cell.Record.Cells["t_tien"].Value.ToString(), new Decimal(0)) + this.ParseDecimal((object)e.Cell.Record.Cells["t_thue"].Value.ToString(), new Decimal(0)));
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "ma_thue":
                            AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl3.IsDataChanged)
                            {
                                num1 = new Decimal(0);
                                num3 = new Decimal(0);
                                num2 = new Decimal(0);
                                if (autoCompleteControl3.RowResult != null)
                                {
                                    e.Cell.Record.Cells["thue_suat"].Value = autoCompleteControl3.RowResult["thue_suat"];
                                    Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0));
                                    Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0));
                                    Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                    if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    {
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(num4 * num6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_thue"].Value = e.Cell.Record.Cells["t_thue_nt"].Value;
                                        e.Cell.Record.Cells["t_tt_nt"].Value = (object)SysFunc.Round(num4 + this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_tt"].Value = e.Cell.Record.Cells["t_tt_nt"].Value;
                                    }
                                    else
                                    {
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(num4 * num6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(num5 * num6 / new Decimal(100), StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["t_tt_nt"].Value = (object)SysFunc.Round(num4 + this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_tt"].Value = (object)SysFunc.Round(num5 + this.ParseDecimal(e.Cell.Record.Cells["t_thue"].Value, new Decimal(0)), StartUpTrans.M_ROUND);
                                    }
                                    e.Cell.Record.Cells["tk_thue_no"].Value = autoCompleteControl3.RowResult["tk_thue_no"];
                                    if (e.Cell.Record.Index == 0)
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_no"] = (object)e.Cell.Record.Cells["tk_thue_no"].Value.ToString();
                                    AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["tk_thue_no"]).Editor as ControlHostEditor);
                                    if (autoCompleteControl2.RowResult == null)
                                        autoCompleteControl2.SearchInit();
                                    if (autoCompleteControl2.RowResult != null)
                                        e.Cell.Record.Cells["tk_cn"].Value = autoCompleteControl2.RowResult["tk_cn"];
                                    this.Sum_ALL();
                                }
                                break;
                            }
                            break;
                        case "t_thue_nt":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["t_thue_nt"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal num4 = new Decimal(0);
                                Decimal nValue = this.txtTy_gia.nValue;
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    if (this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)) == new Decimal(0))
                                    {
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) * this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0)) / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["t_thue"].Value = e.Cell.Record.Cells["t_thue_nt"].Value;
                                    }
                                    e.Cell.Record.Cells["t_thue"].Value = e.Cell.Record.Cells["t_thue_nt"].Value;
                                    e.Cell.Record.Cells["t_tt_nt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) + this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["t_tt"].Value = e.Cell.Record.Cells["t_tt_nt"].Value;
                                }
                                else
                                {
                                    if (this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)) == new Decimal(0))
                                        e.Cell.Record.Cells["t_thue_nt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) * this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0)) / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    if (this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)) * nValue != new Decimal(0))
                                        e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)) * nValue, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["t_tt_nt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien_nt"].Value, new Decimal(0)) + this.ParseDecimal(e.Cell.Record.Cells["t_thue_nt"].Value, new Decimal(0)), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["t_tt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0)) + this.ParseDecimal(e.Cell.Record.Cells["t_thue"].Value, new Decimal(0)), StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "t_thue":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["t_thue"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.ParseDecimal(e.Cell.Record.Cells["t_thue"].Value, new Decimal(0)) == new Decimal(0))
                                    e.Cell.Record.Cells["t_thue"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0)) * this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0)) / new Decimal(100), StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["t_tt"].Value = (object)SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["t_tien"].Value, new Decimal(0)) + this.ParseDecimal(e.Cell.Record.Cells["t_thue"].Value, new Decimal(0)), StartUpTrans.M_ROUND);
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "tk_thue_no":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl4.RowResult != null)
                                e.Cell.Record.Cells["tk_cn"].Value = autoCompleteControl4.RowResult["tk_cn"];
                            if (e.Cell.Record.Index == 0)
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_no"] = (object)e.Cell.Record.Cells["tk_thue_no"].Value.ToString();
                                break;
                            }
                            break;
                        case "ma_kh2":
                            if (e.Editor.Value == null)
                                break;
                            break;
                        case "ma_thck":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl5 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl5.RowResult != null)
                            {
                                e.Cell.Record.Cells["han_tt"].Value = autoCompleteControl5.RowResult["han_tt"];
                                break;
                            }
                            e.Cell.Record.Cells["han_tt"].Value = (object)0;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCtgt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCtgt_KeyUp(object sender, KeyEventArgs e)
        {
            if (FrmPoctpnc.IsInEditMode.Value)
            {
                Key key = e.Key;
                if (key != Key.F4)
                {
                    if (key == Key.F8)
                    {
                        if (ExMessageBox.Show(825, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.No)
                        {
                            DataRecord dataRecord = this.GrdCtgt.ActiveRecord as DataRecord;
                            if (dataRecord != null)
                            {
                                Cell activeCell = this.GrdCtgt.ActiveCell;
                                if (dataRecord.Index == 0)
                                {
                                    if (this.GrdCtgt.Records.Count == 1)
                                    {
                                        this.GrdCtgt_AddNewRecord(null, null);
                                    }
                                }
                                else if (dataRecord.Index == this.GrdCtgt.Records.Count - 1)
                                {
                                    int num = dataRecord.Index - 1;
                                }
                                int num2 = (this.GrdCtgt.ActiveCell == null) ? 0 : this.GrdCtgt.ActiveCell.Field.Index;
                                this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                                if (num2 >= 0)
                                {
                                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(StartUpTrans.DsTrans.Tables[2].DefaultView[dataRecord.Index].Row);
                                    StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                                    if (this.GrdCtgt.Records.Count > 0)
                                    {
                                        this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
                                    }
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue_nt)", StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue)", StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter);
                                    this.Sum_ALL();
                                }
                            }
                        }
                    }
                }
                else
                {
                    DataRecord dataRecord2 = this.GrdCtgt.ActiveRecord as DataRecord;
                    if (dataRecord2 != null && dataRecord2.Cells["so_ct0"].Value != null && !(dataRecord2.Cells["so_ct0"].Value.ToString() == ""))
                    {
                        this.NewRowCtGt();
                        this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
                        this.GrdCtgt.ActiveCell = (this.GrdCtgt.ActiveRecord as DataRecord).Cells["ma_ms"];
                    }
                }
            }
        }

        private void GrdCtgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpnc.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                this.NewRowCtGt();
                this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                return;
            this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            e.Handled = true;
        }

        private void ChkSuaTien_Click(object sender, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            bool? isChecked = this.ChkSuaTien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.TyGiaValueChange();
        }

        private void V_Nhan()
        {
            try
            {
                StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                bool flag1 = false;
                if (!this.IsSequenceSave)
                {
                    object obj1 = (object)this.ParseDecimal(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"), new Decimal(0));
                    object obj2 = (object)this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"], new Decimal(0));
                    object obj3 = (object)this.ParseDecimal(StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"), new Decimal(0));
                    object obj4 = (object)this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"], new Decimal(0));
                    Decimal result1 = new Decimal(0);
                    Decimal result2 = new Decimal(0);
                    Decimal result3 = new Decimal(0);
                    Decimal result4 = new Decimal(0);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result3);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result2);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result4);
                    this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                    this.GrdCp.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(830, StartupBase.SasObj, "Chưa vào mã khách hàng!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag1 = true;
                    }
                    else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(835, StartupBase.SasObj, "Chưa vào tài khoản có!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_nx.IsFocus = true;
                        flag1 = true;
                    }
                    else if (string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                    {
                        int num = (int)ExMessageBox.Show(840, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                            flag1 = true;
                            this.txtNgay_ct.Focus();
                        }
                        else if (string.IsNullOrEmpty(this.txtMa_qs.Text.ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(841, StartupBase.SasObj, "Chưa vào quyển sổ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtMa_qs.IsFocus = true;
                            flag1 = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(845, StartupBase.SasObj, "Chưa vào số chứng từ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag1 = true;
                        }
                        else if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(850, StartupBase.SasObj, "Chưa chọn phiếu nhập, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCp.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCp.Focus();
                            flag1 = true;
                        }
                        else if (result1 != result2 || result3 != result4)
                        {
                            int num2 = (int)ExMessageBox.Show(860, StartupBase.SasObj, "Tổng chi phí khác với chi phí tổng cộng của các vật tư!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                            {
                                this.GrdCp.ActiveCell = (this.GrdCp.Records[0] as DataRecord).Cells["cp_nt"];
                                this.GrdCp.Focus();
                            }
                            flag1 = true;
                        }
                    }
                    if (!flag1)
                    {
                        if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                        {
                            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                            {
                                if (string.IsNullOrEmpty(dataRowView.Row["ma_ms"].ToString().Trim()) || string.IsNullOrEmpty(dataRowView.Row["so_ct0"].ToString().Trim()))
                                {
                                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(dataRowView.Row);
                                    StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                                }
                            }
                        }
                        if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                        {
                            bool flag2 = false;
                            int num1 = 0;
                            bool flag3 = false;
                            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                            {
                                if (!StartUpTrans.M_MST_CHECK.Equals("0") && (!SysFunc.CheckSumMaSoThue(dataRowView.Row["ma_so_thue"].ToString().Trim()) && !string.IsNullOrEmpty(dataRowView.Row["ma_so_thue"].ToString().Trim()) && !flag2))
                                {
                                    int num2 = (int)ExMessageBox.Show(865, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    flag2 = true;
                                    if (StartUpTrans.M_MST_CHECK.Equals("2"))
                                        return;
                                }
                                if (StartUpTrans.M_CHK_HD_VAO != 0)
                                {
                                    string so_ct0 = dataRowView.Row["so_ct0"].ToString().Trim();
                                    string so_seri0 = dataRowView.Row["so_seri0"].ToString().Trim();
                                    string str;
                                    if (!string.IsNullOrEmpty(dataRowView.Row["ngay_ct0"].ToString().Trim()))
                                    {
                                        DateTime dateTime = Convert.ToDateTime(dataRowView.Row["ngay_ct0"].ToString().Trim());
                                        dateTime = dateTime.Date;
                                        str = dateTime.ToShortDateString().Substring(0, 10);
                                    }
                                    else
                                        str = "";
                                    string ngay_ct0 = str;
                                    string ma_so_thue = dataRowView.Row["ma_so_thue"].ToString().Trim();
                                    if (StartUp.CheckExistHDVao(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString(), so_ct0, so_seri0, ngay_ct0, ma_so_thue) && !flag3)
                                    {
                                        int num2 = (int)ExMessageBox.Show(870, StartupBase.SasObj, string.Format("Hoá đơn số [{0}], ký hiệu [{1}], ngày [{2}], MST [{3}] đã tồn tại!", (object)so_ct0, (object)so_seri0, (object)ngay_ct0, (object)ma_so_thue), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag3 = true;
                                        if (StartUpTrans.M_CHK_HD_VAO == 2)
                                            return;
                                    }
                                }
                                ++num1;
                            }
                            if (!this.CheckVoucherOutofDate())
                                return;
                        }
                        if (this.GrdCtgt.Records.Count > 0)
                        {
                            Decimal num1 = Convert.ToDecimal(obj1.Equals((object)DBNull.Value) ? (object)0 : obj1);
                            Decimal num2 = Convert.ToDecimal(obj2.Equals((object)DBNull.Value) ? (object)0 : obj2);
                            Decimal num3 = Convert.ToDecimal(obj3.Equals((object)DBNull.Value) ? (object)0 : obj3);
                            Decimal num4 = Convert.ToDecimal(obj4.Equals((object)DBNull.Value) ? (object)0 : obj4);
                            Decimal num5;
                            Decimal num6;
                            Decimal num7;
                            Decimal num8;
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                num5 = SysFunc.Round(num1, StartUpTrans.M_ROUND);
                                num6 = num5;
                                num7 = SysFunc.Round(num2, StartUpTrans.M_ROUND);
                                num8 = num7;
                            }
                            else
                            {
                                num5 = SysFunc.Round(num1, StartUpTrans.M_ROUND_NT);
                                num6 = SysFunc.Round(num3, StartUpTrans.M_ROUND);
                                num7 = SysFunc.Round(num2, StartUpTrans.M_ROUND_NT);
                                num8 = SysFunc.Round(num4, StartUpTrans.M_ROUND);
                            }
                            if (num7 != num5 || num8 != num6)
                            {
                                int num9 = (int)ExMessageBox.Show(880, StartupBase.SasObj, "Tổng tiền/ tiền ngoại tệ khác với tổng tiền/ tiền ngoại tệ trong các hóa đơn giá trị gia tăng!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                        }
                    }
                }
                if (!flag1)
                {
                    if (!this.IsSequenceSave)
                    {
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Trim()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString().Trim()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs;
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            Decimal num1 = new Decimal(0);
                            Decimal num2 = new Decimal(0);
                            Decimal num3 = new Decimal(0);
                            Decimal num4 = new Decimal(0);
                            Decimal num5 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"], new Decimal(0));
                            Decimal num6 = SysFunc.Round(num5 + this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp_nt"], new Decimal(0)), StartUpTrans.M_ROUND_NT);
                            Decimal num7 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"], new Decimal(0));
                            Decimal num8 = SysFunc.Round(num7 + this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"], new Decimal(0)), StartUpTrans.M_ROUND);
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"] = (object)num8;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"] = (object)num6;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"] = (object)num7;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"] = (object)num5;
                        }
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[2].DefaultView.Count; ++index)
                        {
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["tk_du"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"];
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["ma_nt"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"];
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["ty_gia"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"];
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["ty_giaf"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_giaf"];
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["status"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
                            StartUpTrans.DsTrans.Tables[2].DefaultView[index]["ma_gd"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"];
                        }
                        if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                        {
                            int index = this.Lay_Index_Record_Co_TienThueMax();
                            if (index != -1)
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index]["han_tt"];
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index]["so_ct0"];
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index]["ngay_ct0"];
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index]["so_seri0"];
                            }
                        }
                        else
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct0"] = (object)"";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct0"] = (object)DBNull.Value;
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri0"] = (object)"";
                        }
                        Decimal num = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                        bool? isChecked = this.ChkSuaTien.IsChecked;
                        if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && num != new Decimal(0))
                            this.CanBangTien();
                        this.PhanBoThueInCT();
                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_ct"] = StartUpTrans.DmctInfo["ct_nxt"];
                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable1.Rows[0]["status"] = (object)0;
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                    DataTable LocalTable3 = StartUpTrans.DsTrans.Tables[2].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                            dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;
                        LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                            dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;
                        LocalTable3.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num9 = (int)ExMessageBox.Show(885, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctgtdbf"].ToString(), LocalTable3, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num10 = (int)ExMessageBox.Show(890, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else
                    {
                        if (!this.IsSequenceSave && !flag1)
                        {
                            this.dsCheckData = StartUp.CheckData();
                            this.dsCheckData.Tables[0].AcceptChanges();
                            if (this.dsCheckData.Tables.Count > 0)
                            {
                                foreach (DataRowView dataRowView in this.dsCheckData.Tables[0].DefaultView)
                                {
                                    if (!flag1)
                                    {
                                        switch (dataRowView[0].ToString())
                                        {
                                            case "PH01":
                                                if (StartUpTrans.M_trung_so.Equals("1"))
                                                {
                                                    if (ExMessageBox.Show(895, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                    {
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag1 = true;
                                                        break;
                                                    }
                                                    break;
                                                }
                                                if (StartUpTrans.M_trung_so.Equals("2"))
                                                {
                                                    int num1 = (int)ExMessageBox.Show(900, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    this.txtSo_ct.SelectAll();
                                                    this.txtSo_ct.Focus();
                                                    flag1 = true;
                                                    break;
                                                }
                                                break;
                                            case "PH02":
                                                int num2 = (int)ExMessageBox.Show(905, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag1 = true;
                                                this.txtMa_nx.IsFocus = true;
                                                break;
                                            case "GT01":
                                                int int16 = (int)Convert.ToInt16(dataRowView[1]);
                                                int num3 = (int)ExMessageBox.Show(910, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag1 = true;
                                                this.tiThue.Focus();
                                                this.GrdCtgt.ActiveCell = (this.GrdCtgt.Records[int16] as DataRecord).Cells["tk_thue_no"];
                                                this.GrdCtgt.Focus();
                                                break;
                                        }
                                        this.dsCheckData.Tables[0].Rows.Remove(dataRowView.Row);
                                    }
                                    else
                                        break;
                                }
                            }
                        }

                        DataTable dtPC = TaoPTC();

                        if (!flag1)
                        {
                            string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                            new Thread((ThreadStart)(() =>
                           {
                               this.Post();

                               if (dtPC != null && dtPC.Rows.Count > 0)
                               {
                                   this.CreatePC(dtPC);
                               }

                               if (this.IsSequenceSave)
                                   return;
                               this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                 {
                     if (!StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(_stt_rec1))
                         return;
                     this.loaddataDu13();
                 }));
                           })).Start();
                            if (!this.IsSequenceSave)
                            {
                                int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                if (FrmPoctpnc.iRow != pos)
                                {
                                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                    DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                    row2.ItemArray = row1.ItemArray;
                                    if (FrmPoctpnc.iRow > pos)
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                    else
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    FrmPoctpnc.iRow = pos;
                                }
                                FormTrans.currActionTask = ActionTask.View;
                                FrmPoctpnc.IsInEditMode.Value = false;
                            }
                        }
                    }
                }
                this.ChkTaoPc.IsEnabled = false;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private DataTable TaoPTC()
        {            // Tao chieu chi
            // --- Xoa PC neu khong tao
            string newstt_recPC = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim();
            string sma_ct_pc = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"].ToString().Trim();
            bool chkTao_pc = ChkTaoPc.IsChecked.HasValue ? ChkTaoPc.IsChecked.Value : false;
            bool _createPC = false;

            if (!String.IsNullOrEmpty(newstt_recPC) && !chkTao_pc)
            {
                bool perpc = SysFunc.CheckPermission(this.BindingSasObj, ActionTask.Delete, this.BindingSasObj.ExcuteScalar(new SqlCommand("Select top 1 menu_id From command Where ma_ct like '" + sma_ct_pc + "'")).ToString().Trim());
                if (perpc && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("2"))
                {
                    this.DeleteVoucherPC(newstt_recPC, sma_ct_pc);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(600, StartupBase.SasObj, "Không có quyền xóa phiếu chi đã tạo số : " + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"].ToString().Trim(), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }

                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pc"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"] = "";
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = string.Format("Update " + StartUp.Tablename + " SET ma_qs_pc='',so_ct_pc='',ma_ct_pc='',stt_rec_pc='',tao_pc = 0 WHERE stt_rec LIKE '{0}'", StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim());
                this.BindingSasObj.ExcuteScalar(sqlcmd);
            }

            if (!chkTao_pc)
                return (DataTable)null;

            DataTable dt = new DataTable();
            dt.Columns.Add("ma_ct", typeof(string));
            dt.Columns.Add("stt_rec", typeof(string));
            dt.Columns.Add("stt_recPC", typeof(string));
            dt.Columns.Add("ma_qs", typeof(string));
            dt.Columns.Add("so_ct", typeof(string));
            dt.Columns.Add("ma_nt", typeof(string));
            dt.Columns.Add("ty_gia", typeof(Decimal));
            dt.Columns.Add("ty_giaf", typeof(Decimal));
            dt.Columns.Add("nguoinop", typeof(string));
            dt.Columns.Add("lydonop", typeof(string));
            dt.Columns.Add("ma_gd", typeof(string));

            DataRow row = dt.NewRow();
            dt.Rows.Add(row);

            DataTable dataTable = (DataTable)null;
            if (!String.IsNullOrEmpty(newstt_recPC))
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai,ty_gia,ty_giaf FROM {0} WHERE stt_rec LIKE '{1}'", sma_ct_pc.Equals("PC1") ? "ph46" : "ph56", newstt_recPC);
                dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];
            }


            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                _createPC = true;
                dt.Rows[0]["ma_ct"] = dataTable.Rows[0]["ma_ct"].ToString().Trim();
                dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dt.Rows[0]["stt_recPC"] = dataTable.Rows[0]["stt_rec"].ToString().Trim();
                dt.Rows[0]["ma_qs"] = dataTable.Rows[0]["ma_qs"].ToString().Trim();
                dt.Rows[0]["so_ct"] = dataTable.Rows[0]["so_ct"].ToString().Trim();
                dt.Rows[0]["ma_nt"] = dataTable.Rows[0]["ma_nt"].ToString().Trim();
                dt.Rows[0]["ty_gia"] = this.ParseDecimal(dataTable.Rows[0]["ty_gia"], new Decimal(0));
                dt.Rows[0]["ty_giaf"] = this.ParseDecimal(dataTable.Rows[0]["ty_giaf"], new Decimal(0));
                dt.Rows[0]["nguoinop"] = dataTable.Rows[0]["ong_ba"].ToString().Trim();
                dt.Rows[0]["lydonop"] = dataTable.Rows[0]["dien_giai"].ToString().Trim();
                dt.Rows[0]["ma_gd"] = dataTable.Rows[0]["ma_gd"].ToString().Trim();
            }
            else
            {
                if (!String.IsNullOrEmpty(newstt_recPC))
                    this.DeleteVoucherPC(newstt_recPC, sma_ct_pc);
            }

            if (String.IsNullOrEmpty(newstt_recPC))
            {
                FrmTaoPC frmTaoPt = new FrmTaoPC();
                frmTaoPt.tbInfoPC = dataTable;
                frmTaoPt.DataContext = dt.DefaultView;
                frmTaoPt.txtMa_qs_pc.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pc"].ToString();
                frmTaoPt.txtso_ct_pc.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"].ToString().Trim().PadLeft(frmTaoPt.txtso_ct_pc.MaxLength);
                frmTaoPt.txtnguoi_nop.Text = this.txtOng_ba.Text;
                frmTaoPt.kind = 1;
                frmTaoPt.Ma_nt_ht = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                frmTaoPt.so_hd = this.txtSo_ct.Text.Trim();
                frmTaoPt.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
                frmTaoPt.filterma_qs = this.txtMa_qs.Filter;
                frmTaoPt.ShowDialog();
                if (!frmTaoPt.isOk)
                {
                    _createPC = false;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pc"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tao_pc"] = (Object)0;
                    SqlCommand sqlcmd = new SqlCommand();
                    sqlcmd.CommandText = string.Format("Update " + StartUp.Tablename + " SET tao_pc = 0 WHERE stt_rec LIKE '{0}'", StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim());
                    this.BindingSasObj.ExcuteScalar(sqlcmd);
                }
                else
                {
                    dt = dt.Copy();
                    newstt_recPC = !frmTaoPt.txtKind.Text.Equals("1") ? DataProvider.NewTrans(StartupBase.SasObj, "BN1", StartUpTrans.Ws_Id) : DataProvider.NewTrans(StartupBase.SasObj, "PC1", StartUpTrans.Ws_Id);
                    dt.Rows[0]["ma_ct"] = frmTaoPt.txtKind.Text.Equals("1") ? "PC1" : "BN1";
                    dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                    dt.Rows[0]["stt_recPC"] = newstt_recPC;
                    dt.Rows[0]["ma_qs"] = frmTaoPt.txtMa_qs_pc.Text;
                    dt.Rows[0]["so_ct"] = frmTaoPt.txtso_ct_pc.Text.PadLeft(frmTaoPt.txtso_ct_pc.MaxLength, ' ');
                    dt.Rows[0]["ma_nt"] = frmTaoPt.txtMa_nt.Text;
                    dt.Rows[0]["ty_gia"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.Rate;
                    dt.Rows[0]["ty_giaf"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.RateF;
                    dt.Rows[0]["nguoinop"] = frmTaoPt.txtnguoi_nop.Text;
                    dt.Rows[0]["lydonop"] = frmTaoPt.txtlydo_nop.Text;
                    dt.Rows[0]["ma_gd"] = frmTaoPt.txtMa_gd.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pc"] = frmTaoPt.txtMa_qs_pc.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"] = frmTaoPt.txtso_ct_pc.Text.Trim().PadLeft(frmTaoPt.txtso_ct_pc.MaxLength);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"] = frmTaoPt.txtKind.Text.Equals("1") ? "PC1" : "BN1";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"] = newstt_recPC;
                }
            }

            bool perpac = SysFunc.CheckPermission(this.BindingSasObj, ActionTask.Add, this.BindingSasObj.ExcuteScalar(new SqlCommand("Select top 1 menu_id From command Where ma_ct like '" + dt.Rows[0]["ma_ct"].ToString().Trim() + "'")).ToString().Trim());
            if (!perpac || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("2"))
            {
                int num1 = (int)ExMessageBox.Show(605, StartupBase.SasObj, "Không có quyền tạo phiếu chi!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return (DataTable)null;
            }

            return dt;
        }
        public void DeleteVoucherPC(string _stt_rec, string _ma_ct)
        {
            try
            {
                string format = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd = new SqlCommand(string.Format(format, "[DeleteVoucher]"));
                sqlcmd.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = _ma_ct;
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = _stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void refreshTao_pc()
        {
            this.ChkTaoPc.IsEnabled = Inlist(this.txtMa_nx.Text.Trim(), StartupBase.SasObj.GetOption("M_TK_TK_VT").ToString().Trim().Split(','));
            if (!this.ChkTaoPc.IsEnabled)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tao_pc"] = (Object)0;
            }
        }

        private bool Inlist(string value, string[] list)
        {
            foreach (string item in list)
            {
                if (value.StartsWith(item))
                    return true;
            }
            return false;
        }
        private void CreatePC(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[POCTPNC-CREATEPC] @Stt_rec, @Stt_recPC, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct");
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = dt.Rows[0]["stt_rec"];
                sqlcmd.Parameters.Add("@Stt_recPC", SqlDbType.VarChar).Value = dt.Rows[0]["stt_recPC"];
                sqlcmd.Parameters.Add("@ma_qs", SqlDbType.VarChar).Value = dt.Rows[0]["ma_qs"];
                sqlcmd.Parameters.Add("@so_ct", SqlDbType.VarChar).Value = dt.Rows[0]["so_ct"];
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nt"];
                sqlcmd.Parameters.Add("@ty_gia", SqlDbType.Decimal).Value = dt.Rows[0]["ty_gia"];
                sqlcmd.Parameters.Add("@ty_giaf", SqlDbType.Decimal).Value = dt.Rows[0]["ty_giaf"];
                sqlcmd.Parameters.Add("@nguoinop", SqlDbType.NVarChar).Value = dt.Rows[0]["nguoinop"];
                sqlcmd.Parameters.Add("@lydonop", SqlDbType.NVarChar).Value = dt.Rows[0]["lydonop"];
                sqlcmd.Parameters.Add("@ma_gd", SqlDbType.VarChar).Value = dt.Rows[0]["ma_gd"];
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = dt.Rows[0]["ma_ct"];
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);             
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[POCTPNC-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }
        private void cbMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.cbMa_nt.IsDataChanged)
                return;
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (this.cbMa_nt.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.cbMa_nt.RowResult["loai_tg"];
                if (this.cbMa_nt.RowResult["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    this.txtTy_gia.Value = (object)1;
                else
                    this.txtTy_gia.Value = (object)StartUp.GetRates(this.cbMa_nt.RowResult["ma_nt"].ToString().Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
            }
            this.TyGiaValueChange();
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpnc.IsInEditMode.Value || this.txtMa_kh.RowResult == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = (object)this.txtMa_kh.RowResult["ten_kh"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = (object)this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = (object)this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"] = this.txtMa_kh.RowResult["ma_thck"];
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) ? (object)this.txtMa_kh.RowResult["tk"].ToString().Trim() : (object)this.txtMa_nx.Text.Trim();
            if (string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_kh.RowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
            this.loaddataDu13();
        }

        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            this.txtDia_chi.IsTabStop = false;
            if (Keyboard.IsKeyDown(Key.Tab) && Keyboard.Modifiers == ModifierKeys.Shift)
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Shift, Key.Tab);
            else
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_nx.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_tk"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_tk2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
                refreshTao_pc();
            }
            this.loaddataDu13();
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrdLayoutNT, KeyboardNavigationMode.Continue);
            if (this.GrdCp.Records.Count > 0)
            {
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
                this.GrdCp.ActiveRecord = this.GrdCp.Records[this.GrdCp.Records.Count - 1];
            }
            else
            {
                (this.FindName("btnChonpn") as Button).Focus();
                e.Handled = true;
            }
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.View || FormTrans.currActionTask == ActionTask.None)
                return;
            if (this.txtTy_gia.OldValue != this.txtTy_gia.nValue)
                this.TyGiaValueChange();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (string.IsNullOrEmpty(this.txtso_pn.Text.Trim()))
                   this.btnChonpn.Focus();
               else
                   this.txttong_cp_nt.Focus();
           }));
        }

        public void TyGiaValueChange()
        {
            if (this.cbMa_nt.RowResult != null)
            {
                if (this.cbMa_nt.RowResult["ma_nt"].ToString() == StartUpTrans.M_ma_nt0)
                {
                    this.txtTy_gia.Value = (object)1;
                    this.txttong_cp.Visibility = Visibility.Hidden;
                }
                else
                    this.txttong_cp.Visibility = Visibility.Visible;
            }
            if (string.IsNullOrEmpty(this.txtTy_gia.Text.ToString()))
                this.txtTy_gia.Value = (object)0;
            if (this.ParseDecimal(this.txtTy_gia.Value, new Decimal(0)) == new Decimal(0))
                this.txttong_cp.IsReadOnly = false;
            else
                this.txttong_cp.IsReadOnly = true;
            try
            {
                if (FormTrans.currActionTask == ActionTask.Delete || !FrmPoctpnc.IsInEditMode.Value || (this.txtTy_gia.Value == null || this.txtTy_gia.Value == DBNull.Value || !(this.txtTy_gia.nValue != new Decimal(0))))
                    return;
                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);
                Decimal num4 = new Decimal(0);
                Decimal num5 = new Decimal(0);
                Decimal num6 = new Decimal(0);
                Decimal nValue = this.txtTy_gia.nValue;
                num4 = this.txttong_cp_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txttong_cp_nt.Value);
                if (this.GrdCp.Records.Count > 0 && (this.GrdCp.DataSource as DataView).Table.DefaultView[0]["ma_vt"] != DBNull.Value)
                {
                    Decimal num7 = this.txttong_cp_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txttong_cp_nt.Value.ToString());
                    if (this.GrdCp.Records.Count > 0)
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] = (object)SysFunc.Round(num7 * nValue, StartUpTrans.M_ROUND_NT);
                        if (!string.IsNullOrEmpty(this.txtloai_pb.Text.Trim().ToString()))
                            this.PhanBo();
                    }
                    if (this.GrdCtgt.Records.Count > 0)
                    {
                        for (int index = 0; index < this.GrdCtgt.Records.Count; ++index)
                        {
                            if ((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien_nt"].Value != DBNull.Value)
                            {
                                Decimal num8 = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien_nt"].Value);
                                (this.GrdCtgt.DataSource as DataView)[index]["t_tien"] = (object)SysFunc.Round(nValue * num8, StartUpTrans.M_ROUND_NT);
                            }
                            if ((this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue_nt"].Value != DBNull.Value)
                            {
                                Decimal num8 = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue_nt"].Value);
                                (this.GrdCtgt.DataSource as DataView)[index]["t_thue"] = (object)SysFunc.Round(nValue * num8, StartUpTrans.M_ROUND_NT);
                            }
                            if ((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien"].Value != DBNull.Value && (this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue"].Value != DBNull.Value)
                            {
                                num3 = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien"].Value);
                                num6 = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue"].Value);
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

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value == DBNull.Value)
                this.txtNgay_ct.Value = (object)DateTime.Now;
            if (this.txtNgay_ct.IsFocusWithin || FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy || (!StartUpTrans.M_ngay_lct.Equals("0") || string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString())))
                return;
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
            int num = (int)ExMessageBox.Show(920, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpnc.IsInEditMode.Value || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
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
            this.txtSo_ct.MaxLength = ListFunc.GetLengthColumn(ListFunc.GetSqlTableFieldList(StartupBase.SasObj, "v_PH73"), "so_ct");
        }

        private void Sum_ALL()
        {
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
            Decimal num5;
            Decimal num6;
            Decimal num7;
            Decimal num8;
            if (this.cbMa_nt.Text.Equals(StartUpTrans.M_ma_nt0))
            {
                num5 = SysFunc.Round(this.ParseDecimal((object)this.txttong_cp_nt.nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num6 = num5;
                num7 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue_nt)", StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num8 = num7;
            }
            else
            {
                Decimal nValue = this.txttong_cp.nValue;
                num6 = SysFunc.Round(this.ParseDecimal((object)nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                nValue = this.txttong_cp_nt.nValue;
                num5 = SysFunc.Round(this.ParseDecimal((object)nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num8 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue)", StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num7 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[2].Compute("sum(t_thue_nt)", StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] = (object)num6;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"] = (object)num5;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)num8;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num7;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(num6 + num8);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(num5 + num7);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (ma_nt == StartUpTrans.M_ma_nt0)
            {
                this.GrdCp.FieldLayouts[0].Fields["tien0"].Visibility = Visibility.Hidden;
                this.GrdCp.FieldLayouts[0].Fields["cp"].Visibility = Visibility.Hidden;
                this.GrdCp.FieldLayouts[0].Fields["tien0"].Settings.CellMaxWidth = 0.0;
                this.GrdCp.FieldLayouts[0].Fields["cp"].Settings.CellMaxWidth = 0.0;
                this.GrdCtgt.FieldLayouts[0].Fields["t_tien"].Visibility = Visibility.Hidden;
                this.GrdCtgt.FieldLayouts[0].Fields["t_thue"].Visibility = Visibility.Hidden;
                this.GrdCtgt.FieldLayouts[0].Fields["tk_cn"].Visibility = Visibility.Hidden;
                this.GrdCtgt.FieldLayouts[0].Fields["t_tien"].Settings.CellMaxWidth = 0.0;
                this.GrdCtgt.FieldLayouts[0].Fields["t_thue"].Settings.CellMaxWidth = 0.0;
                this.GrdCtgt.FieldLayouts[0].Fields["tk_cn"].Settings.CellMaxWidth = 0.0;
                this.txtTy_gia.IsReadOnly = true;
                this.txttong_cp.Visibility = Visibility.Hidden;
            }
            else
            {
                this.txttong_cp.Visibility = Visibility.Visible;
                this.GrdCp.FieldLayouts[0].Fields["tien0"].Visibility = Visibility.Visible;
                this.GrdCp.FieldLayouts[0].Fields["cp"].Visibility = Visibility.Visible;
                this.GrdCp.FieldLayouts[0].Fields["tien0"].Settings.CellMaxWidth = this.GrdCp.FieldLayouts[0].Fields["tien0"].Width.Value.Value;
                FieldSettings settings1 = this.GrdCp.FieldLayouts[0].Fields["cp"].Settings;
                FieldLength fieldLength = this.GrdCp.FieldLayouts[0].Fields["cp"].Width.Value;
                double num1 = fieldLength.Value;
                settings1.CellMaxWidth = num1;
                this.GrdCtgt.FieldLayouts[0].Fields["t_tien"].Visibility = Visibility.Visible;
                this.GrdCtgt.FieldLayouts[0].Fields["t_thue"].Visibility = Visibility.Visible;
                this.GrdCtgt.FieldLayouts[0].Fields["tk_cn"].Visibility = Visibility.Hidden;
                FieldSettings settings2 = this.GrdCtgt.FieldLayouts[0].Fields["t_tien"].Settings;
                fieldLength = this.GrdCtgt.FieldLayouts[0].Fields["t_tien"].Width.Value;
                double num2 = fieldLength.Value;
                settings2.CellMaxWidth = num2;
                FieldSettings settings3 = this.GrdCtgt.FieldLayouts[0].Fields["t_thue"].Settings;
                fieldLength = this.GrdCtgt.FieldLayouts[0].Fields["t_thue"].Width.Value;
                double num3 = fieldLength.Value;
                settings3.CellMaxWidth = num3;
                this.GrdCtgt.FieldLayouts[0].Fields["tk_cn"].Settings.CellMaxWidth = 0.0;
                if (FrmPoctpnc.IsInEditMode.Value)
                    this.txtTy_gia.IsReadOnly = false;
                else
                    this.txtTy_gia.IsReadOnly = true;
            }
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            this.ChangeLanguage();
        }

        private void IsVisibilityFieldsXamDataGridBySua_Tien()
        {
            int result1 = 0;
            int.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString(), out result1);
            switch (result1)
            {
                case 0:
                    Decimal result2 = new Decimal(0);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result2);
                    if (result2 == new Decimal(0))
                    {
                        this.txttong_cp.IsReadOnly = false;
                        this.txttong_cp.IsTabStop = true;
                        break;
                    }
                    if (this.ParseDecimal(this.txtTy_gia.Value, new Decimal(0)) == new Decimal(0))
                    {
                        this.txttong_cp.IsReadOnly = false;
                        this.txttong_cp.IsTabStop = true;
                    }
                    else
                    {
                        this.txttong_cp.IsReadOnly = true;
                        this.txttong_cp.IsTabStop = false;
                    }
                    break;
                case 1:
                    if (FrmPoctpnc.IsInEditMode.Value)
                    {
                        this.txttong_cp.IsReadOnly = false;
                        this.txttong_cp.IsTabStop = true;
                        break;
                    }
                    this.txttong_cp.IsReadOnly = true;
                    this.txttong_cp.IsTabStop = false;
                    break;
            }
            this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
        }

        private void PhanBoThueInCT()
        {
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal result5 = new Decimal(0);
            Decimal result6 = new Decimal(0);
            Decimal result7 = new Decimal(0);
            Decimal result8 = new Decimal(0);
            Decimal result9 = new Decimal(0);
            Decimal result10 = new Decimal(0);
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal result11 = new Decimal(0);
            string str = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result11);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", "stt_rec= '" + str + "'").ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", "stt_rec= '" + str + "'").ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp_nt)", "stt_rec= '" + str + "'").ToString(), out result5);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp)", "stt_rec= '" + str + "'").ToString(), out result6);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"].ToString(), out result9);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"].ToString(), out result10);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp_nt"].ToString(), out result7);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"].ToString(), out result8);
                Decimal num5;
                Decimal num6;
                if (this.cbMa_nt.Text != StartUpTrans.M_ma_nt0)
                {
                    num5 = !(result9 == new Decimal(0)) ? (result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result9 + result7) / (result3 + result5) * result1, StartUpTrans.M_ROUND_NT)) : new Decimal(0);
                    num6 = !(result10 == new Decimal(0)) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result10 + result8) / (result4 + result6) * result2, StartUpTrans.M_ROUND)) : new Decimal(0);
                }
                else
                {
                    num5 = result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result9 + result7) / (result3 + result5) * result1, StartUpTrans.M_ROUND_NT);
                    num6 = result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result10 + result8) / (result4 + result6) * result2, StartUpTrans.M_ROUND);
                }
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue_nt"] = (object)num5;
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue"] = (object)num6;
                num3 += num5;
                num4 += num6;
            }
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue_nt"].ToString()) + (result1 - num3));
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["thue"].ToString()) + (result2 - num4));
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
        }

        private void PhanBo()
        {
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal result5 = new Decimal(0);
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal result6 = new Decimal(0);
            string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result6);
            int result7 = 0;
            int.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_pb"].ToString(), out result7);
            switch (result7)
            {
                case 1:
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", "stt_rec= '" + str + "'").ToString(), out result4);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", "stt_rec= '" + str + "'").ToString(), out result5);
                    Decimal result8 = new Decimal(0);
                    Decimal result9 = new Decimal(0);
                    Decimal result10 = new Decimal(0);
                    Decimal result11 = new Decimal(0);
                    Decimal result12 = new Decimal(0);
                    for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                    {
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"].ToString(), out result8);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"].ToString(), out result9);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia0"].ToString(), out result10);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt0"].ToString(), out result11);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString(), out result12);
                        Decimal num5 = !(this.cbMa_nt.Text != StartUpTrans.M_ma_nt0) ? (result5 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result9 / result5 * result2, StartUpTrans.M_ROUND_NT)) : (!(result9 == new Decimal(0)) ? (result5 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result9 / result5 * result2, StartUpTrans.M_ROUND_NT)) : (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result8 / result4 * result2, StartUpTrans.M_ROUND_NT)));
                        Decimal num6 = !(num5 != new Decimal(0)) ? (!(result5 != new Decimal(0)) ? (!(result4 != new Decimal(0)) ? new Decimal(0) : SysFunc.Round(result8 / result4 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(result9 / result5 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(num5 * result6, StartUpTrans.M_ROUND);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"] = (object)num6;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp_nt"] = (object)num5;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia"] = (object)((result12 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num6 / result12, StartUpTrans.M_ROUND_GIA)) + result10);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"] = (object)((result12 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num5 / result12, StartUpTrans.M_ROUND_GIA_NT)) + result11);
                        num1 += num6;
                        num2 += num5;
                    }
                    break;
                case 2:
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", "stt_rec= '" + str + "'").ToString(), out result3);
                    Decimal result13 = new Decimal(0);
                    Decimal result14 = new Decimal(0);
                    Decimal result15 = new Decimal(0);
                    for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                    {
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString(), out result13);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia0"].ToString(), out result14);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt0"].ToString(), out result15);
                        Decimal num5 = result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result13 / result3 * result2, StartUpTrans.M_ROUND_NT);
                        Decimal num6 = SysFunc.Round(num5 * result6, StartUpTrans.M_ROUND);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"] = (object)num6;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp_nt"] = (object)num5;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia"] = (object)((result13 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num6 / result13, StartUpTrans.M_ROUND_GIA)) + result14);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"] = (object)((result13 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num5 / result13, StartUpTrans.M_ROUND_GIA_NT)) + result15);
                        num1 += num6;
                        num2 += num5;
                    }
                    break;
            }
            if (result7 != 1 && result7 != 2)
                return;
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp"].ToString()) + (result1 - num1));
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp_nt"].ToString()) + (result2 - num2));
        }

        private void txttong_cp_nt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            int num;
            if (!(this.txttong_cp_nt.OldValue != this.txttong_cp_nt.nValue))
            {
                bool? isChecked = this.ChkSuaTien.IsChecked;
                num = (isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 ? 1 : 0;
            }
            else
                num = 0;
            if (num != 0)
                return;
            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                this.txttong_cp.nValue = this.txttong_cp_nt.nValue;
            else if (this.txttong_cp_nt.nValue * this.txtTy_gia.nValue != new Decimal(0))
                this.txttong_cp.nValue = this.txttong_cp_nt.nValue * this.txtTy_gia.nValue;
            this.Sum_ALL();
        }

        private void txttong_cp_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result2);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(result1 + result2);
        }

        private void btnPhanBo_Click(object sender, RoutedEventArgs e)
        {
            this.PhanBo();
            int num = (int)ExMessageBox.Show(925, StartupBase.SasObj, "Đã thực hiện xong phân bổ chi phí!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count <= 0)
                return;
            this.GrdCp.Focus();
            this.GrdCp.ActiveCell = (this.GrdCp.Records[0] as DataRecord).Cells["cp_nt"];
        }

        private void GrdCp_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                switch (e.Cell.Field.Name)
                {
                    case "cp_nt":
                        if (e.Editor.Value == DBNull.Value)
                            e.Cell.Record.Cells["cp_nt"].Value = (object)0;
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal num1 = new Decimal(0);
                            Decimal num2 = new Decimal(0);
                            Decimal nValue1 = this.txtTy_gia.nValue;
                            Decimal nValue2 = (e.Editor as NumericTextBox).nValue;
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                e.Cell.Record.Cells["cp"].Value = e.Cell.Record.Cells["cp_nt"].Value;
                            else if (nValue2 * nValue1 != new Decimal(0))
                                e.Cell.Record.Cells["cp"].Value = (object)SysFunc.Round(nValue2 * nValue1, StartUpTrans.M_ROUND);
                        }
                        this.Sum_ALL();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private int Lay_Record_Co_TienHangMax()
        {
            int num1 = 0;
            double num2 = 0.0;
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (double.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"].ToString()) > num2)
                {
                    num2 = double.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"].ToString());
                    num1 = index;
                }
            }
            return num1;
        }

        private int Lay_Index_Record_Co_TienThueMax()
        {
            int num1 = -1;
            double num2 = 0.0;
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[2].DefaultView.Count; ++index)
            {
                if (double.Parse(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["t_thue"].ToString()) > num2)
                {
                    num2 = double.Parse(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["t_thue"].ToString());
                    num1 = index;
                }
            }
            return num1;
        }

        private void txtloai_pb_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(this.txtloai_pb.Text.Trim()))
                    return;
                this.txtloai_pb.Text = "1";
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void loaddataDu13()
        {
            this.txtso_du_kh.Value = (object)ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString());
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = new Decimal(0);
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
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

        public void CanBangTien()
        {
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(1);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result2);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"], new Decimal(0)) != new Decimal(0))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"] = (object)this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"], new Decimal(0));
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"], new Decimal(0)) + this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"], new Decimal(0)));
                    break;
                }
            }
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal result5 = new Decimal(0);
            Decimal result6 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result5);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result6);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(result3 + result5);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(result4 + result6);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            StartUpTrans.DsTrans.Tables[2].AcceptChanges();
        }

        private void btnChonpn_Click(object sender, RoutedEventArgs e)
        {
            if (!FrmPoctpnc.IsInEditMode.Value)
                return;
            try
            {
                FrmTimPN frmTimPn = new FrmTimPN(this.txtMa_kh.Text.Trim());
                frmTimPn.ShowDialog();
                if (!string.IsNullOrEmpty(this.txtso_pn.Text.Trim()))
                    this.TabInfo.SelectedIndex = 1;
                if (frmTimPn.isOk)
                {
                    int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                    for (int index = 0; index < count; ++index)
                        StartUpTrans.DsTrans.Tables[1].DefaultView.Delete(0);
                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_pn"] = frmTimPn.dsPn.Tables[0].DefaultView[0]["ngay_ct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_pn"] = (object)frmTimPn.dsPn.Tables[0].DefaultView[0]["so_ct"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pn"] = (object)frmTimPn.dsPn.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh_pn"] = (object)frmTimPn.dsPn.Tables[0].DefaultView[0]["ma_kh"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh_pn"] = (object)frmTimPn.dsPn.Tables[0].DefaultView[0]["ten_kh"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh_pn2"] = (object)frmTimPn.dsPn.Tables[0].DefaultView[0]["ten_kh2"].ToString();
                    DataTable tblPH = frmTimPn.dsPn.Tables[0];
                    foreach (DataRow r in tblPH.Rows)
                    {
                        if (Convert.ToBoolean(r["choose"]))
                        {
                            string upper1 = this.cbMa_nt.Text.ToUpper();
                            string upper2 = frmTimPn.dsPn.Tables[0].DefaultView[0]["ma_nt"].ToString().ToUpper();
                            string stt_rec = r["stt_rec"].ToString();
                            frmTimPn.dsPn.Tables[1].DefaultView.RowFilter = string.Format("{0} = '{1}'", (object)" stt_rec ", (object)stt_rec);
                            for (int index = 0; index < frmTimPn.dsPn.Tables[1].DefaultView.Count; ++index)
                            {
                                DataRow row1 = frmTimPn.dsPn.Tables[1].DefaultView[index].Row;
                                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                                DataTable table = frmTimPn.dsPn.Tables[1].Clone();
                                table.Rows.Add(row1.ItemArray);
                                DataTable dataTable = StartUpTrans.DsTrans.Tables[1].Clone();
                                dataTable.Merge(table, true, MissingSchemaAction.Ignore);
                                if (dataTable.Rows.Count > 0)
                                    row2.ItemArray = dataTable.Rows[0].ItemArray;
                                row2["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                if (upper1.Equals(upper2))
                                {
                                    if (upper1.Equals(StartUpTrans.M_ma_nt0))
                                    {
                                        row2["tien_nt0"] = row1["tien0"];
                                        row2["tien0"] = row1["tien0"];
                                        row2["gia_nt0"] = row1["gia0"];
                                        row2["gia0"] = row1["gia0"];
                                    }
                                    else
                                    {
                                        row2["tien_nt0"] = row1["tien_nt0"];
                                        row2["tien0"] = row1["tien0"];
                                        row2["gia_nt0"] = row1["gia_nt0"];
                                        row2["gia0"] = row1["gia0"];
                                    }
                                }
                                else if (upper1.Equals(StartUpTrans.M_ma_nt0))
                                {
                                    row2["tien_nt0"] = row1["tien0"];
                                    row2["tien0"] = row1["tien0"];
                                    row2["gia_nt0"] = row1["gia0"];
                                    row2["gia0"] = row1["gia0"];
                                }
                                else
                                {
                                    row2["tien_nt0"] = (object)SysFunc.Round(Convert.ToDecimal(row1["tien0"]) / this.txtTy_gia.nValue, StartUpTrans.M_ROUND_NT);
                                    row2["tien0"] = row1["tien0"];
                                    row2["gia_nt0"] = (object)SysFunc.Round(Convert.ToDecimal(row1["gia0"]) / this.txtTy_gia.nValue, StartUpTrans.M_ROUND_GIA_NT);
                                    row2["gia0"] = row1["gia0"];
                                }
                                row2["cp"] = (object)0;
                                row2["cp_nt"] = (object)0;
                                row2["gia"] = row2["gia0"];
                                row2["gia_nt"] = row2["gia_nt0"];
                                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                            }
                        }
                    }

                    this.Sum_ALL();
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                        this.TabInfo.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCp_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpnc.IsInEditMode.Value || !Keyboard.IsKeyDown(Key.Left) && !Keyboard.IsKeyDown(Key.Right))
                return;
            this.NewRowCtGt();
            this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[this.GrdCtgt.Records.Count - 1];
            this.GrdCtgt.ActiveCell = (this.GrdCtgt.ActiveRecord as DataRecord).Cells["ma_ms"];
        }

        private bool GrdCp_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
            return true;
        }

        private void txtHan_ck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void FormMain_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void TabInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.TabInfo.SelectedIndex != 0)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               this.txttong_cp_nt.Focus();
               this.txttong_cp_nt.SelectAll();
           }));
        }
        private void btnViewPC_Click(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.View)
                return;
            try
            {

                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim()))
                    return;
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"].ToString().Trim().Equals("PC1"))
                {
                    SqlCommand sqlcmd1 = new SqlCommand();
                    sqlcmd1.CommandText = "Select count(1) from ph46 WHERE stt_rec = @stt_rec_pc";
                    sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pc", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString();
                    if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                    {
                        if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu chi không tồn tại, có xóa thông tin phiếu chi trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) != MessageBoxResult.Yes)
                            return;
                        SqlCommand sqlcmd2 = new SqlCommand();
                        sqlcmd2.CommandText = "UPDATE "+StartUp.Tablename+" Set stt_rec_pc = '', so_ct_pc = '', ma_ct_pc = '', tao_pc=0 WHERE stt_rec = @stt_rec_pc; ";
                        sqlcmd2.CommandText += "UPDATE cttt20 Set tat_toan = 0, stt_rec_tt = '' WHERE stt_rec = @stt_rec";
                        sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_pc", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                        sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                        this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"] = "";
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"] = "";
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"] = "";
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tao_pc"] = (Object)0;
                    }
                    else
                    {
                        SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "PC1", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString(), this.BindingSasObj.M_StartUp_Path, this.BindingSasObj.M_ProcessName);
                    }

                }
                else
                {
                    SqlCommand sqlcmd1 = new SqlCommand();
                    sqlcmd1.CommandText = "Select count(1) from ph56 WHERE stt_Rec = @stt_rec_pc";
                    sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pc", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString();
                    if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                    {
                        if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu chi không tồn tại, có xóa thông tin phiếu chi trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                        {
                            SqlCommand sqlcmd2 = new SqlCommand();
                            sqlcmd2.CommandText = "UPDATE " + StartUp.Tablename + " Set stt_rec_pc = '', so_ct_pc = '', ma_ct_pc = '', tao_pc=0 WHERE stt_rec = @stt_rec_pc; ";
                            sqlcmd2.CommandText += "UPDATE cttt20 Set tat_toan = 0, stt_rec_tt = '' WHERE stt_rec = @stt_rec";
                            sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_pc", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                            sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                            this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"] = "";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pc"] = "";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"] = "";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tao_pc"] = (Object)0;
                        }
                    }
                    else
                        SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "BN1", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString(), this.BindingSasObj.M_StartUp_Path, this.BindingSasObj.M_ProcessName);
                }

            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
    }
}
