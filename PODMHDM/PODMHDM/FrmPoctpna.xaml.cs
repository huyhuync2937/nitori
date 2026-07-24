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
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;
using static Infragistics.Shared.DynamicResourceString;
using static System.Windows.Forms.AxHost;

namespace PODMHDM
{
    public partial class FrmPoctpna : FormTrans
    {
        public static int iRow = 0;
        public static int OldiRow = 0;
        public string Old_ma_kho = string.Empty;
        private bool txtDiaChiFocusable = true;
        private string ma_hd;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject Voucher_Lan0;
        private DataSet DsVitual;
        private DataSet dsCheckData;
        public string ma_tra_cuu_kh = "";
        public string ma_tra_cuu_vt = "";
        private DataTable dtMa_ncc;
        public static string KeyFilter = "";
        public static FormBrowse obrowseBKCT = (FormBrowse)null;
        private static SqlCommand sqlcmdBKCT;
        private static string strBrowseBKCT = "";
        public FrmPoctpna()
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
                    FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtso_du_kh.Visibility = this.tblso_du_kh.Visibility = Visibility.Hidden;
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
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                }
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
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
            this.GrdLayout10.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout20.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout21.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout22.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdTongChiPhi.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.GrdCp.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtStatus.IsEnabled = false;
        }

        private void V_Dau()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Truoc()
        {
            if (FrmPoctpna.iRow <= 1)
                return;
            --FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Sau()
        {
            if (FrmPoctpna.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Cuoi()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.BtnChonNCC.Focus()));
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)str;
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                row["status"] = StartUpTrans.DmctInfo["ma_post"];
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ma_nt"];
                    row["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ma_qs"].ToString().Trim());
                }
                row["sua_tien"] = (object)0;
                row["ty_giaf"] = !row["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)StartUp.GetRates(row["ma_nt"].ToString().Trim(), Convert.ToDateTime(row["ngay_ct"]).Date) : (object)1;
                row["status"] = StartUpTrans.DmctInfo["ma_post"];
                row["t_cp_nt"] = (object)0;
                row["t_cp"] = (object)0;
                row["t_tien"] = (object)0;
                row["t_tien_nt"] = (object)0;
                row["t_tien0"] = (object)0;
                row["t_tien_nt0"] = (object)0;
                row["t_thue_nt"] = (object)0;
                row["t_thue"] = (object)0;
                row["t_tt_nt"] = (object)0;
                row["t_tt"] = (object)0;
                row["t_so_luong"] = (object)0;
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
            FrmPoctpnaCopy frmPoctpnaCopy = new FrmPoctpnaCopy();
            frmPoctpnaCopy.Closed += new EventHandler(this._formcopy_Closed);
            if (this.M_LAN != "V")
                frmPoctpnaCopy.Title = "Copy";
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
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
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
        }

        private bool KiemTraCoPhatSinh()
        {
            SqlCommand sqlcmd = new SqlCommand("select so_ct from ct00 where ma_hdm = @ma_hd");
            sqlcmd.Parameters.Add("@ma_hd", SqlDbType.Char).Value = (object)this.ma_hd;
            if (StartupBase.SasObj.ExcuteScalar(sqlcmd) != null)
                return true;
            sqlcmd.CommandText = "select so_ct from ct70 where ma_hdm = @ma_hd";
            if (StartupBase.SasObj.ExcuteScalar(sqlcmd) != null)
                return true;
            sqlcmd.CommandText = "select so_ct from cttt30 where ma_hdm = @ma_hd";
            return StartupBase.SasObj.ExcuteScalar(sqlcmd) != null;
        }

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num1 = (int)ExMessageBox.Show(1200, StartupBase.SasObj, "Không có dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                this.ma_hd = this.txtSo_ct.Text.Trim();
                string status = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString();
                string ma_cv = StartupBase.SasObj.UserInfo.Rows[0]["ma_cv"].ToString().Trim();

                switch (ma_cv)
                {
                    case "PIC":
                        if (status != "0" & status != "2" & status != "1")
                        {
                            int num2 = (int)ExMessageBox.Show(1210, StartupBase.SasObj, "Không thể sửa đơn hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                        break;
                }

                if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)))
                {
                    int num2 = (int)ExMessageBox.Show(1205, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                    FormTrans.currActionTask = ActionTask.Edit;
                    this.DsVitual = new DataSet();
                    this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                    this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                    FrmPoctpna.IsInEditMode.Value = true;
                    this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                    this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                }
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

        protected override bool CheckCanDelete()
        {
            this.ma_hd = this.txtSo_ct.Text.Trim();
            if (!this.KiemTraCoPhatSinh())
                return true;
            int num = (int)ExMessageBox.Show(1375, StartupBase.SasObj, "Hợp đồng đã có phát sinh, không được xóa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            return false;
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
            formView.frmBrw.LanguageID = "PODMHDMXemView";
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
            FrmIn frmIn = new FrmIn();
            if (this.M_LAN != "V")
                frmIn.Title = "Print";
            frmIn.ShowDialog();
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
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
                dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView.Count <= 0 ? (object)this.Old_ma_kho : StartUpTrans.DsTrans.Tables[1].DefaultView[StartUpTrans.DsTrans.Tables[1].DefaultView.Count - 1]["ma_kho_i"];
                dataRow["so_luong"] = (object)0;
                dataRow["gia_nt0"] = (object)0;
                dataRow["tien_nt0"] = (object)0;
                dataRow["tien0"] = (object)0;
                dataRow["cp_nt"] = (object)0;
                dataRow["cp"] = (object)0;
                // dataRow["so_luong_duyet"] = (object)0;
                //  dataRow["duyet"] = (object)false;
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
                if (!FrmPoctpna.IsInEditMode.Value || (this.GrdCt.ActiveCell == null || StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) != null))
                    return;
                Decimal num1;
                Decimal num2;
                Decimal num3;
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null || !e.Cell.IsDataChanged)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            ma_tra_cuu_vt = autoCompleteControl1.RowResult["ma_tra_cuu"].ToString();
                            if (!string.IsNullOrEmpty(ma_tra_cuu_vt))
                            {
                                this.txtSo_ct.Text = this.GetLastSoctPO_VT(txtMa_kh.Text.ToString(), txtNgay_ct.dValue, e.Cell.Record.Cells["ma_vt"].Value.ToString());

                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
                            }

                            e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl1.RowResult["ten_vt"];
                            e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl1.RowResult["ten_vt2"];
                            e.Cell.Record.Cells["dvt"].Value = autoCompleteControl1.RowResult["dvt"];

                            DateTime ngayCt = Convert.ToDateTime(txtNgay_ct.Value);
                            int soNgay = Convert.ToInt32(autoCompleteControl1.RowResult["sl_td1"]);

                            e.Cell.Record.Cells["han_gh"].Value = ngayCt.AddDays(soNgay).ToString("dd-MM-yyyy");

                            if (string.IsNullOrEmpty((e.Cell.Record.DataItem as DataRowView)["tk_vt"].ToString()))
                                (e.Cell.Record.DataItem as DataRowView)["tk_vt"] = autoCompleteControl1.RowResult["tk_vt"];
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                            if (autoCompleteControl2 != null)
                            {
                                autoCompleteControl2.SearchInit();
                                if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                                    (e.Cell.Record.DataItem as DataRowView)["tk_vt"] = autoCompleteControl2.RowResult["tk_dl"];
                            }
                            (e.Cell.Record.DataItem as DataRowView)["sua_tk_vt"] = autoCompleteControl1.RowResult["sua_tk_vt"];

                            CellCollection cells = e.Cell.Record.Cells;
                            if (this.txtNgay_ct.dValue != new DateTime())
                            {
                                DataRow dataRow = StartUp.Getdmgia0(e.Editor.Value.ToString(), string.Format("{0:yyyyMMdd}", this.txtNgay_ct.dValue));
                                if (dataRow != null)
                                {
                                    cells["gia_nt0"].Value = !this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? dataRow["gia_nt0"] : dataRow["gia0"];
                                    cells["gia0"].Value = dataRow["gia0"];
                                }
                            }

                            if (autoCompleteControl1.RowResult["vt_ton_kho"].ToString().Equals("0"))
                            {
                                e.Cell.Record.Cells["so_luong"].Value = (object)0;
                                e.Cell.Record.Cells["gia_nt0"].Value = (object)0;
                                e.Cell.Record.Cells["gia0"].Value = (object)0;
                                if (string.IsNullOrEmpty(autoCompleteControl1.RowResult["dvt"].ToString().Trim()))
                                    CellValuePresenter.FromCell(e.Cell.Record.Cells["so_luong"]).Editor.IsReadOnly = true;
                            }
                            else
                                CellValuePresenter.FromCell(e.Cell.Record.Cells["so_luong"]).Editor.IsReadOnly = false;


                            DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                            //CellCollection cells = e.Cell.Record.Cells;
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
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl3 != null)
                        {
                            autoCompleteControl3.SearchInit();
                            if (autoCompleteControl3.RowResult != null && (autoCompleteControl3.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl3.RowResult["tk_dl"].ToString().Trim())))
                                (e.Cell.Record.DataItem as DataRowView)["tk_vt"] = autoCompleteControl3.RowResult["tk_dl"];
                            break;
                        }
                        break;
                    case "so_luong":
                        try
                        {
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["so_luong"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal nValue1 = (e.Editor as NumericTextBox).nValue;

                                DataRow rowQD = GetDVTKL(
                                    e.Cell.Record.Cells["ma_vt"].Value.ToString()
                                );

                                if (rowQD != null && decimal.TryParse(rowQD["hs_qd"]?.ToString(), out decimal hs_qd) && hs_qd != 0)
                                {
                                    var dvt1 = e.Cell.Record.Cells["dvt1"].Value?.ToString().Trim();
                                    var dvt = e.Cell.Record.Cells["dvt"].Value?.ToString().Trim();

                                    bool isSameUnit = dvt1 == dvt;

                                    if (isSameUnit)
                                        nValue1 /= hs_qd;
                                    else
                                        nValue1 *= hs_qd;
                                }
                                Decimal num4 = new Decimal(0);
                                Decimal result1 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                Decimal result3 = new Decimal(0);
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal nValue = (e.Editor as NumericTextBox).nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result3);
                                Decimal.TryParse(e.Cell.Record.Cells["gia_nt0"].Value.ToString(), out result1);
                                Decimal.TryParse(e.Cell.Record.Cells["gia0"].Value.ToString(), out result2);
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    if (result1 * nValue != new Decimal(0))
                                    {
                                        Decimal num5 = SysFunc.Round(result1 * nValue, StartUpTrans.M_ROUND);
                                        Decimal num6 = num5;
                                        e.Cell.Record.Cells["tien_nt0"].Value = (object)num5;
                                        e.Cell.Record.Cells["tien0"].Value = (object)num6;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num5 * result3 / new Decimal(100), StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                    }
                                }
                                else
                                {
                                    if (result1 * nValue != new Decimal(0))
                                    {
                                        Decimal num5 = SysFunc.Round(result1 * nValue, StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["tien_nt0"].Value = (object)num5;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num5 * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    }
                                    if (result2 * nValue != new Decimal(0))
                                    {
                                        Decimal num5 = SysFunc.Round(result2 * nValue, StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["tien0"].Value = (object)num5;
                                        e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num5 * result3 / new Decimal(100), StartUpTrans.M_ROUND);
                                    }
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                            break;
                        }
                    case "gia_nt0":
                        if (e.Editor.Value == DBNull.Value)
                            e.Cell.Record.Cells["gia_nt0"].Value = (object)0;
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal result1 = new Decimal(0);
                            Decimal num4 = new Decimal(0);
                            Decimal num5 = new Decimal(0);
                            num2 = new Decimal(0);
                            num3 = new Decimal(0);
                            Decimal result2 = new Decimal(0);
                            Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result1);
                            Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                            Decimal nValue2 = this.txtTy_gia.nValue;
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                bool? isChecked = this.ChkSuaTien.IsChecked;
                                if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0 && nValue1 * result1 != new Decimal(0))
                                {
                                    Decimal num6 = SysFunc.Round(result1 * nValue1, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien_nt0"].Value = (object)num6;
                                    e.Cell.Record.Cells["tien0"].Value = (object)num6;
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num6 * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                }
                                e.Cell.Record.Cells["gia0"].Value = (object)nValue1;
                            }
                            else
                            {
                                if (nValue1 * result1 != new Decimal(0))
                                {
                                    bool? isChecked = this.ChkSuaTien.IsChecked;
                                    if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                                    {
                                        num5 = SysFunc.Round(result1 * nValue1, StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["tien_nt0"].Value = (object)num5;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num5 * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    }
                                }
                                if (nValue1 * nValue2 != new Decimal(0))
                                    e.Cell.Record.Cells["gia0"].Value = (object)SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND_GIA);
                                if (num5 * nValue2 != new Decimal(0))
                                {
                                    Decimal num6 = SysFunc.Round(num5 * nValue2, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien0"].Value = (object)num6;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num6 * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                            }
                            this.Sum_ALL();
                            break;
                        }
                        break;
                    case "tien_nt0":
                        if (e.Editor.Value == DBNull.Value)
                            e.Cell.Record.Cells["tien_nt0"].Value = (object)0;
                        if (e.Cell.IsDataChanged)
                        {
                            num3 = new Decimal(0);
                            num1 = new Decimal(0);
                            num2 = new Decimal(0);
                            Decimal result = new Decimal(0);
                            Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                            Decimal nValue2 = this.txtTy_gia.nValue;
                            Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result);
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt0"].Value, new Decimal(0));
                            if (num5 == new Decimal(0))
                            {
                                if (num4 != new Decimal(0))
                                    num5 = SysFunc.Round(nValue1 / num4, StartUpTrans.M_ROUND_GIA);
                                e.Cell.Record.Cells["gia_nt0"].Value = (object)num5;
                            }
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien0"].Value = (object)nValue1;
                                e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(nValue1 * result / new Decimal(100), StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                            }
                            else
                            {
                                if (nValue1 * nValue2 != new Decimal(0))
                                {
                                    Decimal num6 = SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien0"].Value = (object)num6;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num6 * result / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                                e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(nValue1 * result / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            }
                            this.Sum_ALL();
                            break;
                        }
                        break;
                    case "gia0":
                        if (e.Editor.Value == DBNull.Value)
                            e.Cell.Record.Cells["gia0"].Value = (object)0;
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal result1 = new Decimal(0);
                            Decimal num4 = new Decimal(0);
                            num2 = new Decimal(0);
                            Decimal result2 = new Decimal(0);
                            Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                            Decimal nValue = (e.Editor as NumericTextBox).nValue;
                            Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result1);
                            if (nValue * result1 != new Decimal(0))
                            {
                                Decimal num5 = SysFunc.Round(nValue * result1, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["tien0"].Value = (object)num5;
                                e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num5 * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                            }
                            this.Sum_ALL();
                            break;
                        }
                        break;
                    case "tien0":
                        if (e.Cell.IsDataChanged)
                        {
                            num2 = new Decimal(0);
                            Decimal result = new Decimal(0);
                            Decimal nValue = (e.Editor as NumericTextBox).nValue;
                            Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result);
                            e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(nValue * result / new Decimal(100), StartUpTrans.M_ROUND);
                            this.Sum_ALL();
                            break;
                        }
                        break;
                    case "ma_thue":
                        AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl4.IsDataChanged)
                        {
                            Decimal num4 = new Decimal(0);
                            Decimal num5 = new Decimal(0);
                            Decimal num6 = new Decimal(0);
                            if (autoCompleteControl4.RowResult != null)
                            {
                                e.Cell.Record.Cells["thue_suat"].Value = autoCompleteControl4.RowResult["thue_suat"];
                                Decimal num7 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt0"].Value, new Decimal(0));
                                Decimal num8 = this.ParseDecimal(e.Cell.Record.Cells["tien0"].Value, new Decimal(0));
                                Decimal num9 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num7 * num9 / new Decimal(100), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                }
                                else
                                {
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round(num7 * num9 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num8 * num9 / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                            }
                            break;
                        }
                        break;
                    case "thue_nt":
                        if (e.Editor.Value == DBNull.Value)
                            e.Cell.Record.Cells["thue_nt"].Value = (object)0;
                        if (e.Cell.IsDataChanged)
                        {
                            Decimal num4 = new Decimal(0);
                            num3 = new Decimal(0);
                            Decimal nValue1 = this.txtTy_gia.nValue;
                            Decimal nValue2 = (e.Editor as NumericTextBox).nValue;
                            e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(nValue2 * nValue1, StartUpTrans.M_ROUND);
                            this.Sum_ALL();
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

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.txtHan_ck.IsFocus = true;
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (FrmPoctpna.IsInEditMode.Value)
            {
                Key key = e.Key;
                if (key != Key.F4)
                {
                    if (key == Key.F8)
                    {
                        if (ExMessageBox.Show(1210, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.No)
                        {
                            DataRecord dataRecord = this.GrdCt.ActiveRecord as DataRecord;
                            if (dataRecord != null)
                            {
                                Cell activeCell = this.GrdCt.ActiveCell;
                                if (dataRecord.Index == 0)
                                {
                                    if (this.GrdCt.Records.Count == 1)
                                    {
                                        this.GrdCt_AddNewRecord(null, null);
                                    }
                                }
                                else if (dataRecord.Index == this.GrdCt.Records.Count - 1)
                                {
                                    int num = dataRecord.Index - 1;
                                }
                                int num2 = (this.GrdCt.ActiveCell == null) ? 0 : this.GrdCt.ActiveCell.Field.Index;
                                this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                                if (num2 >= 0)
                                {
                                    StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[dataRecord.Index].Row);
                                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                                    if (this.GrdCt.Records.Count > 0)
                                    {
                                        this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                                    }
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt0"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(cp_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(cp)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
                                    this.Sum_ALL();
                                }
                            }
                        }
                    }
                }
                else
                {
                    DataRecord dataRecord = this.GrdCt.ActiveRecord as DataRecord;
                    if (dataRecord != null && dataRecord.Cells["ma_vt"].Value != null && !(dataRecord.Cells["ma_vt"].Value.ToString() == ""))
                    {
                        switch (Keyboard.Modifiers)
                        {
                            case ModifierKeys.None:
                                this.NewRowCt();
                                this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                                this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"];
                                break;
                            case ModifierKeys.Control:
                                this.InsertRecord(delegate
                                {
                                    this.NewRowCt();
                                }, this.GrdCt, "ma_vt");
                                break;
                        }
                    }
                }
            }
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt"].Value == null || activeRecord.Cells["ma_vt"].Value.ToString() == ""))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            this.GrdCt.ActiveCell = (this.GrdCt.Records[this.GrdCt.Records.Count - 1] as DataRecord).Cells["ma_vt"];
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
                bool flag = false;
                if (!this.IsSequenceSave)
                {
                    Decimal result1 = new Decimal(0);
                    Decimal result2 = new Decimal(0);
                    Decimal result3 = new Decimal(0);
                    Decimal result4 = new Decimal(0);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(cp)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result3);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result2);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result4);
                    this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(1215, StartupBase.SasObj, "Chưa vào mã khách hàng!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag = true;
                    }
                    else if (string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                    {
                        int num = (int)ExMessageBox.Show(1225, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(6101, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            this.txtMa_qs.IsFocus = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(275, StartupBase.SasObj, "Chưa vào số hóa đơn!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            this.txtSo_ct.Focus();
                            flag = true;
                        }
                        else if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(1230, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCt.Focus();
                            flag = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(1240, StartupBase.SasObj, "Chưa vào số chứng từ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            flag = true;
                        }
                        else if (result1 != result2 || result3 != result4)
                        {
                            int num2 = (int)ExMessageBox.Show(1255, StartupBase.SasObj, "Tổng chi phí khác với chi phí tổng cộng của các vật tư!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
                            this.GrdCp.ActiveCell = (this.GrdCp.Records[0] as DataRecord).Cells["cp_nt"];
                            this.GrdCp.Focus();
                            flag = true;
                        }
                        else if (!StartUpTrans.M_MST_CHECK.Equals("0") && (!SysFunc.CheckSumMaSoThue(this.txtMaSoThue.Text.Trim()) && !string.IsNullOrEmpty(this.txtMaSoThue.Text.Trim())))
                        {
                            int num2 = (int)ExMessageBox.Show(1260, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            if (StartUpTrans.M_MST_CHECK.Equals("2"))
                                return;
                        }
                    }
                    if (!flag)
                    {
                        this.Sum_ALL();
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                        {
                            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                            {
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString()))
                                {
                                    int num = (int)ExMessageBox.Show(1265, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_vt"].ToString()))
                                {
                                    int num = (int)ExMessageBox.Show(1275, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_vt"];
                                    this.GrdCt.Focus();
                                    return;
                                }
                            }
                        }
                    }
                }
                if (!flag)
                {
                    if (!this.IsSequenceSave)
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
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
                            Decimal num9 = new Decimal(0);
                            Decimal num10 = new Decimal(0);
                            Decimal num11 = new Decimal(0);
                            Decimal num12 = new Decimal(0);
                            Decimal num13 = new Decimal(0);
                            Decimal num14 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"], new Decimal(0));
                            Decimal num15;
                            Decimal num16;
                            if (num14 > new Decimal(0))
                            {
                                num15 = SysFunc.Round(num6 / num14, StartUpTrans.M_ROUND_GIA_NT);
                                num16 = SysFunc.Round(num8 / num14, StartUpTrans.M_ROUND_GIA);
                            }
                            else
                            {
                                Decimal num17;
                                num12 = num17 = new Decimal(0);
                                num13 = num17;
                                num16 = num17;
                                num15 = num17;
                            }
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia"] = (object)num16;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"] = (object)num15;
                        }
                        Decimal num = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                        bool? isChecked = this.ChkSuaTien.IsChecked;
                        if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && num != new Decimal(0))
                            this.CanBangTien();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt0"], new Decimal(0)) + this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"], new Decimal(0)));
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"], new Decimal(0)) * this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0)));
                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ky"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_hd"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_hdm"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_ct"] = StartUpTrans.DmctInfo["ct_nxt"];
                    if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("search"))
                    {
                        DataTable table = StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable();
                        SysFunc.SetStrSearch(StartupBase.SasObj, "dmhdm", ref table);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["search"] = (object)table.Rows[0]["search"].ToString().Trim();
                    }
                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    //if (!this.IsSequenceSave)
                    //    LocalTable1.Rows[0]["status"] = (object)0;
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                        {
                            dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;
                            dataRowView.Row["ma_hdm_i"] = (object)this.txtSo_ct.Text.Trim();
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
                        LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num18 = (int)ExMessageBox.Show(1285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                                    if (ExMessageBox.Show(1290, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                    {
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag = true;
                                                        break;
                                                    }
                                                    break;
                                                }
                                                int num1 = (int)ExMessageBox.Show(1295, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.txtSo_ct.SelectAll();
                                                this.txtSo_ct.Focus();
                                                flag = true;
                                                break;
                                            case "PH02":
                                                int num2 = (int)ExMessageBox.Show(1300, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = true;
                                                this.txtMa_nx.IsFocus = true;
                                                break;
                                            case "CT01":
                                                int int16 = (int)Convert.ToInt16(dataRowView[1]);
                                                int num3 = (int)ExMessageBox.Show(1305, StartupBase.SasObj, "Tk vật tư là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = true;
                                                this.tiHT.Focus();
                                                this.GrdCt.ActiveCell = (this.GrdCt.Records[int16] as DataRecord).Cells["tk_vt"];
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
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[PODMHDM-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }


        public string GetLastSoctPO(string ma_kh, DateTime ngay_ct)
        {
            string str = "";
            string ngay = txtNgay_ct.dValue.ToString("yyMMdd");

            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[GetSoPO] @ma_kh , @ngay_ct");
                sqlcmd.Parameters.Add("@ma_kh", SqlDbType.VarChar).Value = (object)ma_kh;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.SmallDateTime).Value = (object)ngay_ct;
                DataTable dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
                if (dataTable.Rows.Count == 0)
                {
                    str = ma_tra_cuu_kh + ngay;
                }
                else
                {
                    string lastSoct = dataTable.Rows[0][0].ToString();

                    int pos = lastSoct.LastIndexOf('-');

                    if (pos < 0)
                    {
                        // Chưa có hậu tố -xx
                        str = lastSoct + "-01";
                    }
                    else
                    {
                        string soCuoi = lastSoct.Substring(pos + 1);

                        if (int.TryParse(soCuoi, out int so))
                        {
                            str = lastSoct.Substring(0, pos) + "-" + (so + 1).ToString("00");
                        }
                        else
                        {
                            str = lastSoct + "-01";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return str;
        }

        public string GetLastSoctPO_VT(string ma_kh, DateTime ngay_ct, string ma_vt)
        {
            string str = "";
            string ngay = txtNgay_ct.dValue.ToString("yyMMdd");

            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[GetSoPO_VT] @ma_kh , @ngay_ct, @ma_vt");
                sqlcmd.Parameters.Add("@ma_kh", SqlDbType.VarChar).Value = (object)ma_kh;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.SmallDateTime).Value = (object)ngay_ct;
                sqlcmd.Parameters.Add("@ma_vt", SqlDbType.VarChar).Value = (object)ma_vt;

                DataTable dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
                if (dataTable.Rows.Count == 0)
                {
                    str = ma_tra_cuu_kh + ngay + ma_tra_cuu_vt;
                }
                else
                {
                    string lastSoct = dataTable.Rows[0][0].ToString();

                    int pos = lastSoct.LastIndexOf('-');

                    if (pos < 0)
                    {
                        // Chưa có hậu tố -xx
                        str = lastSoct + "-01";
                    }
                    else
                    {
                        string soCuoi = lastSoct.Substring(pos + 1);

                        if (int.TryParse(soCuoi, out int so))
                        {
                            str = lastSoct.Substring(0, pos) + "-" + (so + 1).ToString("00");
                        }
                        else
                        {
                            str = lastSoct + "-01";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return str;
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
            if (!FrmPoctpna.IsInEditMode.Value || this.txtMa_kh.RowResult == null)
                return;
            if (this.txtMa_kh.RowResult["dia_chi"].ToString().Trim().Equals(""))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_kh.RowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = (object)this.txtMa_kh.RowResult["ten_kh"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = (object)this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = (object)this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["gc_td1"] = (object)this.txtMa_kh.RowResult["gc_td1"].ToString().Trim();
            //StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_td1"] = (object)this.txtMa_kh.RowResult["sl_td1"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"] = (object)this.txtMa_kh.RowResult["gc_td1"].ToString().Trim();
            ma_tra_cuu_kh = this.txtMa_kh.RowResult["ma_tra_cuu"].ToString().Trim();


            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) ? (object)this.txtMa_kh.RowResult["tk"].ToString().Trim() : (object)this.txtMa_nx.Text.Trim();
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            DataRow rowResult = this.txtMa_kh.RowResult;
            if (string.IsNullOrEmpty(dataRowView["ma_thck"].ToString().Trim()))
            {
                dataRowView["ma_thck"] = rowResult["ma_thck"];
                dataRowView["ten_thck"] = rowResult["ten_thck"];
                dataRowView["ten_thck2"] = rowResult["ten_thck2"];
            }
            if (this.ParseInt((object)dataRowView["han_tt"].ToString(), 0) == 0)
                dataRowView["han_tt"] = (object)this.ParseInt(rowResult["han_tt"], 0);
        }

        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            this.txtDia_chi.IsTabStop = false;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        //private void txtDk_tt_GotFocus(object sender, RoutedEventArgs e)
        //{
        //    if (this.txtDiaChiFocusable)
        //        return;
        //    this.txtDk_tt.IsTabStop = false;
        //    SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        //}

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_nx.RowResult == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_tk"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_tk2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
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
                this.txtTy_gia.Value = this.cbMa_nt.RowResult["ma_nt"].ToString() == StartUpTrans.M_ma_nt0 ? (object)1 : this.txtTy_gia.Value;
            if (string.IsNullOrEmpty(this.txtTy_gia.Text.ToString()))
                this.txtTy_gia.Value = (object)0;
            try
            {
                if (FormTrans.currActionTask == ActionTask.Delete || !FrmPoctpna.IsInEditMode.Value || (this.txtTy_gia.Value == null || this.txtTy_gia.Value == DBNull.Value || !(this.txtTy_gia.nValue != new Decimal(0))))
                    return;
                Decimal num1 = new Decimal(0);
                Decimal num2 = new Decimal(0);
                Decimal num3 = new Decimal(0);
                Decimal num4 = new Decimal(0);
                Decimal num5 = new Decimal(0);
                Decimal num6 = new Decimal(0);
                Decimal num7 = new Decimal(0);
                Decimal num8 = new Decimal(0);
                Decimal num9 = new Decimal(0);
                Decimal nValue = this.txtTy_gia.nValue;
                num6 = this.txtT_Tien_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_Tien_nt.Value);
                num7 = this.txttong_cp_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txttong_cp_nt.Value);
                if (this.GrdCt.Records.Count > 0 && (this.GrdCt.DataSource as DataView).Table.DefaultView[0]["ma_vt"] != DBNull.Value)
                {
                    for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                    {
                        if ((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt0"].Value != DBNull.Value)
                        {
                            Decimal num10 = (this.GrdCt.DataSource as DataView)[index]["so_luong"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["so_luong"].Value);
                            Decimal num11 = (this.GrdCt.DataSource as DataView)[index]["gia_nt0"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["gia_nt0"].Value);
                            if (num10 * num11 != new Decimal(0))
                            {
                                num2 = SysFunc.Round(num10 * num11, StartUpTrans.M_ROUND_NT);
                                (this.GrdCt.DataSource as DataView)[index]["tien_nt0"] = (object)num2;
                            }
                            if (nValue * num11 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["gia0"] = (object)SysFunc.Round(nValue * num11, StartUpTrans.M_ROUND_GIA);
                            if (nValue * num2 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["tien0"] = (object)SysFunc.Round(nValue * num2, StartUpTrans.M_ROUND);
                            Decimal num12 = new Decimal(0);
                            Decimal num13 = new Decimal(0);
                            Decimal num14 = new Decimal(0);
                            Decimal num15 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["tien_nt0"], new Decimal(0));
                            Decimal num16 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["tien0"], new Decimal(0));
                            Decimal num17 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["thue_suat"], new Decimal(0));
                            if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                (this.GrdCt.DataSource as DataView)[index]["thue_nt"] = (object)SysFunc.Round(num15 * num17 / new Decimal(100), StartUpTrans.M_ROUND);
                                (this.GrdCt.DataSource as DataView)[index]["thue"] = (this.GrdCt.DataSource as DataView)[index]["thue_nt"];
                            }
                            else
                            {
                                (this.GrdCt.DataSource as DataView)[index]["thue_nt"] = (object)SysFunc.Round(num15 * num17 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                (this.GrdCt.DataSource as DataView)[index]["thue"] = (object)SysFunc.Round(num16 * num17 / new Decimal(100), StartUpTrans.M_ROUND);
                            }
                        }
                    }
                    Decimal num18 = this.txttong_cp_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txttong_cp_nt.Value.ToString());
                    if (this.GrdCp.Records.Count > 0)
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] = (object)SysFunc.Round(num18 * nValue, StartUpTrans.M_ROUND);
                        this.PhanBo();
                    }
                    this.Sum_ALL();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                return;
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
            {
                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim()) || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"].ToString().Trim()))
                {
                    //this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                    this.txtSo_ct.Text = this.GetLastSoctPO(txtMa_kh.Text.ToString(), txtNgay_ct.dValue);

                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;

                }
                else
                    this.txtSo_ct.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim();
            }
            else
            {
                this.txtSo_ct.Text = this.GetLastSoctPO(txtMa_kh.Text.ToString(), txtNgay_ct.dValue);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
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

        private void Sum_ALL()
        {
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal num5 = new Decimal(0);
            Decimal num6 = new Decimal(0);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            Decimal num7;
            Decimal num8;
            Decimal num9;
            Decimal num10;
            Decimal num11;
            Decimal num12;
            if (this.cbMa_nt.Text.Equals(StartUpTrans.M_ma_nt0))
            {
                num7 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num8 = num7;
                num9 = SysFunc.Round(this.ParseDecimal((object)this.txttong_cp_nt.nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num10 = num9;
                num11 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num12 = num11;
            }
            else
            {
                num8 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num7 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                Decimal nValue = this.txttong_cp.nValue;
                num10 = SysFunc.Round(this.ParseDecimal((object)nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                nValue = this.txttong_cp_nt.nValue;
                num9 = SysFunc.Round(this.ParseDecimal((object)nValue.ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num12 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num11 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"] = (object)num8;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt0"] = (object)num7;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"] = (object)num10;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"] = (object)num9;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)num12;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num11;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(num8 + num10 + num12);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(num7 + num9 + num11);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (ma_nt == StartUpTrans.M_ma_nt0)
                this.txtTy_gia.IsReadOnly = true;
            else if (FrmPoctpna.IsInEditMode.Value)
                this.txtTy_gia.IsReadOnly = false;
            else
                this.txtTy_gia.IsReadOnly = true;
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
                    this.txttong_cp.IsReadOnly = true;
                    this.txttong_cp.IsTabStop = false;
                    break;
                case 1:
                    if (FrmPoctpna.IsInEditMode.Value)
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
        }

        private void PhanBo()
        {
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal num1 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal num5 = new Decimal(0);
            Decimal result5 = new Decimal(0);
            string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", "stt_rec= '" + str + "'").ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", "stt_rec= '" + str + "'").ToString(), out result4);
            Decimal result6 = new Decimal(0);
            Decimal result7 = new Decimal(0);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"].ToString(), out result6);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"].ToString(), out result7);
                Decimal num6 = !(this.cbMa_nt.Text != StartUpTrans.M_ma_nt0) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result7 / result4 * result2, StartUpTrans.M_ROUND_NT)) : (!(result7 == new Decimal(0)) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result7 / result4 * result2, StartUpTrans.M_ROUND_NT)) : (result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result6 / result3 * result2, StartUpTrans.M_ROUND_NT)));
                Decimal num7 = !(num6 != new Decimal(0)) ? (!(result4 != new Decimal(0)) ? (!(result3 != new Decimal(0)) ? new Decimal(0) : SysFunc.Round(result6 / result3 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(result7 / result4 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(num6 * result5, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"] = (object)num7;
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp_nt"] = (object)num6;
                num2 += num7;
                num3 += num6;
            }
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp"].ToString()) + (result1 - num2));
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["cp_nt"].ToString()) + (result2 - num3));
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
            Decimal result3 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result3);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(result1 + result2 + result3);
        }

        private void btnPhanBo_Click(object sender, RoutedEventArgs e)
        {
            this.PhanBo();
            int num = (int)ExMessageBox.Show(1310, StartupBase.SasObj, "Đã thực hiện xong phân bổ chi phí!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private bool GrdCp_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D3);
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            return false;
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
                        if (!e.Cell.IsDataChanged)
                            break;
                        Decimal num1 = new Decimal(0);
                        Decimal num2 = new Decimal(0);
                        Decimal nValue1 = this.txtTy_gia.nValue;
                        Decimal nValue2 = (e.Editor as NumericTextBox).nValue;
                        if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                            e.Cell.Record.Cells["cp"].Value = e.Cell.Record.Cells["cp_nt"].Value;
                        else if (nValue2 * nValue1 != new Decimal(0))
                            e.Cell.Record.Cells["cp"].Value = (object)SysFunc.Round(nValue2 * nValue1, StartUpTrans.M_ROUND);
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
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal result5 = new Decimal(1);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt0"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien0)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"] = (object)SysFunc.Round(result1 * result5, StartUpTrans.M_ROUND);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt0"], new Decimal(0)) != new Decimal(0))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"], new Decimal(0)) + (this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"], new Decimal(0)) - result4));
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien0"], new Decimal(0)) + this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["cp"], new Decimal(0)));
                    break;
                }
            }
            Decimal result6 = new Decimal(0);
            Decimal result7 = new Decimal(0);
            Decimal result8 = new Decimal(0);
            Decimal result9 = new Decimal(0);
            Decimal result10 = new Decimal(0);
            Decimal result11 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt0"].ToString(), out result6);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien0"].ToString(), out result7);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp_nt"].ToString(), out result8);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_cp"].ToString(), out result9);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result10);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result11);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(result6 + result8 + result10);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(result7 + result9 + result11);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void txttong_cp_nt_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value != DBNull.Value)
            {
                if (!(string.IsNullOrEmpty(this.txtMa_qs.Text)))
                {
                    this.txtSo_ct.Text = this.GetLastSoctPO(txtMa_kh.Text.ToString(), txtNgay_ct.dValue);

                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
                }
                return;
            }
            this.txtNgay_ct.Value = (object)DateTime.Now;

        }

        private void txtSo_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Edit || !(this.txtSo_ct.Text.Trim() != this.ma_hd) || !this.KiemTraCoPhatSinh())
                return;
            int num = (int)ExMessageBox.Show(1370, StartupBase.SasObj, "Hợp đồng đã có phát sinh, không được sửa số hợp đồng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtSo_ct.Text = this.ma_hd;
        }

        private void txtHan_tt_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtHan_tt.SelectAll();
        }

        private void txtHan_ck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtHan_ck.RowResult != null)
            {
                if (this.txtHan_tt.Value == DBNull.Value || this.txtHan_tt.nValue == new Decimal(0))
                    this.txtHan_tt.Value = this.txtHan_ck.RowResult["han_tt"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck"] = this.txtHan_ck.RowResult["ten_thck"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck2"] = this.txtHan_ck.RowResult["ten_thck2"];
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck"] = (object)"";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck2"] = (object)"";
            }
        }

        private void txtHan_tt_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.txtHan_tt.IsFocusWithin)
                   return;
               (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
           }));
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
        private void BtnChonNCC_Click(object sender, RoutedEventArgs e)
        {
            FrmLoc Locdonhang = new FrmLoc();
            Locdonhang.ShowDialog();
            if (StartUp.isOk)
            {
                if (StartUp.HDBData.Tables[1].DefaultView.Count > 0)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_yeucau"] = StartUp.HDBData.Tables[1].DefaultView[0]["ma_hd"].ToString().Trim();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"] = StartUp.HDBData.Tables[0].DefaultView[0]["ncc"].ToString().Trim();
                    StartUpTrans.DsTrans.Tables[1].Rows.Clear();
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
                            dataRow["ma_vt"] = r["ma_vt"];
                            dataRow["ten_vt"] = r["ten_vt"];
                            dataRow["he_so1"] = 0;
                            dataRow["dvt1"] = r["dvt1"];
                            dataRow["so_luong"] = r["so_luong"];
                            dataRow["gia0"] = r["gia2"];
                            dataRow["gia_nt0"] = r["gia_nt2"];
                            dataRow["tien0"] = r["tien2"];
                            dataRow["tien_nt0"] = r["tien_nt2"];
                            dataRow["tk_vt"] = r["tk_vt"];
                            dataRow["so_yeucau"] = r["ma_hd"];
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
            }
        }

        private void btnChoose_Click(object sender, RoutedEventArgs e)
        {
            FrmLocVT LocVT = new FrmLocVT();
            LocVT.ShowDialog();
            if (StartUp.isOk)
            {
                if (StartUp.dataRowArray.Count() > 0)
                {
               
                    StartUpTrans.DsTrans.Tables[1].Rows.Clear();
                    foreach (DataRow r in StartUp.dataRowArray)
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
                            //dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num2);
                            //dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            //dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                            dataRow["ma_vt"] = r["ma_vt"];
                            dataRow["ten_vt"] = r["ten_vt"];
                            //dataRow["he_so1"] = 0;
                            //dataRow["dvt1"] = r["dvt1"];
                            dataRow["so_luong"] = r["so_luong"];
                            //dataRow["gia0"] = r["gia2"];
                            //dataRow["gia_nt0"] = r["gia_nt2"];
                            //dataRow["tien0"] = r["tien2"];
                            //dataRow["tien_nt0"] = r["tien_nt2"];
                            //dataRow["tk_vt"] = r["tk_vt"];
                            //dataRow["so_yeucau"] = r["ma_hd"];
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
            }
        }

        private void btnSendEmail_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_kh.Text))
                return;
            try
            {
                string msg = "";
                string voucherIds = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                if (string.IsNullOrEmpty(voucherIds))
                    return;


                SqlCommand sql_link = new SqlCommand("SELECT top 1  ten_td FROM dmtd5 ");
                string link = "";
                DataTable table_link = StartupBase.SasObj.ExcuteReader(sql_link).Tables[0];
                foreach (DataRow row in table_link.Rows)
                {
                    link = row["ten_td"].ToString().Trim();
                }

                if (string.IsNullOrEmpty(link))
                {
                    MessageBox.Show("Hóa đơn không phải hóa đơn tạo từ PMKT hoặc Không tồn tại file PDF hóa đơn điện tử không gửi Email được", "Thông báo");
                }
                else
                {
                    DataTable dt = null;
                    using (SqlCommand sqlcmd = new SqlCommand())
                    {
                        sqlcmd.CommandText = string.Format("select top 1 * from dmemail where user_id = {0}", StartUpTrans.M_User_Id);
                        dt = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                    }
                    string EmailFrom = string.Empty;
                    string EmailPassFrom = string.Empty;
                    //string EmailTo = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["e_mail"].ToString().Trim();

                    List<string> listEmail = new List<string>();

                    FrmPoctpna.KeyFilter = string.IsNullOrEmpty(this.txtMa_kh.Text) ? "1=1" : "ma_kh = '" + this.txtMa_kh.Text.Trim().Replace("'", "''") + "'";

                    //SqlCommand sql = new SqlCommand("SELECT e_mail FROM dmkh WHERE " + (string.IsNullOrEmpty(FrmPoctpna.KeyFilter) ? "" : " and " + FrmPoctpna.KeyFilter));

                    SqlCommand sql = new SqlCommand("SELECT tk_portal FROM dmkh WHERE " + FrmPoctpna.KeyFilter);

                    DataTable table = StartupBase.SasObj.ExcuteReader(sql).Tables[0];

                    if(table == null || table.Rows.Count == 0)
                    {
                        MessageBox.Show("Khách hàng không có email để gửi", "Thông báo");
                        return;
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        string EmailTo = row["tk_portal"].ToString().Trim();

                        if (!string.IsNullOrEmpty(EmailTo))
                        {
                            if (dt != null && dt.Rows.Count > 0)
                            {
                                EmailFrom = dt.Rows[0]["e_mail"].ToString().Trim();
                                //EmailPassFrom = SasUtilities.ClsEmail.DeCrypt(dt.Rows[0]["pass"].ToString().Trim(), "0123456789").ToString().Trim();
                                EmailPassFrom = dt.Rows[0]["pass"].ToString().Trim();

                            }
                            EmailFrom = EmailFrom ?? "";
                            EmailPassFrom = EmailPassFrom ?? "";
                            FrmGuiEmail frm = new FrmGuiEmail();
                            frm.txtFEmail.Text = EmailFrom.ToString().Trim();
                            frm.txtTEmail.Text = EmailTo.ToString().Trim();
                            frm.ShowDialog();
                            if (!frm.isError)
                            {
                                //Thực hiện gửi Email
                                FrmWaiting frmWaiting = new FrmWaiting(600);
                                frmWaiting.Set(200);
                                frmWaiting.Show();
                                msg = SasUtilities.ClsEmail.sendmail(EmailFrom, EmailPassFrom, frm.txtTEmail.Text.ToString().Trim(), "", "", frm.txtTieude.Text.ToString().Trim(), frm.txtNoidung1.Text.ToString().Trim().Trim(), link);
                                frmWaiting.Set(600);
                                frmWaiting.Close();

                                if (msg.Length > 0)
                                {
                                    MessageBox.Show(msg, "Thong bao");
                                }
                                else
                                {
                                    MessageBox.Show("Gửi Email thành công", "Thong bao");
                                }

                            }
                        }
                        else
                        {
                            MessageBox.Show("Khách hàng không có email để gửi", "Thông báo");
                            return;
                        }
                    }



                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static DataRow GetDVTKL(string ma_vt)
        {
            string sql = "SELECT * FROM v_dmvtdvt WHERE ma_vt = @ma_vt and hs_qd != 1";

            using (SqlCommand sqlCommand = new SqlCommand(sql))
            {
                sqlCommand.Parameters.Add("@ma_vt", SqlDbType.Char, 16).Value = ma_vt;

                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlCommand);
                return (dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                    ? dataSet.Tables[0].Rows[0]
                    : null;
            }
        }
        private void BtnDvt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CallGridReportBKCT(true);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        public static void CallGridReportBKCT(bool isFirstLoad)
        {
            try
            {
                string stt_rec = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim();
                string stt_rec0 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec0"].ToString().Trim();


                FrmPoctpna.sqlcmdBKCT = new SqlCommand();
                FrmPoctpna.sqlcmdBKCT.CommandText = "Gethdmgh";
                FrmPoctpna.sqlcmdBKCT.CommandType = CommandType.StoredProcedure;
                FrmPoctpna.sqlcmdBKCT.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = stt_rec.ToString().Trim();
                FrmPoctpna.sqlcmdBKCT.Parameters.Add("@stt_rec0", SqlDbType.VarChar).Value = stt_rec0.ToString().Trim();


                DataTable dataTable = StartupBase.SasObj.ExcuteReader(FrmPoctpna.sqlcmdBKCT).Tables[0].Copy();
                dataTable.TableName = "tbDetail";

                strBrowseBKCT = "chon:80:h=Chọn;id:80:h=Id;ngay_giao:80:h=Ngày giao;dvt:80:h=DVT;packing:80:h=Packing;he_so:80:h=Hệ số;so_luong:80:h=Số lượng";


                FrmPoctpna.obrowseBKCT = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, strBrowseBKCT);
                FrmPoctpna.obrowseBKCT.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(FrmPoctpna.FormBrowse_Esc);
                //FrmPoctpna.obrowseBKCT.frmBrw.PreviewKeyDown += new KeyEventHandler(FrmPoctpna.FrmBrwBKCT_PreviewKeyDown);
                FrmPoctpna.obrowseBKCT.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(FrmPoctpna.obrowseBKCT_CTRL_R);
                FrmPoctpna.obrowseBKCT.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                FrmPoctpna.obrowseBKCT.frmBrw.Title = SysFunc.Cat_Dau((StartupBase.M_LAN.Equals("V") ? "Chi tiết lịch giao hàng" : "Delivery schedule details"));
                object name = FrmPoctpna.obrowseBKCT.frmBrw.ToolBar.FindName("tbReport");
                if (name != null)
                {
                    ToolBar toolBar = name as ToolBar;
                    for (int i = toolBar.Items.Count - 1; i > 0; i--)
                    {
                        if ((toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnRefresh" && (toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnExport")
                        {
                            toolBar.Items.Remove((toolBar.Items[i] as SasControls.ToolBarButton));
                        }
                    }
                    SasControls.ToolBarButton toolBarButton3 = new SasControls.ToolBarButton();
                    toolBarButton3.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton3.Name = "btnMoi";
                    toolBarButton3.Text = "Thêm mới";
                    toolBarButton3.ToolTip = "F4";
                    toolBarButton3.ImagePath = "Images\\UpdateSearch.png";
                    toolBarButton3.Click += new RoutedEventHandler(ToolBarButtonF4_Click);
                    toolBar.Items.Insert(1, toolBarButton3);

                    SasControls.ToolBarButton toolBarButton1 = new SasControls.ToolBarButton();
                    toolBarButton1.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton1.Name = "btnXoa";
                    toolBarButton1.Text = "Xóa";
                    toolBarButton1.ToolTip = "F5";
                    toolBarButton1.ImagePath = "Images\\AddNew.png";
                    toolBarButton1.Click += new RoutedEventHandler(ToolBarButtonF5_Click);
                    toolBar.Items.Insert(2, toolBarButton1);

                    SasControls.ToolBarButton toolBarButton2 = new SasControls.ToolBarButton();
                    toolBarButton2.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton2.Name = "btnUpdate";
                    toolBarButton2.Text = "Cập nhật";
                    toolBarButton2.ToolTip = "F6";
                    toolBarButton2.ImagePath = "Images\\Edit.png";
                    toolBarButton2.Click += new RoutedEventHandler(ToolBarButtonF6_Click);
                    toolBar.Items.Insert(3, toolBarButton2);
                }


                FrmPoctpna.obrowseBKCT.frmBrw.LanguageID = "PODMHDM_brwBKCT";
                FrmPoctpna.obrowseBKCT.ShowDialog();
            }
            catch(Exception e)
            {

            }
            
        }
        private static void ToolBarButtonF4_Click(object sender, RoutedEventArgs e)
        {
            string ma_vt = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString().Trim();

            string stt_rec = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim();
            string stt_rec0 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec0"].ToString().Trim();
            SqlCommand sqlcmd = new SqlCommand("select dvt,packing from dmvt where ma_vt = @ma_vt");
            sqlcmd.Parameters.Add("@ma_vt", SqlDbType.NVarChar).Value = (object)ma_vt;

            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];


            FrmSetValue frmSetValue = new FrmSetValue(table);
            frmSetValue.ShowDialog();

            if (frmSetValue.DialogResult == true)
            {
                string packing = frmSetValue.txtpacking.Text.Trim();
                string he_so = frmSetValue.txthe_so.Text.Trim();
                string dvt = frmSetValue.txtDvt1.Text.Trim();
                string so_luong = frmSetValue.txtso_luong.Text.Trim();
                string ngay_giao = Convert.ToDateTime(frmSetValue.txtNgay_bh.Value)
                    .ToString("yyyy-MM-dd");


                string packingSql = Convert.ToDecimal(packing).ToString(CultureInfo.InvariantCulture);
                string heSoSql = Convert.ToDecimal(he_so).ToString(CultureInfo.InvariantCulture);
                string soLuongSql = Convert.ToDecimal(so_luong).ToString(CultureInfo.InvariantCulture);

                string updateCommand = string.Format(
                    @"INSERT INTO dmhdmctgh
      (stt_rec, stt_rec0, packing, he_so, dvt, so_luong, ngay_giao,ma_vt)
      VALUES
      ('{0}', '{1}', {2}, {3}, '{4}', {5}, '{6}','{7}')",
                    stt_rec, stt_rec0, packingSql, heSoSql, dvt, soLuongSql, ngay_giao,ma_vt);

                SqlCommand cmd = new SqlCommand(updateCommand);
                StartupBase.SasObj.ExcuteNonQuery(cmd);
                CallGridReportBKCT(false);

            }

        }


        private static void ToolBarButtonF5_Click(object sender, RoutedEventArgs e)
        {
            DataView dataView = FrmPoctpna.obrowseBKCT.DataGrid.DataSource as DataView;
            if (FrmPoctpna.obrowseBKCT.ActiveRecord == null)
                return;

            if (FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell != null && FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell.IsInEditMode)
                FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell.EndEditMode();
            FrmPoctpna.obrowseBKCT.ActiveRecord.Update();
            DataTable distinctValues = dataView.ToTable(true, "chon", "id");

            DataRowView[] array = (from DataRowView x in distinctValues.DefaultView
                                   where (bool)x["chon"]
                                   select x).ToArray();
            if (array.Length > 0)
            {
                string listSo_ct = string.Join(", ", array.Select((DataRowView x) => x["id"].ToString().Trim()).ToArray());
                foreach (DataRowView row in array)
                {
                    int id = Convert.ToInt32(row["id"]);

                    string sql = "DELETE FROM dmhdmctgh WHERE id = " + id;
                    SqlCommand cmd = new SqlCommand(sql);

                    StartUp.SasObj.ExcuteNonQuery(cmd);
                }
            }

        }

        private static void ToolBarButtonF6_Click(object sender, RoutedEventArgs e)
        {
            DataView dataView = FrmPoctpna.obrowseBKCT.DataGrid.DataSource as DataView;
            if (FrmPoctpna.obrowseBKCT.ActiveRecord == null)
                return;
            string ma_vt = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString().Trim();

            if (FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell != null && FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell.IsInEditMode)
                FrmPoctpna.obrowseBKCT.DataGrid.ActiveCell.EndEditMode();
            FrmPoctpna.obrowseBKCT.ActiveRecord.Update();
            DataView dv = new DataView(dataView.Table);
            dv.RowFilter = "chon = true";

            DataTable distinctValues = dv.ToTable(true, "chon", "id","dvt", "ngay_giao", "so_luong", "he_so", "packing");

            FrmSetValue frmSetValue = new FrmSetValue(distinctValues);
            frmSetValue.ShowDialog();

            if (frmSetValue.DialogResult == true)
            {
                string packing = frmSetValue.txtpacking.Text.Trim();
                string he_so = frmSetValue.txthe_so.Text.Trim();
                string dvt = frmSetValue.txtDvt1.Text.Trim();
                string so_luong = frmSetValue.txtso_luong.Text.Trim();
                string ngay_giao = Convert.ToDateTime(frmSetValue.txtNgay_bh.Value)
                    .ToString("yyyy-MM-dd");
                string id = distinctValues.Rows[0]["id"].ToString();

                string packingSql = Convert.ToDecimal(packing).ToString(CultureInfo.InvariantCulture);
                string heSoSql = Convert.ToDecimal(he_so).ToString(CultureInfo.InvariantCulture);
                string soLuongSql = Convert.ToDecimal(so_luong).ToString(CultureInfo.InvariantCulture);

                string updateCommand = string.Format(
      @"UPDATE dmhdmctgh
      SET packing = {1},
          he_so = {2},
          dvt = '{3}',
          so_luong = {4},
          ngay_giao = '{5}',ma_vt='{6}'
      WHERE id = {0}",
      id,
      packingSql,
      heSoSql,
      dvt,
      soLuongSql,
      ngay_giao, ma_vt);

                SqlCommand cmd = new SqlCommand(updateCommand);
                StartupBase.SasObj.ExcuteNonQuery(cmd);
                CallGridReportBKCT(false);

            }

        }

        private static void FormBrowse_Esc(object sender, EventArgs e)
        {
        }
        public static void obrowseBKCT_CTRL_R(object sender, EventArgs e)
        {
            CallGridReportBKCT(false);
        }
    }
}
