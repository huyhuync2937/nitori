using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasControls.ControlLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Linq;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using System.IO;
using System.Reflection;
using SasFormReport;

namespace COSXKSX.KSXS
{
    public partial class FrmPoctpna : FormTrans
    {
        public static int iRow = 0;
        public static int OldiRow = 0;
        public string Old_ma_kho = string.Empty;
        private string so_lsx_old;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        private DataSet DsVitual;
        public int so_ct_px_length = 12;
        public FrmPoctpna()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
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
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmPoctpna.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                this.LanguageProvider.Language = StartUpTrans.M_LAN;
                this.Voucher_Ma_nt0.Text = StartUpTrans.M_ma_nt0;
                this.Voucher_Ma_nt0.Value = true;

                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCtmaymoc, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCtnguonluc, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCtnguyenlieu, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.LoadData();
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                }
                //(this.Toolbar.FindName("btnPrint") as SasVoucherLib.ToolBarButton).Visibility = Visibility.Collapsed;
                //this.SetFocusToolbar();
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
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";

            this.txtstatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.GrdCtnguyenlieu.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[2].DefaultView;
            this.GrdCtmaymoc.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[3].DefaultView;
            this.GrdCtnguonluc.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[4].DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtstatus.IsEnabled = false;
        }

        private void V_Dau()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Truoc()
        {
            if (FrmPoctpna.iRow <= 1)
                return;
            --FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Sau()
        {
            if (FrmPoctpna.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmPoctpna.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Cuoi()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
        }

        private void V_Moi()
        {
            try
            {
                string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
                FormTrans.currActionTask = ActionTask.Add;
                if (string.IsNullOrEmpty(str))
                    return;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMabpht.IsFocus = true));
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)str;
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["status"] = (object)0;
                row["ma_qs"] = (object)"KSXS";
                row["ngay_kh1"] = row["ngay_kh2"] = row["ngay_ksx"] = (object)DateTime.Now.Date;
                row["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.NewRowCtsanpham();
                this.NewRowCtnguyenlieu();
                this.NewRowCtmaymoc();
                this.NewRowCtnguonluc();
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str + "'";

                FrmPoctpna.OldiRow = FrmPoctpna.iRow;
                FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode.Value = true;
                this.TabInfo.SelectedIndex = 0;
                GetSoCT();
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMabpht.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow].ItemArray;
                row1["stt_rec_px"] = (object)string.Empty;
                row1["ma_ct_px"] = (object)string.Empty;
                row1["so_ct_px"] = (object)string.Empty;
                row1["ma_qs_px"] = (object)string.Empty;
                row1["stt_rec"] = (object)str;
                row1["status"] = (object)1;
                row1["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs;
                row1["ngay_ksx"] = (object)FrmPoctpnaCopy.ngay_ct;
                row1["ma_qs"] = (object)string.Empty;
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_lsxtmp"] = row1["so_ct"];
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
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
                if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                {
                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"))
                    {
                        DataRow row2 = StartUpTrans.DsTrans.Tables[2].NewRow();
                        row2.ItemArray = dataRow.ItemArray;
                        row2["stt_rec"] = (object)str;
                        StartUpTrans.DsTrans.Tables[2].Rows.Add(row2);
                    }
                }
                if (StartUpTrans.DsTrans.Tables[3].DefaultView.Count > 0)
                {
                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[3].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"))
                    {
                        DataRow row2 = StartUpTrans.DsTrans.Tables[3].NewRow();
                        row2.ItemArray = dataRow.ItemArray;
                        row2["stt_rec"] = (object)str;
                        StartUpTrans.DsTrans.Tables[3].Rows.Add(row2);
                    }
                }
                if (StartUpTrans.DsTrans.Tables[4].DefaultView.Count > 0)
                {
                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[4].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'"))
                    {
                        DataRow row2 = StartUpTrans.DsTrans.Tables[4].NewRow();
                        row2.ItemArray = dataRow.ItemArray;
                        row2["stt_rec"] = (object)str;
                        StartUpTrans.DsTrans.Tables[4].Rows.Add(row2);
                    }
                }
                FrmPoctpna.OldiRow = FrmPoctpna.iRow;
                FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmPoctpna.IsInEditMode.Value = true;
            }
        }

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num = (int)ExMessageBox.Show(2215, StartupBase.SasObj, "Không có dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMabpht.IsFocus = true));
                FormTrans.currActionTask = ActionTask.Edit;
                this.so_lsx_old = this.txtSo_lsx.Text;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[3].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[4].DefaultView.ToTable());
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
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
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
                    if (StartUpTrans.DsTrans.Tables[3].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[3].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[3].Rows.Remove(row);
                    }
                    if (StartUpTrans.DsTrans.Tables[4].Rows.Count > 0)
                    {
                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[4].Select("stt_rec='" + str + "'"))
                            StartUpTrans.DsTrans.Tables[4].Rows.Remove(row);
                    }
                    StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmPoctpna.iRow);
                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row1, FrmPoctpna.iRow);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    StartUpTrans.DsTrans.Tables[2].Merge(this.DsVitual.Tables[2]);
                    StartUpTrans.DsTrans.Tables[3].Merge(this.DsVitual.Tables[3]);
                    StartUpTrans.DsTrans.Tables[4].Merge(this.DsVitual.Tables[4]);
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
                StartUp.DeleteVoucher(_stt_rec);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
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
                if (StartUpTrans.DsTrans.Tables[3].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[3].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[3].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[4].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[4].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[4].Rows.Remove(row);
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
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            dataTable.DefaultView.Sort = "ngay_ksx asc, so_ct asc";
            DataTable dataTable1 = dataTable.DefaultView.ToTable();
            FormView formView = new FormView(StartupBase.SasObj, dataTable1.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, strBrowse, strBrowseCt, "stt_rec");
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);

            formView.frmBrw.LanguageID = "COSXKSX.KSXS_4";
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
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str + "'";

            }
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
        }

        private void NewRowCtsanpham()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dataRow["ngay_kh1"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh1"];
                dataRow["ngay_kh2"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh2"];
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                int result = 0;
                int num1 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (result >= num1 ? result : num1) + 1;
                dataRow["stt_rec0"] = (object)string.Format("{0:000}", (object)num2);
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void NewRowCtnguyenlieu()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[2].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dataRow["ngay_kh1"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh1"];
                dataRow["ngay_kh2"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh2"];
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                dataRow["gia"] = (object)0;
                dataRow["tien"] = (object)0;
                int result = 0;
                int num1 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (result >= num1 ? result : num1) + 1;
                dataRow["stt_rec0"] = (object)string.Format("{0:000}", (object)num2);
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[2].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void NewRowCtmaymoc()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[3].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dataRow["ngay_kh1"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh1"];
                dataRow["ngay_kh2"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh2"];
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                int result = 0;
                int num1 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[3].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (result >= num1 ? result : num1) + 1;
                dataRow["stt_rec0"] = (object)string.Format("{0:000}", (object)num2);
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[3].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[3].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void NewRowCtnguonluc()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[4].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                dataRow["ngay_kh1"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh1"];
                dataRow["ngay_kh2"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh2"];
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                int result = 0;
                int num1 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[4].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result);
                }
                int num2 = (result >= num1 ? result : num1) + 1;
                dataRow["stt_rec0"] = (object)string.Format("{0:000}", (object)num2);
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[4].DefaultView, dataRow, 1);
                StartUpTrans.DsTrans.Tables[4].Rows.Add(dataRow);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCtsanpham();
            return true;
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmPoctpna.IsInEditMode.Value || (this.GrdCt.ActiveCell == null || StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) != null))
                    return;
                switch (e.Cell.Field.Name)
                {
                    case "ma_sp":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_sp"].Value = autoCompleteControl.RowResult["ten_vt"];
                            e.Cell.Record.Cells["ten_sp2"].Value = autoCompleteControl.RowResult["ten_vt2"];
                            e.Cell.Record.Cells["dvt"].Value = autoCompleteControl.RowResult["dvt"];

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
                    case "ma_cd":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_cd"].Value = autoCompleteControl1.RowResult["ten_px"];
                            e.Cell.Record.Cells["ten_cd2"].Value = autoCompleteControl1.RowResult["ten_px2"];
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
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_sp"].Value == null || activeRecord.Cells["ma_sp"].Value.ToString() == "")
                        break;
                    this.NewRowCtsanpham();
                    this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                    this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_sp"];
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
                    }
                    break;
            }
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_sp"].Value == null || activeRecord.Cells["ma_sp"].Value.ToString() == ""))
                return;
            this.NewRowCtsanpham();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            this.GrdCt.ActiveCell = (this.GrdCt.Records[this.GrdCt.Records.Count - 1] as DataRecord).Cells["ma_sp"];
        }

        private void V_Nhan()
        {
            try
            {
                StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                StartUpTrans.DsTrans.Tables[3].AcceptChanges();
                StartUpTrans.DsTrans.Tables[4].AcceptChanges();
                bool flag = false;
                if (!this.IsSequenceSave)
                {
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2230, StartupBase.SasObj, "Chưa vào số kế hoạch sản xuất!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtSo_lsx.Focus();
                        flag = true;
                    }
                    else
                    {
                        string str = SysFunc.CheckInValidCode(StartupBase.SasObj, this.txtSo_lsx.Text.Trim());
                        if (str != "")
                        {
                            int num = (int)ExMessageBox.Show(1940, StartupBase.SasObj, "Số kế hoạch sản xuất không được chứa các ký tự [" + str + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_lsx.SelectAll();
                            this.txtSo_lsx.Focus();
                            flag = true;
                        }
                        SqlCommand sqlcmd = new SqlCommand();
                        if (FormTrans.currActionTask == ActionTask.Edit)
                        {
                            sqlcmd.CommandText = "Select so_ct From ksxphs Where  rtrim(ltrim(so_ct)) = @so_ct and rtrim(ltrim(so_ct)) <> @so_lsx_cu";
                            sqlcmd.Parameters.Add("@so_ct", SqlDbType.Char, 16).Value = (object)this.txtSo_lsx.Text.ToString().Trim();
                            sqlcmd.Parameters.Add("@so_lsx_cu", SqlDbType.Char, 16).Value = (object)this.so_lsx_old.Trim();
                        }
                        else
                        {
                            sqlcmd.CommandText = "Select so_ct From ksxphs Where so_ct = @so_ct";
                            sqlcmd.Parameters.Add("@so_ct", SqlDbType.Char, 16).Value = (object)this.txtSo_lsx.Text.ToString();
                        }
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                        {
                            int num = (int)ExMessageBox.Show(4595, StartupBase.SasObj, "Số chứng từ kế hoạch sản xuất đã có!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_lsx.Focus();
                            flag = true;
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(this.txtNgaybd_kh.Text))
                            {
                                int num = (int)ExMessageBox.Show(1975, StartupBase.SasObj, "Ngày bắt đầu kế hoạch sản xuất không được để trống!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtNgaybd_kh.Focus();
                                return;
                            }
                            if (string.IsNullOrEmpty(this.txtNgaykt_kh.Text))
                            {
                                int num = (int)ExMessageBox.Show(1975, StartupBase.SasObj, "Ngày kết thúc kế hoạch sản xuất không được để trống!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtNgaykt_kh.Focus();
                                return;
                            }
                            DateTime dateTime1 = Convert.ToDateTime(this.txtngay_lsx.dValue);
                            DateTime? nullable = StartUp.M_NGAY_BAT_DAU;
                            if ((nullable.HasValue ? (dateTime1 < nullable.GetValueOrDefault() ? 1 : 0) : 0) != 0)
                            {
                                int num = (int)ExMessageBox.Show(1905, StartupBase.SasObj, "Ngày kế hoạch sản xuất không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtngay_lsx.Focus();
                                flag = true;
                            }
                            else
                            {
                                DateTime dateTime2 = Convert.ToDateTime(this.txtngay_lsx.dValue);
                                nullable = StartUp.M_NGAY_KET_THUC;
                                if ((nullable.HasValue ? (dateTime2 > nullable.GetValueOrDefault() ? 1 : 0) : 0) != 0)
                                {
                                    int num = (int)ExMessageBox.Show(1910, StartupBase.SasObj, "Ngày kế hoạch sản xuất không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.txtngay_lsx.Focus();
                                    flag = true;
                                }
                            }
                        }
                    }
                    if (!flag)
                    {
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                        {
                            int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào chi tiết, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.GrdCt.Focus();
                            return;
                        }
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_sp"].ToString()))
                            {
                                StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[index].Row);
                                //int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào chi tiết, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                //this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_sp"];
                                // this.GrdCt.Focus();
                                // return;

                            }
                        }
                    }
                }
                if (!flag)
                {
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("search"))
                    {
                        DataTable table = StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable();
                        SysFunc.SetStrSearch(StartupBase.SasObj, "ksxphs", ref table);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["search"] = (object)table.Rows[0]["search"].ToString().Trim();
                    }
                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    LocalTable1.Rows[0]["so_ct"] = LocalTable1.Rows[0]["so_ct"];
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                    DataTable LocalTable3 = StartUpTrans.DsTrans.Tables[2].Clone();
                    DataTable LocalTable4 = StartUpTrans.DsTrans.Tables[3].Clone();
                    DataTable LocalTable5 = StartUpTrans.DsTrans.Tables[4].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                    {
                        LocalTable3.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[3].DefaultView)
                    {
                        LocalTable4.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[4].DefaultView)
                    {
                        LocalTable5.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUp.tablesanpham, LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUp.tablenguyenlieu, LocalTable3, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUp.tablemaymoc, LocalTable4, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUp.tablenguonluc, LocalTable5, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else if (!flag && !this.IsSequenceSave)
                    {
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
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
                        }
                        FormTrans.currActionTask = ActionTask.View;
                        FrmPoctpna.IsInEditMode.Value = false;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private bool GrdCp_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D3);
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
            return false;
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void txtMabpht_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            if (this.txtMabpht.RowResult == null)
                this.tblList_Bp.Text = "";
            else
                this.tblList_Bp.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMabpht.RowResult["ten_bp"].ToString() : this.txtMabpht.RowResult["ten_bp2"].ToString();
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            GetSoCT();
        }
        void GetSoCT()
        {
            if (!FrmPoctpna.IsInEditMode.Value || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                return;
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
            {
                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsxtmp"].ToString().Trim()) || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"].ToString().Trim()))
                {
                    this.txtSo_lsx.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsxtmp"] = (object)this.txtSo_lsx.Text;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
                }
                else
                    this.txtSo_lsx.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsxtmp"].ToString().Trim();
            }
            if (this.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtSo_lsx.Text, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
            {
                this.txtSo_lsx.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsxtmp"] = (object)this.txtSo_lsx.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
            }
        }
        private void txtMa_dvcs_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            this.lbtenMa_dvcs.Text = this.txtMa_dvcs.RowResult != null ? (StartUpTrans.M_LAN.Equals("V") ? this.txtMa_dvcs.RowResult["ten_dvcs"].ToString() : this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString()) : "";
        }

        private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || !(this.txtstatus.Text == ""))
                return;
            this.txtstatus.Text = "1";
        }

        private void BtnXuatNVL_Click(object sender, RoutedEventArgs e)
        {
            if (FrmPoctpna.IsInEditMode.Value)
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
            dt.Columns.Add("ong_ba", typeof(string));
            dt.Columns.Add("dien_giai", typeof(string));
            dt.Columns.Add("ma_gd", typeof(string));
            dt.Columns.Add("ma_kh", typeof(string));
            dt.Columns.Add("ma_kho", typeof(string));
            dt.Columns.Add("ma_nx", typeof(string));

            DataRow row = dt.NewRow();
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                newstt_recPXD = dataTable.Rows[0]["stt_rec"].ToString();
                row["stt_recpx"] = newstt_recPXD;
                row["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                row["ma_ct"] = "PXD";
                row["ma_nt"] = dataTable.Rows[0]["ma_nt"].ToString();
                row["ma_gd"] = dataTable.Rows[0]["ma_gd"].ToString();
                row["ma_qs"] = dataTable.Rows[0]["ma_qs"].ToString();
                row["so_ct"] = dataTable.Rows[0]["so_ct"].ToString();
                row["ma_kh"] = dataTable.Rows[0]["ma_kh"].ToString();
                row["ma_kho"] = dataTable.Rows[0]["ma_kho"].ToString();
                row["ong_ba"] = dataTable.Rows[0]["ong_ba"].ToString();
                row["ma_nx"] = dataTable.Rows[0]["ma_nx"].ToString();
                row["dien_giai"] = dataTable.Rows[0]["dien_giai"].ToString();
            }
            else
            {
                row["stt_recpx"] = newstt_recPXD;
                row["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                row["ma_ct"] = "PXD";
                row["ma_nt"] = StartUp.M_ma_nt0;
                row["ma_qs"] = "";
                row["ma_gd"] = "4";
                row["so_ct"] = "";
                row["ma_kh"] = "";
                row["ma_kho"] = "";
                row["ong_ba"] = "";
                row["ma_nx"] = "";
                row["dien_giai"] = !StartUpTrans.M_LAN.Equals("V") ? string.Format("Export NVL {0}, invoice date {1}", this.txtSo_lsx.Text.Trim(), this.txtngay_lsx.dValue.ToShortDateString()) : string.Format("Xuất NVL hóa đơn số {0}, ngày {1}", this.txtSo_lsx.Text.Trim(), this.txtngay_lsx.dValue.ToShortDateString());
            }
            dt.Rows.Add(row);

            FrmTaoPXNVL frmTaoPXNVL = new FrmTaoPXNVL();
            frmTaoPXNVL.DataContext = dt.DefaultView;
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
                dt.Rows[0]["ty_gia"] = 1;
                dt.Rows[0]["ty_giaf"] = 1;

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

                if (isSuccess)
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Successfully created NVL export slip!" : "Tạo phiếu xuất NVL thành công!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                else
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Create an NVL output slip with error. Please check the quota declaration!" : "Tạo phiếu xuất NVL bị lỗi.Hãy kiểm tra lại khai báo định mức!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());

            }
        }

        private void CreatePXD(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[COSXKSX.KSXS-CREATEPXD] @Stt_rec, @Stt_recpx, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @ong_ba, @dien_giai, @ma_gd, @ma_ct, @ma_kh, @ma_kho,@ma_nx");
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = dt.Rows[0]["stt_rec"];
                sqlcmd.Parameters.Add("@Stt_recpx", SqlDbType.VarChar).Value = dt.Rows[0]["stt_recpx"];
                sqlcmd.Parameters.Add("@ma_qs", SqlDbType.VarChar).Value = dt.Rows[0]["ma_qs"];
                sqlcmd.Parameters.Add("@so_ct", SqlDbType.VarChar).Value = dt.Rows[0]["so_ct"];
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nt"];
                sqlcmd.Parameters.Add("@ty_gia", SqlDbType.Decimal).Value = dt.Rows[0]["ty_gia"];
                sqlcmd.Parameters.Add("@ty_giaf", SqlDbType.Decimal).Value = dt.Rows[0]["ty_giaf"];
                sqlcmd.Parameters.Add("@ong_ba", SqlDbType.NVarChar).Value = dt.Rows[0]["ong_ba"];
                sqlcmd.Parameters.Add("@dien_giai", SqlDbType.NVarChar).Value = dt.Rows[0]["dien_giai"];
                sqlcmd.Parameters.Add("@ma_gd", SqlDbType.VarChar).Value = dt.Rows[0]["ma_gd"];
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = dt.Rows[0]["ma_ct"];
                sqlcmd.Parameters.Add("@ma_kh", SqlDbType.VarChar).Value = dt.Rows[0]["ma_kh"];
                sqlcmd.Parameters.Add("@ma_kho", SqlDbType.VarChar).Value = dt.Rows[0]["ma_kho"];
                sqlcmd.Parameters.Add("@ma_nx", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nx"];
                ErrorLog.WriteToLogFile("", StartupBase.SasObj.SqlString(sqlcmd));
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
                sqlcmd2.CommandText = "UPDATE ph81 Set stt_rec_px = '', so_ct_px = '', ma_ct_px = '', ma_ps_px = '' WHERE stt_rec = @stt_rec_px; ";
                sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_px", SqlDbType.VarChar)).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_px"] = "";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"] = "";
            }
            else
            {
                SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "PXD", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString(), Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), this.BindingSasObj.M_ProcessName);
            }
        }
        private void V_In()
        {
            FrmIn frmIn = new FrmIn();
            if (this.M_LAN != "V")
                frmIn.Title = "Print";
            frmIn.ShowDialog();
            //DataTable tbl1 = StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable();
            //tbl1.TableName = "tblPH";
            //DataTable tbl2 = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable();
            //tbl2.TableName = "tblsanpham";
            //DataTable tbl3 = StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable();
            //tbl3.TableName = "tblnguyenlieu";
            //DataTable tbl4 = StartUpTrans.DsTrans.Tables[3].DefaultView.ToTable();
            //tbl4.TableName = "tblmaymoc";
            //DataTable tbl5 = StartUpTrans.DsTrans.Tables[4].DefaultView.ToTable();
            //tbl5.TableName = "tblnguonluc";
            //StartUp.DataSource.Tables.Add(tbl1);
            //StartUp.DataSource.Tables.Add(tbl2);
            //StartUp.DataSource.Tables.Add(tbl3);
            //StartUp.DataSource.Tables.Add(tbl4);
            //StartUp.DataSource.Tables.Add(tbl5);
            //new ReportManager(StartupBase.SasObj, StartUp.CommandInfo["rep_file"].ToString(), 0).Preview(StartUp.DataSource);
        }

        private void txtDoitac_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            if (this.txtDoitac.RowResult == null)
                this.tblTenDt.Text = "";
            else
                this.tblTenDt.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtDoitac.RowResult["ten_kh"].ToString() : this.txtDoitac.RowResult["ten_kh2"].ToString();
        }

        private void GrdCtnguyenlieu_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmPoctpna.IsInEditMode.Value || (this.GrdCtnguyenlieu.ActiveCell == null || StartUpTrans.DsTrans.Tables[2].GetChanges(DataRowState.Deleted) != null))
                    return;
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl.RowResult["ten_vt"];
                            e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl.RowResult["ten_vt2"];
                            e.Cell.Record.Cells["dvt"].Value = autoCompleteControl.RowResult["dvt"];
                            DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                            CellCollection cells = e.Cell.Record.Cells;
                            if (string.IsNullOrEmpty(cells["dvt1"].Value.ToString()))
                            {
                                cells["dvt1"].Value = cells["dvt"].Value;
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
                                    }
                                }
                            }
                        }
                        break;
                    case "ma_cd":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_cd"].Value = autoCompleteControl1.RowResult["ten_px"];
                            e.Cell.Record.Cells["ten_cd2"].Value = autoCompleteControl1.RowResult["ten_px2"];
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCtnguyenlieu_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCtnguyenlieu_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCtnguyenlieu.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == "")
                        break;
                    this.NewRowCtnguyenlieu();
                    this.GrdCtnguyenlieu.ActiveRecord = this.GrdCtnguyenlieu.Records[this.GrdCtnguyenlieu.Records.Count - 1];
                    this.GrdCtnguyenlieu.ActiveCell = (this.GrdCtnguyenlieu.ActiveRecord as DataRecord).Cells["ma_cd"];
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(2225, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtnguyenlieu.ActiveRecord is DataRecord activeRecord1))
                        break;
                    int num1 = 0;
                    Cell activeCell = this.GrdCtnguyenlieu.ActiveCell;
                    if (activeRecord1.Index == 0)
                    {
                        if (this.GrdCtnguyenlieu.Records.Count == 1)
                            this.GrdCtnguyenlieu_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord1.Index == this.GrdCtnguyenlieu.Records.Count - 1)
                        num1 = activeRecord1.Index - 1;
                    int num2 = this.GrdCtnguyenlieu.ActiveCell == null ? 0 : this.GrdCtnguyenlieu.ActiveCell.Field.Index;
                    this.GrdCtnguyenlieu.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[2].Rows.Remove(StartUpTrans.DsTrans.Tables[2].DefaultView[activeRecord1.Index].Row);
                        StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                        if (this.GrdCtnguyenlieu.Records.Count > 0)
                            this.GrdCtnguyenlieu.ActiveRecord = this.GrdCtnguyenlieu.Records[num1 > this.GrdCtnguyenlieu.Records.Count - 1 ? this.GrdCtnguyenlieu.Records.Count - 1 : num1];
                    }
                    break;
            }
        }

        private void GrdCtnguyenlieu_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCtnguyenlieu.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == ""))
                return;
            this.NewRowCtnguyenlieu();
            this.GrdCtnguyenlieu.ActiveRecord = this.GrdCtnguyenlieu.Records[this.GrdCtnguyenlieu.Records.Count - 1];
            this.GrdCtnguyenlieu.ActiveCell = (this.GrdCtnguyenlieu.Records[this.GrdCtnguyenlieu.Records.Count - 1] as DataRecord).Cells["ma_cd"];
        }

        private bool GrdCtnguyenlieu_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCtnguyenlieu();
            return true;
        }

        private bool GrdCtmaymoc_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCtmaymoc();
            return true;
        }

        private void GrdCtmaymoc_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCtmaymoc_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCtmaymoc.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == "")
                        break;
                    this.NewRowCtmaymoc();
                    this.GrdCtmaymoc.ActiveRecord = this.GrdCtmaymoc.Records[this.GrdCtmaymoc.Records.Count - 1];
                    this.GrdCtmaymoc.ActiveCell = (this.GrdCtmaymoc.ActiveRecord as DataRecord).Cells["ma_cd"];
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(2225, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtmaymoc.ActiveRecord is DataRecord activeRecord1))
                        break;
                    int num1 = 0;
                    Cell activeCell = this.GrdCtmaymoc.ActiveCell;
                    if (activeRecord1.Index == 0)
                    {
                        if (this.GrdCtmaymoc.Records.Count == 1)
                            this.GrdCtmaymoc_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord1.Index == this.GrdCtmaymoc.Records.Count - 1)
                        num1 = activeRecord1.Index - 1;
                    int num2 = this.GrdCtmaymoc.ActiveCell == null ? 0 : this.GrdCtmaymoc.ActiveCell.Field.Index;
                    this.GrdCtmaymoc.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[3].Rows.Remove(StartUpTrans.DsTrans.Tables[3].DefaultView[activeRecord1.Index].Row);
                        StartUpTrans.DsTrans.Tables[3].AcceptChanges();
                        if (this.GrdCtmaymoc.Records.Count > 0)
                            this.GrdCtmaymoc.ActiveRecord = this.GrdCtmaymoc.Records[num1 > this.GrdCtmaymoc.Records.Count - 1 ? this.GrdCtmaymoc.Records.Count - 1 : num1];
                    }
                    break;
            }
        }

        private void GrdCtmaymoc_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCtmaymoc.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == ""))
                return;
            this.NewRowCtmaymoc();
            this.GrdCtmaymoc.ActiveRecord = this.GrdCtmaymoc.Records[this.GrdCtmaymoc.Records.Count - 1];
            this.GrdCtmaymoc.ActiveCell = (this.GrdCtmaymoc.Records[this.GrdCtmaymoc.Records.Count - 1] as DataRecord).Cells["ma_cd"];
        }

        private void GrdCtmaymoc_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmPoctpna.IsInEditMode.Value || (this.GrdCtmaymoc.ActiveCell == null || StartUpTrans.DsTrans.Tables[3].GetChanges(DataRowState.Deleted) != null))
                    return;
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl.RowResult["ten_ts"];
                            e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl.RowResult["ten_ts2"];
                            e.Cell.Record.Cells["dvt"].Value = autoCompleteControl.RowResult["dvt"];
                            DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                            CellCollection cells = e.Cell.Record.Cells;
                            if (string.IsNullOrEmpty(cells["dvt"].Value.ToString()))
                            {
                                AutoCompleteTextBox autoCompleteControlmavtdvt1 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["dvt1"]).Editor as ControlHostEditor);
                                if (autoCompleteControlmavtdvt1 != null)
                                {
                                    autoCompleteControlmavtdvt1.SearchInit();
                                    if (autoCompleteControlmavtdvt1.RowResult == null)
                                    {
                                        cells["dvt"].Value = cells["dvt"].Value;
                                        autoCompleteControlmavtdvt1.SearchInit();
                                    }
                                }                              
                            }                           
                        }
                        break;
                    case "ma_cd":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_cd"].Value = autoCompleteControl1.RowResult["ten_px"];
                            e.Cell.Record.Cells["ten_cd2"].Value = autoCompleteControl1.RowResult["ten_px2"];
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCtnguonluc_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!FrmPoctpna.IsInEditMode.Value || (this.GrdCtnguonluc.ActiveCell == null || StartUpTrans.DsTrans.Tables[4].GetChanges(DataRowState.Deleted) != null))
                    return;
                switch (e.Cell.Field.Name)
                {
                    case "ma_ns":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_ns"].Value = autoCompleteControl.RowResult["ten_ns"];
                            e.Cell.Record.Cells["ten_ns2"].Value = autoCompleteControl.RowResult["ten_ns2"];
                        }
                        break;
                    case "ma_cd":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl1.RowResult != null)
                        {
                            e.Cell.Record.Cells["ten_cd"].Value = autoCompleteControl1.RowResult["ten_px"];
                            e.Cell.Record.Cells["ten_cd2"].Value = autoCompleteControl1.RowResult["ten_px2"];
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private bool GrdCtnguonluc_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCtnguonluc();
            return true;
        }

        private void GrdCtnguonluc_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)) || (!(this.GrdCtnguonluc.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == ""))
                return;
            this.NewRowCtnguonluc();
            this.GrdCtnguonluc.ActiveRecord = this.GrdCtnguonluc.Records[this.GrdCtnguonluc.Records.Count - 1];
            this.GrdCtnguonluc.ActiveCell = (this.GrdCtnguonluc.Records[this.GrdCtnguonluc.Records.Count - 1] as DataRecord).Cells["ma_cd"];
        }

        private void GrdCtnguonluc_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCtnguonluc.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_cd"].Value == null || activeRecord.Cells["ma_cd"].Value.ToString() == "")
                        break;
                    this.NewRowCtnguonluc();
                    this.GrdCtnguonluc.ActiveRecord = this.GrdCtnguonluc.Records[this.GrdCtnguonluc.Records.Count - 1];
                    this.GrdCtnguonluc.ActiveCell = (this.GrdCtnguonluc.ActiveRecord as DataRecord).Cells["ma_cd"];
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(2225, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtnguonluc.ActiveRecord is DataRecord activeRecord1))
                        break;
                    int num1 = 0;
                    Cell activeCell = this.GrdCtnguonluc.ActiveCell;
                    if (activeRecord1.Index == 0)
                    {
                        if (this.GrdCtnguonluc.Records.Count == 1)
                            this.GrdCtnguonluc_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                    }
                    else if (activeRecord1.Index == this.GrdCtnguonluc.Records.Count - 1)
                        num1 = activeRecord1.Index - 1;
                    int num2 = this.GrdCtnguonluc.ActiveCell == null ? 0 : this.GrdCtnguonluc.ActiveCell.Field.Index;
                    this.GrdCtnguonluc.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[4].Rows.Remove(StartUpTrans.DsTrans.Tables[4].DefaultView[activeRecord1.Index].Row);
                        StartUpTrans.DsTrans.Tables[4].AcceptChanges();
                        if (this.GrdCtnguonluc.Records.Count > 0)
                            this.GrdCtnguonluc.ActiveRecord = this.GrdCtnguonluc.Records[num1 > this.GrdCtnguonluc.Records.Count - 1 ? this.GrdCtnguonluc.Records.Count - 1 : num1];
                    }
                    break;
            }
        }

        private void GrdCtnguonluc_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }      

        private void BtnTAODMNVL_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isadd = false;              
                SqlCommand cmd;
                string sqlmaky = "Select top(1) ma_ky from dmkygt where @ngay between ngay1 and ngay2";
                cmd = new SqlCommand(sqlmaky);
                cmd.Parameters.Add("@ngay", SqlDbType.SmallDateTime).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ksx"].ToString()).ToString("yyyy-MM-dd");
                DataTable tblmaky = StartUp.SasObj.ExcuteReader(cmd).Tables[0];
                if (tblmaky == null || tblmaky.Rows.Count < 1)
                {
                    MessageBox.Show("Chưa khai báo kỳ giá thành.", "Thông báo");
                    return;
                }

                foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim() + "'"))
                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                foreach (DataRowView view in StartUpTrans.DsTrans.Tables[1].DefaultView)
                {
                    string sql = "select a.ma_vt,sum(a.sl_dm)sl_dm, max(b.ten_vt)ten_vt, max(b.ten_vt2)ten_vt2, max(b.dvt)dvt from [cosxlsx-Dmdmvtct] a left join dmvt b on a.ma_vt = b.ma_vt  where a.ma_sp = @ma_sp  and a.ma_ky= @ma_ky group by a.ma_vt";
                    cmd = new SqlCommand(sql);
                    cmd.Parameters.Add("@ma_sp", SqlDbType.Char).Value = (object)view["ma_sp"].ToString().Trim();
                    cmd.Parameters.Add("@ma_ky", SqlDbType.Char).Value = (object)tblmaky.Rows[0]["ma_ky"].ToString().Trim();
                    DataTable tblPhanbo = StartUp.SasObj.ExcuteReader(cmd).Tables[0];

                    if (tblPhanbo != null && tblPhanbo.Rows.Count > 0)
                    {
                        foreach (DataRow row in tblPhanbo.Rows)
                        {
                            isadd = false;
                            foreach (DataRowView r in StartUpTrans.DsTrans.Tables[2].DefaultView)
                            {
                                if (r["ma_vt"].ToString().Trim().Equals(row["ma_vt"].ToString().Trim()))
                                {
                                    double soluong_sp = 0;
                                    double soluong_dm = 0;
                                    double.TryParse(view["so_luong"].ToString(), out soluong_sp);
                                    double.TryParse(row["sl_dm"].ToString(), out soluong_dm);
                                    double oldSoluong = Convert.ToDouble(r["so_luong"]);
                                    r["so_luong"] = oldSoluong + (soluong_sp * soluong_dm);
                                    isadd = true;
                                    break;
                                }                               
                            }
                            if (!isadd)
                            {
                                DataRow dataRow = StartUpTrans.DsTrans.Tables[2].NewRow();
                                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                dataRow["ngay_kh1"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh1"];
                                dataRow["ngay_kh2"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_kh2"];
                                dataRow["ma_vt"] = row["ma_vt"];
                                dataRow["ten_vt"] = row["ten_vt"];
                                dataRow["ma_cd"] = view["ma_cd"];
                                dataRow["ten_cd"] = view["ten_cd"];
                                dataRow["ten_cd2"] = view["ten_cd2"];
                                dataRow["ma_sp"] = view["ma_sp"];
                                dataRow["he_so1"] = (object)0;
                                dataRow["so_luong1"] = (object)0;
                                dataRow["gia"] = (object)0;
                                dataRow["tien"] = (object)0;
                                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                                double soluong_sp = 0;
                                double soluong_dm = 0;
                                double.TryParse(view["so_luong"].ToString(), out soluong_sp);
                                double.TryParse(row["sl_dm"].ToString(), out soluong_dm);
                                dataRow["so_luong"] = soluong_sp * soluong_dm;
                                dataRow["dvt1"] = row["dvt"];
                                dataRow["dvt"] = row["dvt"];
                                int result = 0;
                                int num1 = 0;
                                if (this.GrdCt.Records.Count > 0)
                                {
                                    string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                                    if (str != null)
                                        int.TryParse(str.ToString(), out result);
                                }
                                int num2 = (result >= num1 ? result : num1) + 1;
                                dataRow["stt_rec0"] = (object)string.Format("{0:000}", (object)num2);
                                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[2].DefaultView, dataRow, 1);
                                StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow);
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
        private void BtnDMNVL_Click(object sender, RoutedEventArgs e)
        {
            SysFunc.CallModule("COSXLSX.CODMNVL.exe", new string[1] { StartupBase.Menu_Id }, StartupBase.SasObj.M_StartUp_Path, StartupBase.SasObj.M_ProcessName, StartupBase.SasObj);
        }

        private void btnApgiavon_Click(object sender, RoutedEventArgs e)
        {
            FrmApgiavon frmgiavon = new FrmApgiavon();
            frmgiavon.ShowDialog();
        }
    }
}
