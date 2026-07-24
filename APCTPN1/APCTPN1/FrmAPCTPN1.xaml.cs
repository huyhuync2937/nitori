using ArapLib;
using Infragistics.Windows.Controls;
using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
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

namespace APCTPN1
{
    public partial class FrmAPCTPN1 : FormTrans
    {
        public static int iRow = 0;
        private int iOldRow = 0;
        private bool txtDiaChiFocusable = true;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject Voucher_Lan0;
        private DataSet dsCheckData;
        private DataSet DsVitual;

        public FrmAPCTPN1()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.BindingSasObj = StartupBase.SasObj;
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.cbMa_nt;
            this.C_So_ct = this.txtSo_ct;
        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmAPCTPN1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                DateTime now = DateTime.Now;
                Debug.WriteLine(string.Format("#5: {0}", (object)now.ToString()));
                FrmAPCTPN1.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtSoDuKH.Visibility = this.tblSoDuKH.Visibility = Visibility.Hidden;
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmAPCTPN1.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                this.GrdCt.Lan = this.M_LAN;
                this.GrdCtgt.Lan = this.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCtgt, StartUpTrans.Ma_ct, 2);
                now = DateTime.Now;
                Debug.WriteLine(string.Format("#6: {0}", (object)now.ToString()));
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                    this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
                    this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
                    this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
                    this.GrdCtgt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[2].DefaultView;
                    this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
                    if (StartUpTrans.tbStatus.DefaultView.Count == 1)
                        this.txtStatus.IsEnabled = false;
                    StartUpTrans.DsTrans.Tables[1].DefaultView.ListChanged += new ListChangedEventHandler(this.DefaultView_ListChanged);
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = true;
                }
                now = DateTime.Now;
                Debug.WriteLine(string.Format("#7: {0}", (object)now.ToString()));
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                now = DateTime.Now;
                Debug.WriteLine(string.Format("#8: {0}", (object)now.ToString()));
                this.SetStatusVisibleField();
                now = DateTime.Now;
                Debug.WriteLine(string.Format("#9: {0}", (object)now.ToString()));
                this.LoadDataDu13();
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

        private void DefaultView_ListChanged(object sender, ListChangedEventArgs e)
        {
            this.SetStatusVisibleField();
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.ChkSuaTien_Click(sender, e);
            if (menuItemName.Equals("btnSave"))
                return;
            this.LoadDataDu13();
        }

        private void V_Truoc()
        {
            if (FrmAPCTPN1.iRow <= 1)
                return;
            --FrmAPCTPN1.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Sau()
        {
            if (FrmAPCTPN1.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmAPCTPN1.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Dau()
        {
            FrmAPCTPN1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Cuoi()
        {
            FrmAPCTPN1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Huy()
        {
            FrmAPCTPN1.IsInEditMode.Value = false;
            if (this.DsVitual == null || StartUpTrans.DsTrans.Tables[0].Rows.Count <= 0)
                return;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmAPCTPN1.iRow = this.iOldRow;
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
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
                    StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmAPCTPN1.iRow);
                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row1.ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row1, FrmAPCTPN1.iRow);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    StartUpTrans.DsTrans.Tables[2].Merge(this.DsVitual.Tables[2]);
                    ChkTaoPc.IsEnabled = false;
                    break;
            }
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim()) && ExMessageBox.Show(391, StartupBase.SasObj, "Phiếu nhập đã được thanh toán, có muốn xóa phiếu thanh toán hay không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                this.DeleteVoucherPC(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pc"].ToString().Trim());
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUpTrans.UpdateTkSd13(1, 0);
                StartUp.DeleteVoucher(_stt_rec);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmAPCTPN1.iRow);
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
                    FrmAPCTPN1.iRow = FrmAPCTPN1.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmAPCTPN1.iRow - 1 : FrmAPCTPN1.iRow;
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["stt_rec"].ToString() + "'";
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
            new FrmIn().ShowDialog();
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            FrmAPCTPN1Copy frmApctpN1Copy = new FrmAPCTPN1Copy();
            frmApctpN1Copy.Closed += new EventHandler(this._formcopy_Closed);
            frmApctpN1Copy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!(sender as FrmAPCTPN1Copy).isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow].ItemArray;
                row1["stt_rec"] = (object)str;
                row1["ngay_ct"] = (object)FrmAPCTPN1Copy.ngay_ct;
                row1["ngay_lct"] = (object)FrmAPCTPN1Copy.ngay_ct;
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
                this.iOldRow = FrmAPCTPN1.iRow;
                FrmAPCTPN1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmAPCTPN1.IsInEditMode.Value = true;
                this.SetStatusVisibleField();
                this.UpdateTotalThue();
            }
        }

        private void NewRowCt()
        {
            DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
            dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            int result1 = 0;
            int result2 = 0;
            if (this.GrdCt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result1);
            }
            if (this.GrdCtgt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result2);
            }
            int num = (result1 >= result2 ? result1 : result2) + 1;
            dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
            dataRow["tien_nt"] = (object)0;
            dataRow["tien"] = (object)0;
            int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
            dataRow["dien_giaii"] = count <= 0 ? StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row["dien_giai"] : StartUpTrans.DsTrans.Tables[1].DefaultView[count - 1].Row["dien_giaii"];
            FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
            StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
        }

        private bool GrdCt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            this.NewRowCt();
            return true;
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void GrdCt_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                Decimal num1;
                Decimal num2;
                Decimal num3;
                Decimal num4;
                if (this.IsEditMode && this.GrdCt.ActiveCell != null && StartUpTrans.DsTrans.Tables[1].DefaultView.Count > this.GrdCt.ActiveRecord.Index && StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "tk_vt":
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl.RowResult != null && !autoCompleteControl.Text.Trim().Equals(""))
                            {
                                e.Cell.Record.Cells["ten_tk"].Value = autoCompleteControl.RowResult["ten_tk"];
                                e.Cell.Record.Cells["ten_tk2"].Value = autoCompleteControl.RowResult["ten_tk2"];
                            }
                            if (autoCompleteControl.Text.Trim().Equals(""))
                            {
                                e.Cell.Record.Cells["ten_tk"].Value = (object)"";
                                e.Cell.Record.Cells["ten_tk2"].Value = (object)"";
                                break;
                            }
                            break;
                        case "so_luong":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                                Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal nValue = this.txtTy_gia.nValue;
                                if (num6 == new Decimal(0))
                                {
                                    e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                    e.Cell.Record.Cells["gia"].Value = (object)0;
                                    break;
                                }
                                if (num5 * num6 != new Decimal(0))
                                {
                                    Decimal num7 = !this.cbMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? SysFunc.Round(num5 * num6, StartUpTrans.M_ROUND_NT) : SysFunc.Round(num5 * num6, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)num7;
                                    if (!this.ChkSuaTien.IsChecked.Value)
                                        e.Cell.Record.Cells["tien"].Value = (object)(num7 * nValue == new Decimal(0) ? num7 : SysFunc.Round(num7 * nValue, StartUpTrans.M_ROUND));
                                    this.GrdCt_EditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien_nt"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien_nt"]).Editor, true));
                                }
                                break;
                            }
                            break;
                        case "gia_nt":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                                Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal nValue = this.txtTy_gia.nValue;
                                if (num5 * num6 != new Decimal(0))
                                {
                                    Decimal num7 = !this.cbMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? SysFunc.Round(num5 * num6, StartUpTrans.M_ROUND_NT) : SysFunc.Round(num5 * num6, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)num7;
                                    if (!this.ChkSuaTien.IsChecked.Value)
                                    {
                                        e.Cell.Record.Cells["gia"].Value = (object)(num7 * nValue == new Decimal(0) ? num7 : SysFunc.Round(num5 * nValue, StartUpTrans.M_ROUND_GIA));
                                        e.Cell.Record.Cells["tien"].Value = (object)(num7 * nValue == new Decimal(0) ? num7 : SysFunc.Round(num7 * nValue, StartUpTrans.M_ROUND));
                                    }
                                }
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    e.Cell.Record.Cells["gia"].Value = e.Cell.Record.Cells["gia_nt"].Value;
                                this.GrdCt_EditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien_nt"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien_nt"]).Editor, true));
                                break;
                            }
                            break;
                        case "gia":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                Decimal num5 = new Decimal(0);
                                num3 = new Decimal(0);
                                Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["gia"].Value, new Decimal(0));
                                Decimal num7 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                if (num7 * num6 != new Decimal(0))
                                    e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(num7 * num6, StartUpTrans.M_ROUND);
                                this.GrdCt_EditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien"]).Editor, true));
                                break;
                            }
                            break;
                        case "tien_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    Decimal num5 = new Decimal(0);
                                    Decimal num6 = SysFunc.Round(this.txtTy_gia.nValue * (e.Editor as NumericTextBox).nValue, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                        e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                    else if (num6 != new Decimal(0))
                                        e.Cell.Record.Cells["tien"].Value = (object)num6;
                                }
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
                        case "tien":
                            if (e.Cell.IsDataChanged)
                            {
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

        private void V_Sua()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 0)
            {
                int num = (int)ExMessageBox.Show(70, StartupBase.SasObj, "Không có dữ liệu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1 || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pc"].ToString().Trim() != "" && ExMessageBox.Show(150, StartupBase.SasObj, "Đã tạo phiếu chi tự động cho phiếu nhập. Có sửa không?", "", MessageBoxButton.YesNo, MessageBoxImage.Asterisk, MessageBoxResult.No) == MessageBoxResult.No)
            {
                return;
            }
            else
            {
                this.txtMa_kh.IsFocus = true;
                FormTrans.currActionTask = ActionTask.Edit;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                FrmAPCTPN1.IsInEditMode.Value = true;
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                refreshTao_pc(1);
            }
        }

        private void ChkSuaTien_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tien.Value = this.ChkSuaTien.IsChecked.Value;
            bool? isChecked = this.ChkSuaTien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.CalculateTyGia();
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
                this.txtMa_kh.IsFocus = true;
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1["stt_rec"] = (object)str;
                row1["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row1["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row1["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row1["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ma_qs"].ToString().Trim());
                }

                if (row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    row1["ty_giaf"] = (object)1;
                }
                else
                {
                    row1["ty_giaf"] = StartUp.GetRates(row1["ma_nt"].ToString().Trim(), Convert.ToDateTime(row1["ngay_ct"]).Date);
                }
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["t_tien_nt"] = (object)0;
                row1["t_tien"] = (object)0;
                row1["t_thue_nt"] = (object)0;
                row1["t_thue"] = (object)0;
                row1["t_tt_nt"] = (object)0;
                row1["t_tt"] = (object)0;
                row1["tao_pc"] = (object)0;
                row1["stt_rec_pc"] = "";
                row1["so_ct_pc"] = "";
                row1["ma_ct_pc"] = "";
                row1["ma_qs_pc"] = "";
                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                row2["stt_rec"] = (object)str;
                row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)1);
                row2["ma_ct"] = (object)StartUpTrans.Ma_ct;
                if (this.txtNgay_ct.Value != null)
                {
                    row2["ngay_ct"] = this.txtNgay_ct.dValue.Date;
                }
                else
                {
                    row2["ngay_ct"] = DateTime.Now.Date;
                }
                row2["tien_nt"] = (object)0;
                row2["tien"] = (object)0;

                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.iOldRow = FrmAPCTPN1.iRow;
                FrmAPCTPN1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmAPCTPN1.IsInEditMode.Value = true;
                this.txtTen_kh.Text = "";
                this.txtTenTK.Text = "";
                this.TabInfo.SelectedIndex = 0;
                refreshTao_pc(2);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_Nhan()
        {
            try
            {
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                bool flag1 = false;
                if (!this.IsSequenceSave)
                {
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (this.GrdCt.Records.Count == 0)
                    {
                        int num = (int)ExMessageBox.Show(75, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                        return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(80, StartupBase.SasObj, "Chưa có mã khách hàng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag1 = true;
                    }
                    else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(85, StartupBase.SasObj, "Chưa vào tài khoản có!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_nx.IsFocus = true;
                        flag1 = true;
                    }
                    else if (this.txtNgay_ct.dValue == new DateTime())
                    {
                        int num = (int)ExMessageBox.Show(90, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                    goto label_18;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_18:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtNgay_ct.Focus();
                            flag1 = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["tk_vt"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(95, StartupBase.SasObj, "Chưa vào tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCt.Focus();
                            flag1 = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(6101, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = true;
                            this.txtMa_qs.IsFocus = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(100, StartupBase.SasObj, "Chưa vào số chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag1 = true;
                        }
                    }
                    if (!flag1)
                    {
                        if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                        {
                            bool flag2 = false;
                            bool flag3 = false;
                            for (int index = 0; index < StartUpTrans.DsTrans.Tables[2].DefaultView.Count; ++index)
                            {
                                DataRowView dataRowView = StartUpTrans.DsTrans.Tables[2].DefaultView[index];
                                if (string.IsNullOrEmpty(dataRowView.Row["ma_ms"].ToString().Trim()) || string.IsNullOrEmpty(dataRowView.Row["so_ct0"].ToString().Trim()))
                                {
                                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(dataRowView.Row);
                                    StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                                }
                                else
                                {
                                    if (!StartUpTrans.M_MST_CHECK.Equals("0") && (!SysFunc.CheckSumMaSoThue(dataRowView.Row["ma_so_thue"].ToString().Trim()) && !string.IsNullOrEmpty(dataRowView.Row["ma_so_thue"].ToString().Trim()) && !flag2))
                                    {
                                        int num = (int)ExMessageBox.Show(105, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                            int num = (int)ExMessageBox.Show(110, StartupBase.SasObj, string.Format("Hoá đơn số [{0}], ký hiệu [{1}], ngày [{2}], MST [{3}] đã tồn tại!", (object)so_ct0, (object)so_seri0, (object)ngay_ct0, (object)ma_so_thue), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag3 = true;
                                            if (StartUpTrans.M_CHK_HD_VAO == 2)
                                                return;
                                        }
                                    }
                                }
                            }
                            if (!this.CheckVoucherOutofDate())
                                flag1 = true;
                        }
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                        {
                            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                            {
                                if (string.IsNullOrEmpty(dataRowView.Row["tk_vt"].ToString().Trim()))
                                {
                                    StartUpTrans.DsTrans.Tables[1].Rows.Remove(dataRowView.Row);
                                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                                }
                            }
                        }
                    }
                }
                if (!flag1)
                {
                    this.UpdateTotalHT();
                    if (!this.IsSequenceSave)
                    {
                        object obj1 = StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj2 = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj3 = StartUpTrans.DsTrans.Tables[2].Compute("sum(t_tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj4 = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        Decimal num1 = Convert.ToDecimal(obj1.Equals((object)DBNull.Value) ? (object)0 : obj1);
                        Decimal num2 = Convert.ToDecimal(obj2.Equals((object)DBNull.Value) ? (object)0 : obj2);
                        Decimal num3 = Convert.ToDecimal(obj3.Equals((object)DBNull.Value) ? (object)0 : obj3);
                        Decimal num4 = Convert.ToDecimal(obj4.Equals((object)DBNull.Value) ? (object)0 : obj4);
                        if (FormTrans.currActionTask == ActionTask.Copy && (num2 != num1 || num3 != num4))
                        {
                            if (this.txtT_thue.nValue != new Decimal(0) && this.txtT_thue_nt.nValue != new Decimal(0) && this.txtT_thue_Nt0.nValue != new Decimal(0))
                            {
                                int num5 = (int)ExMessageBox.Show(120, StartupBase.SasObj, "Tổng tiền/ tiền ngoại tệ khác với tổng tiền/ tiền ngoại tệ trong các hóa đơn giá trị gia tăng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = true;
                            }
                        }
                        else if ((num2 != num1 || num3 != num4) && this.GrdCtgt.Records.Count > 0)
                        {
                            int num6 = (int)ExMessageBox.Show(125, StartupBase.SasObj, "Tổng tiền/ tiền ngoại tệ khác với tổng tiền/ tiền ngoại tệ trong các hóa đơn giá trị gia tăng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                        Decimal nValue = this.txtTy_gia.nValue;
                        if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) && this.GrdCt.Records.Count > 0 && nValue != new Decimal(0) && !this.ChkSuaTien.IsChecked.Value)
                        {
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
                            Decimal num7;
                            if (result1 == new Decimal(0))
                            {
                                Decimal num5 = SysFunc.Round(nValue * num2, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                this.txtT_Tien_Nt0.Value = (object)num5;
                                Decimal result2 = new Decimal(0);
                                Decimal? nullable = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien")));
                                if (nullable.HasValue)
                                    Decimal.TryParse(nullable.ToString(), out result2);
                                (this.GrdCt.Records[0] as DataRecord).Cells["tien"].Value = (object)(Convert.ToDecimal((this.GrdCt.Records[0] as DataRecord).Cells["tien"].Value) + (num5 - result2));
                                num7 = new Decimal(0);
                                this.txtT_tt_Nt0.Value = (object)((this.txtT_thue_Nt0.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_thue_Nt0.nValue)) + num5);
                            }
                            else if (result1 < (Decimal)this.GrdCt.Records.Count)
                            {
                                Decimal num5 = SysFunc.Round(nValue * num2, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                this.txtT_Tien_Nt0.Value = (object)num5;
                                Decimal result2 = new Decimal(0);
                                Decimal? nullable = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien")));
                                if (nullable.HasValue)
                                    Decimal.TryParse(nullable.ToString(), out result2);
                                for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                                {
                                    DataRecord record = this.GrdCt.Records[index] as DataRecord;
                                    Decimal result3 = new Decimal(0);
                                    Decimal result4 = new Decimal(0);
                                    Decimal.TryParse(record.Cells["tien_nt"].Value.ToString(), out result3);
                                    Decimal.TryParse(record.Cells["tien"].Value.ToString(), out result4);
                                    if (!(result3 == new Decimal(0)) || !(result4 != new Decimal(0)))
                                    {
                                        record.Cells["tien"].Value = (object)(result4 + (num5 - result2));
                                        break;
                                    }
                                }
                                num7 = new Decimal(0);
                                this.txtT_tt_Nt0.Value = (object)((this.txtT_thue_Nt0.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_thue_Nt0.nValue)) + num5);
                            }
                        }
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = (object)StartupBase.SasObj.GetOption("M_MA_DVCS").ToString();
                        int index1 = this.Lay_Index_Record_Co_TienThueMax();
                        if (index1 != -1)
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index1]["so_ct0"];
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index1]["ngay_ct0"];
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri0"] = StartUpTrans.DsTrans.Tables[2].DefaultView[index1]["so_seri0"];
                        }
                    }
                    PhanBoThueInCT();
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable1.Rows[0]["status"] = (object)0;
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                    DataTable LocalTable3 = StartUpTrans.DsTrans.Tables[2].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                        {
                            dataRowView["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            dataRowView["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            dataRowView["ma_ct"] = (object)StartUpTrans.Ma_ct;
                        }
                        LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                        {
                            dataRowView["ma_nt"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"];
                            dataRowView["ty_gia"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"];
                            dataRowView["ty_giaf"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_giaf"];
                            dataRowView["status"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
                            dataRowView["ma_gd"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"];
                            dataRowView["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                        }
                        LocalTable3.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(130, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return;
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctgtdbf"].ToString(), LocalTable3, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(135, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return;
                    }
                }
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
                                            if (ExMessageBox.Show(140, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
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
                                            int num = (int)ExMessageBox.Show(145, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.txtSo_ct.SelectAll();
                                            this.txtSo_ct.Focus();
                                            flag1 = true;
                                            break;
                                        }
                                        break;
                                    case "PH02":
                                        int num1 = (int)ExMessageBox.Show(150, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.txtMa_nx.IsFocus = true;
                                        break;
                                    case "CT01":
                                        int int16_1 = (int)Convert.ToInt16(dataRowView[1]);
                                        int num2 = (int)ExMessageBox.Show(155, StartupBase.SasObj, "Tk vật tư là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.tiHT.Focus();
                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_1] as DataRecord).Cells["tk_vt"];
                                        this.GrdCt.Focus();
                                        break;
                                    case "GT01":
                                        int int16_2 = (int)Convert.ToInt16(dataRowView[1]);
                                        int num3 = (int)ExMessageBox.Show(160, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.tabItem3.Focus();
                                        this.GrdCtgt.ActiveCell = (this.GrdCtgt.Records[int16_2] as DataRecord).Cells["tk_thue_no"];
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
                 this.LoadDataDu13();
             }));
                   })).Start();
                    if (!this.IsSequenceSave)
                    {
                        int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                        if (FrmAPCTPN1.iRow != pos)
                        {
                            DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                            DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                            row2.ItemArray = row1.ItemArray;
                            if (FrmAPCTPN1.iRow > pos)
                                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                            else
                                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                            StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                            FrmAPCTPN1.iRow = pos;
                        }
                        FrmAPCTPN1.IsInEditMode.Value = false;
                        FormTrans.currActionTask = ActionTask.View;
                    }
                }
                this.ChkTaoPc.IsEnabled = false;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[APCTPN1-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private int Lay_Index_Record_Co_TienThueMax()
        {
            int num1 = -1;
            Decimal num2 = new Decimal(0);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[2].DefaultView.Count; ++index)
            {
                if (this.Parsedecimal(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["t_thue"], new Decimal(0)) > num2)
                {
                    num2 = Decimal.Parse(StartUpTrans.DsTrans.Tables[2].DefaultView[index]["t_thue"].ToString());
                    num1 = index;
                }
            }
            return num1;
        }

        public Decimal Parsedecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void btnNhan_Click(object sender, RoutedEventArgs e)
        {
            this.V_Nhan();
        }

        private void btnXem_Click(object sender, RoutedEventArgs e)
        {
            this.V_Xem();
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
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, strBrowse, strBrowseCt, "stt_rec");
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.frmBrw.LanguageID = "APCTPN1_4";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string str = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmAPCTPN1.iRow = index + 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            }
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmTim3 frmTim3 = new FrmTim3(StartupBase.SasObj, StartUpTrans.filterId, StartUpTrans.filterView);
                SysFunc.LoadIcon((Window)frmTim3);
                frmTim3.ShowDialog();
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
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
                                Decimal num4 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"].ToString());
                                Decimal num5 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"].ToString());
                                if (this.ParseDecimal((object)e.Cell.Record.Cells["t_tien_nt"].Value.ToString(), new Decimal(0)) == new Decimal(0))
                                    e.Cell.Record.Cells["t_tien_nt"].Value = (object)SysFunc.Round(num4 - result1, StartUpTrans.M_ROUND_NT);
                                if (this.ParseDecimal((object)e.Cell.Record.Cells["t_tien"].Value.ToString(), new Decimal(0)) == new Decimal(0))
                                    e.Cell.Record.Cells["t_tien"].Value = (object)SysFunc.Round(num5 - result2, StartUpTrans.M_ROUND);
                                Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                Decimal num7 = SysFunc.Round(num6 * (num4 - result1) / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                Decimal num8 = SysFunc.Round(num6 * (num5 - result2) / new Decimal(100), StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["t_thue_nt"].Value = (object)num7;
                                e.Cell.Record.Cells["t_thue"].Value = (object)num8;
                                e.Cell.Record.Cells["t_tt_nt"].Value = (object)(num7 + (num4 - result1));
                                e.Cell.Record.Cells["t_tt"].Value = (object)(num8 + (num5 - result2));
                                this.UpdateTotalThue();
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
                                if (!string.IsNullOrEmpty(autoCompleteControl1.RowResult["dia_chi"].ToString()))
                                    e.Cell.Record.Cells["dia_chi"].Value = autoCompleteControl1.RowResult["dia_chi"];
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
                                if (!string.IsNullOrEmpty(autoCompleteControl1.RowResult["ma_so_thue"].ToString()))
                                    e.Cell.Record.Cells["ma_so_thue"].Value = autoCompleteControl1.RowResult["ma_so_thue"];
                                e.Cell.Record.Cells["ma_so_thue_dmkh"].Value = autoCompleteControl1.RowResult["ma_so_thue"];
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
                                this.UpdateTotalThue();
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
                                this.UpdateTotalThue();
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
                                    this.UpdateTotalThue();
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
                                this.UpdateTotalThue();
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
                                this.UpdateTotalThue();
                                break;
                            }
                            break;
                        case "tk_thue_no":
                            if (e.Editor.Value == null || e.Cell.Record.Index != 0)
                                break;
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_no"] = (object)e.Cell.Record.Cells["tk_thue_no"].Value.ToString();
                            break;
                        case "ma_kh2":
                            if (e.Editor.Value == null)
                                break;
                            break;
                        case "ma_thck":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl4.RowResult != null)
                            {
                                e.Cell.Record.Cells["han_tt"].Value = autoCompleteControl4.RowResult["han_tt"];
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
        private void PhanBoThueInCT()
        {
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                return;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);

            Decimal result9 = new Decimal(0);
            Decimal result10 = new Decimal(0);
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            Decimal num4 = new Decimal(0);
            Decimal result11 = new Decimal(0);
            string stt_rec = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result11);
            //if ((this.txtMa_kh_i.Text.Equals("") || !this.txtMa_kh_i.Text.Equals(this.txtMa_kh.Text)) && (this.txttk_i.Text.Equals("") || !this.txttk_i.Text.Equals(this.txtMa_nx.Text)))
            //{
            result1 = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == stt_rec && b.Field<string>("ma_kh").Trim() == this.txtMa_kh.Text.Trim())).Sum<DataRow>((Func<DataRow, Decimal>)(x => x.Field<Decimal>("t_thue_nt")));
            result2 = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == stt_rec && b.Field<string>("ma_kh").Trim() == this.txtMa_kh.Text.Trim())).Sum<DataRow>((Func<DataRow, Decimal>)(x => x.Field<Decimal>("t_thue")));
            //}
            //else
            //{
            //    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result1);
            //    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result2);
            //}
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", "stt_rec= '" + stt_rec + "'").ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", "stt_rec= '" + stt_rec + "'").ToString(), out result4);

            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"].ToString(), out result9);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"].ToString(), out result10);
                Decimal num5;
                Decimal num6;
                if (this.cbMa_nt.Text != StartUpTrans.M_ma_nt0)
                {
                    num5 = !(result9 == new Decimal(0)) ? (result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result9 / result3 * result1, StartUpTrans.M_ROUND_NT)) : new Decimal(0);
                    num6 = !(result10 == new Decimal(0)) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result10 / result4 * result2, StartUpTrans.M_ROUND)) : new Decimal(0);
                }
                else
                {
                    num5 = result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result9 / result3 * result1, StartUpTrans.M_ROUND_NT);
                    num6 = result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result10 / result4 * result2, StartUpTrans.M_ROUND);
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
        private bool GrdCtgt_AddNewRecord(object sender, EditModeEndedEventArgs e)
        {
            return this.NewRowCtGt();
        }

        private void cbMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.cbMa_nt.IsDataChanged)
                return;
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.SetStatusVisibleField();
            if (this.cbMa_nt.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.cbMa_nt.RowResult["loai_tg"];
                if (this.cbMa_nt.RowResult["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    this.txtTy_gia.Value = (object)1;
                else
                    this.txtTy_gia.Value = (object)StartUp.GetRates(this.cbMa_nt.RowResult["ma_nt"].ToString().Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
            }
            this.CalculateTyGia();
        }

        private void SetStatusVisibleField()
        {
            this.ChangeLanguage();
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value || (this.txtMa_kh.RowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim())))
                return;
            if (this.M_LAN.ToUpper().Equals("V"))
            {
                this.txtTen_kh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
                StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ten_kh2"] = (object)this.txtMa_kh.RowResult["ten_kh2"].ToString();
            }
            else
            {
                this.txtTen_kh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
                StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ten_kh"] = (object)this.txtMa_kh.RowResult["ten_kh"].ToString();
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
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"] = (object)this.txtMa_kh.RowResult["ma_thck"].ToString().Trim();
            this.txtMaSoThue.Text = this.txtMa_kh.RowResult["ma_so_thue"].ToString();
            this.txtMa_nx.Text = string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) ? this.txtMa_kh.RowResult["tk"].ToString().Trim() : this.txtMa_nx.Text.Trim();
            this.LoadDataDu13();
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

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            this.GrdLayout20.Visibility = Visibility.Hidden;
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.StartEditMode);
                    if (Keyboard.FocusedElement.GetType().Name.Equals("TextBoxAutoComplete") && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus())
                        return;
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells["tk_vt"];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCt, "tk_vt");
                            break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(190, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                        return;
                    if (this.GrdCt.ActiveRecord is DataRecord activeRecord)
                    {
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
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[activeRecord.Index].Row);
                            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                            if (this.GrdCt.Records.Count > 0)
                                this.GrdCt.ActiveRecord = this.GrdCt.Records[num1 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : num1];
                            this.UpdateTotalHT();
                        }
                        break;
                    }
                    break;
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                ;
        }

        private void GrdCtgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
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

        private void GrdCtgt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.StartEditMode);
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
                case Key.F8:
                    if (ExMessageBox.Show(195, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCtgt.ActiveRecord is DataRecord activeRecord))
                        break;
                    int num1 = 0;
                    Cell activeCell = this.GrdCtgt.ActiveCell;
                    if (activeRecord.Index == 0)
                    {
                        if (this.GrdCtgt.Records.Count == 1)
                            this.GrdCtgt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                        num1 = activeRecord.Index;
                    }
                    else if (activeRecord.Index > 0)
                        num1 = activeRecord.Index - 1;
                    int num2 = this.GrdCtgt.ActiveCell == null ? 0 : this.GrdCtgt.ActiveCell.Field.Index;
                    this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num2 >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[2].Rows.Remove(StartUpTrans.DsTrans.Tables[2].DefaultView[activeRecord.Index].Row);
                        StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                        if (this.GrdCtgt.Records.Count > 0)
                            this.GrdCtgt.ActiveRecord = this.GrdCtgt.Records[num1 > this.GrdCtgt.Records.Count - 1 ? this.GrdCtgt.Records.Count - 1 : num1];
                        this.UpdateTotalThue();
                    }
                    break;
            }
        }

        private bool NewRowCtGt()
        {
            try
            {
                DataRow dataRow = StartUpTrans.DsTrans.Tables[2].NewRow();
                dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                int result1 = 0;
                int result2 = 0;
                if (this.GrdCt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result1);
                }
                if (this.GrdCtgt.Records.Count > 0)
                {
                    string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                        int.TryParse(str.ToString(), out result2);
                }
                int num = (result1 >= result2 ? result1 : result2) + 1;
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["ma_ms"] = (object)StartUp.M_ma_ms;
                dataRow["ngay_ct"] = this.txtNgay_ct.Value;
                dataRow["ten_vt"] = (object)this.txtDien_giai.Text;
                dataRow["t_tien_nt"] = (object)0;
                dataRow["t_tien"] = (object)0;
                dataRow["thue_suat"] = (object)0;
                dataRow["t_thue_nt"] = (object)0;
                dataRow["t_thue"] = (object)0;
                FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[2].DefaultView, dataRow, 2);
                StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow);
                this.UpdateTotalThue();
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
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                Decimal result1 = new Decimal(0);
                Decimal num1 = new Decimal(0);
                Decimal? nullable1 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien_nt")));
                if (nullable1.HasValue)
                    Decimal.TryParse(nullable1.ToString(), out result1);
                this.txtT_Tien.Value = (object)result1;
                this.txtT_Tien_nt.Value = (object)result1;
                Decimal num2 = this.txtT_thue_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_thue_nt.nValue);
                this.txtT_tt.Value = (object)(result1 + num2);
                this.txtT_tt_nt.Value = (object)(result1 + num2);
                if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    Decimal result2 = new Decimal(0);
                    Decimal num3 = new Decimal(0);
                    Decimal? nullable2 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien")));
                    if (nullable2.HasValue)
                        Decimal.TryParse(nullable2.ToString(), out result2);
                    Decimal num4 = this.txtT_thue_Nt0.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_thue_Nt0.nValue);
                    this.txtT_Tien_Nt0.Value = (object)result2;
                    this.txtT_tt_Nt0.Value = (object)(result2 + num4);
                }
                else
                {
                    this.txtT_Tien_Nt0.Value = (object)result1;
                    this.txtT_tt_Nt0.Value = (object)(result1 + num2);
                }
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
                StartUpTrans.DsTrans.Tables[2].AcceptChanges();
                Decimal num1 = new Decimal(0);
                Decimal result1 = new Decimal(0);
                Decimal num2 = this.txtT_Tien_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_Tien_nt.nValue);
                Decimal? nullable1 = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("t_thue_nt")));
                if (nullable1.HasValue)
                    Decimal.TryParse(nullable1.ToString(), out result1);
                this.txtT_thue.Value = (object)result1;
                this.txtT_thue_nt.Value = (object)result1;
                this.txtT_tt.Value = (object)(num2 + result1);
                this.txtT_tt_nt.Value = (object)(num2 + result1);

                if (!this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    Decimal num3 = new Decimal(0);
                    Decimal num4 = this.txtT_Tien_Nt0.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_Tien_Nt0.nValue);
                    Decimal result2 = new Decimal(0);
                    Decimal? nullable2 = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("t_thue")));
                    if (nullable2.HasValue)
                        Decimal.TryParse(nullable2.ToString(), out result2);
                    this.txtT_thue_Nt0.Value = (object)result2;
                    this.txtT_tt_Nt0.Value = (object)(num4 + result2);

                }
                else
                {
                    this.txtT_thue_Nt0.Value = (object)result1;
                    this.txtT_tt_Nt0.Value = (object)(num2 + result1);
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) && !this.txtMa_nx.IsReadOnly && this.txtMa_nx.RowResult != null)
            {
                if (this.M_LAN.ToUpper().Equals("V"))
                {
                    this.txtTenTK.Text = this.txtMa_nx.RowResult["ten_nx"].ToString();
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ten_nx2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
                }
                else
                {
                    this.txtTenTK.Text = this.txtMa_nx.RowResult["ten_nx2"].ToString();
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmAPCTPN1.iRow]["ten_nx"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
                }
                refreshTao_pc(1);
            }
            this.LoadDataDu13();
        }
        private void refreshTao_pc(int type)
        {
            if (type == 1)
            {
                this.ChkTaoPc.IsEnabled = Inlist(this.txtMa_nx.Text.Trim(), StartupBase.SasObj.GetOption("M_TK_TK_VT").ToString().Trim().Split(','));
            }
            else
            {
                this.ChkTaoPc.IsEnabled = false;
            }
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
        private void CreatePC(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[APCTPN1-CREATEPC] @Stt_rec, @Stt_recPC, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct");
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
                int result = StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                Debug.WriteLine("Tao PC: " + result);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
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
        private void FormMain_Closed(object sender, EventArgs e)
        {
            if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                return;
            Application.Current.Shutdown();
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void LoadDataDu13()
        {
            this.txtSoDuKH.Value = (object)ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString());
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()) || !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                return;
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim()) || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"].ToString().Trim()))
            {
                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
            }
            else
                this.txtSo_ct.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"].ToString().Trim();
        }

        private void GrdCtgt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.UpdateTotalThue();
            this.GrdCtgt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
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
                if (!Keyboard.IsKeyDown(Key.Return))
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
            int num = (int)ExMessageBox.Show(200, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
            if (!(this.txtTy_gia.nValue != new Decimal(0)))
                return;
            Decimal num1 = new Decimal(0);
            Decimal nValue = this.txtTy_gia.nValue;
            Decimal result;
            for (int index = 0; index < this.GrdCt.Records.Count; ++index)
            {
                if ((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt"].Value != DBNull.Value)
                {
                    result = new Decimal(0);
                    Decimal.TryParse((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt"].Value.ToString(), out result);
                    if (result * nValue > new Decimal(0))
                        (this.GrdCt.Records[index] as DataRecord).Cells["tien"].Value = (object)SysFunc.Round(nValue * result, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                }
            }
            this.UpdateTotalHT();
            for (int index = 0; index < this.GrdCtgt.Records.Count; ++index)
            {
                if ((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien_nt"].Value != DBNull.Value)
                {
                    result = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien_nt"].Value);
                    if (nValue * result > new Decimal(0))
                    {
                        Decimal num2 = SysFunc.Round(nValue * result, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                        (this.GrdCtgt.Records[index] as DataRecord).Cells["t_tien"].Value = (object)num2;
                        if ((this.GrdCtgt.Records[index] as DataRecord).Cells["thue_suat"].Value != DBNull.Value)
                        {
                            Decimal num3 = Convert.ToDecimal((this.GrdCtgt.Records[index] as DataRecord).Cells["thue_suat"].Value);
                            Decimal num4 = SysFunc.Round(result * num3 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND_NT));
                            (this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue_nt"].Value = (object)num4;
                            (this.GrdCtgt.Records[index] as DataRecord).Cells["t_thue"].Value = (object)SysFunc.Round(num2 * num3 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                        }
                    }
                }
            }
            this.UpdateTotalThue();
        }

        private void GrdCt_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Add || (StartUpTrans.DsTrans.Tables[1].DefaultView.Count != 1 || !(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["dien_giaii"].ToString() == string.Empty)))
                return;
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["dien_giaii"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row["dien_giai"];
        }

        private void txtHan_ck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
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

        private void btnImport_Click(object sender, RoutedEventArgs e)
        {
            if (!FrmAPCTPN1.IsInEditMode.Value)
                return;
            try
            {
                FrmImport frm = new FrmImport();
                frm.ShowDialog();
                if (frm.isOk && frm.dsInvoice != null && frm.dsInvoice.Tables[0].Rows.Count > 0 && frm.dsInvoice.Tables[1].Rows.Count > 0 && frm.dsInvoice.Tables[2].Rows.Count > 0)
                {
                    int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                    for (int index = 0; index < count; ++index)
                        StartUpTrans.DsTrans.Tables[1].DefaultView.Delete(0);
                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    string upper1 = this.cbMa_nt.Text.ToUpper();
                    string upper2 = frm.dsInvoice.Tables[0].DefaultView[0]["ma_nt"].ToString().ToUpper();

                    //Tổng hợp PH
                    DataRow rowph = frm.dsInvoice.Tables[0].DefaultView[0].Row;
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"] = rowph["ma_kh"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = rowph["ten_kh"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = rowph["ten_kh"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = rowph["ma_so_thue"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = rowph["dia_chi"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = rowph["ong_ba"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = rowph["ma_nx"].ToString();

                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = (DateTime)rowph["ngay_lct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"] = (DateTime)rowph["ngay_ct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"] = rowph["ma_nt"].ToString();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"] = rowph["ty_gia"] == DBNull.Value ? 1 : Convert.ToDecimal(rowph["ty_gia"]);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_giaf"] = rowph["ty_giaf"] == DBNull.Value ? 1 : Convert.ToDecimal(rowph["ty_giaf"]);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = rowph["t_tien_nt"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = rowph["t_tien"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = rowph["t_thue_nt"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = rowph["t_thue"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (decimal)rowph["t_thue_nt"] == 0 ? rowph["t_tien_nt"] : rowph["t_tt_nt"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (decimal)rowph["t_thue"] == 0 ? rowph["t_tien"] : rowph["t_tt"];

                    //Tab thue đầu vào
                    int count2 = StartUpTrans.DsTrans.Tables[2].DefaultView.Count;
                    for (int index = 0; index < count2; ++index)
                        StartUpTrans.DsTrans.Tables[2].DefaultView.Delete(0);
                    StartUpTrans.DsTrans.Tables[2].AcceptChanges();

                    DataRow row11 = frm.dsInvoice.Tables[2].DefaultView[0].Row;
                    DataRow row21 = StartUpTrans.DsTrans.Tables[2].NewRow();
                    DataTable table21 = frm.dsInvoice.Tables[2].Clone();
                    table21.Rows.Add(row11.ItemArray);
                    DataTable dataTable21 = StartUpTrans.DsTrans.Tables[2].Clone();
                    dataTable21.Merge(table21, true, MissingSchemaAction.Ignore);
                    if (dataTable21.Rows.Count > 0)
                    {
                        row21.ItemArray = dataTable21.Rows[0].ItemArray;
                        row21["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                        row21["ma_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"];
                        row21["ma_ms"] = StartUp.M_ma_ms;
                        StartUpTrans.DsTrans.Tables[2].Rows.Add(row21);
                    }

                    //Chi tiết CT
                    for (int index = 0; index < frm.dsInvoice.Tables[1].DefaultView.Count; ++index)
                    {
                        DataRow row1 = frm.dsInvoice.Tables[1].DefaultView[index].Row;
                        DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                        DataTable table = frm.dsInvoice.Tables[1].Clone();
                        table.Rows.Add(row1.ItemArray);
                        DataTable dataTable = StartUpTrans.DsTrans.Tables[1].Clone();
                        dataTable.Merge(table, true, MissingSchemaAction.Ignore);
                        if (dataTable.Rows.Count > 0)
                            row2.ItemArray = dataTable.Rows[0].ItemArray;

                        Decimal result1 = new Decimal(0);
                        Decimal.TryParse(row1["so_luong"].ToString(), out result1);
                        if (upper1.Equals(upper2))
                        {
                            if (upper1.Equals(StartUpTrans.M_ma_nt0))
                            {
                                row2["gia_nt"] = row1["gia"];
                                row2["gia"] = row1["gia"];
                                row2["tien_nt"] = row1["tien"];
                                row2["tien"] = row1["tien"];
                            }
                            else
                            {
                                row2["gia_nt"] = row1["gia_nt"];
                                row2["gia"] = (object)SysFunc.Round(Convert.ToDecimal(row1["gia_nt"].ToString()) * this.txtTy_gia.nValue, StartUpTrans.M_ROUND_GIA);
                                row2["tien_nt"] = row1["tien_nt"];
                                row2["tien"] = (object)SysFunc.Round(Convert.ToDecimal(row2["gia"].ToString()) * result1, StartUpTrans.M_ROUND_GIA);
                            }
                        }
                        else if (upper1.Equals(StartUpTrans.M_ma_nt0))
                        {
                            row2["gia_nt"] = row1["gia_nt"];
                            row2["gia0"] = row1["gia_nt"];
                            row2["tien_nt"] = row1["tien_nt"];
                            row2["tien"] = row1["tien_nt"];
                        }
                        else
                        {
                            row2["gia_nt"] = (object)SysFunc.Round(Convert.ToDecimal(row1["gia"]) / this.txtTy_gia.nValue, StartUpTrans.M_ROUND_GIA_NT);
                            row2["gia"] = row1["gia0"];
                            row2["tien_nt"] = (object)SysFunc.Round(Convert.ToDecimal(row2["gia_nt"]) * result1, StartUpTrans.M_ROUND_NT);
                            row2["tien"] = row1["tien"];
                        }
                        row2["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                        int result2 = 0;
                        int result3 = 0;
                        if (this.GrdCt.Records.Count > 0)
                        {
                            string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                            if (str != null)
                                int.TryParse(str.ToString(), out result2);
                        }
                        if (this.GrdCtgt.Records.Count > 0)
                        {
                            string str = StartUpTrans.DsTrans.Tables[2].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                            if (str != null)
                                int.TryParse(str.ToString(), out result3);
                        }
                        int num = (result2 >= result3 ? result2 : result3) + 1;
                        row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), num);
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                    }
                    this.txtMa_kh.Focus();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
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
