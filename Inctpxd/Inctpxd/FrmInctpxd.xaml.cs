using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
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
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace Inctpxd
{
    public partial class FrmInctpxd : FormTrans
    {
        public static int iRow = 0;
        private int iRow_old = 0;
        private FrmCopy _formcopy = (FrmCopy)null;
        private bool txtDiaChiFocusable = true;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject IsCheckedPx_gia_dd;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject M_Ngay_lct;
        private DataSet dsCheckData;
        public DataSet DsVitual;

        public FrmInctpxd()
        {
            this.InitializeComponent();
            this.LanguageProvider.Language = StartUpTrans.M_LAN;
            this.BindingSasObj = StartupBase.SasObj;
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_Ma_nt = this.txtMa_nt;
            this.C_So_ct = this.txtSo_ct;
        }

        private void FrmInctpxd_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmInctpxd.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                this.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource("IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource("Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource("Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource("IsCheckedSua_tien");
                this.IsCheckedPx_gia_dd = (CodeValueBindingObject)this.FormMain.FindResource("IsCheckedPx_gia_dd");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource("Ty_Gia_ValueChange");
                this.M_Ngay_lct = (CodeValueBindingObject)this.FormMain.FindResource("M_Ngay_lct");
                this.M_Ngay_lct.Value = StartUpTrans.M_ngay_lct.Equals("1");
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)this.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                if (StartupBase.SasObj.GetOption(this.stt_mau_temlate.ToString(), "M_DM_VT_CK").ToString().Trim() == "0")
                {
                    Grid parent = (Grid)this.GrdLayout11.Parent;
                    if (parent != null)
                        parent.RowDefinitions[2].Height = new GridLength(0.0);
                    this.GrdLayout11.Visibility = Visibility.Collapsed;
                }
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.IsCheckedPx_gia_dd.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["px_gia_dd"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = false;
                }
                this.Voucher_Lan0.Value = this.M_LAN.Equals("V");
                this.TabInfo.SelectedIndex = 0;
                this.SetFocusToolbar();

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
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout27.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtStatus.IsEnabled = false;
        }

        private void V_Sau()
        {
            if (FrmInctpxd.iRow < StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                ++FrmInctpxd.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void V_Truoc()
        {
            if (FrmInctpxd.iRow > 1)
                --FrmInctpxd.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void V_Dau()
        {
            FrmInctpxd.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count > 1 ? 1 : 0;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void V_Cuoi()
        {
            FrmInctpxd.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
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
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1["stt_rec"] = (object)stt_rec;
                row1["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row1["row_id"] = (object)0;
                DateTime dateTime;
                if (SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)))
                {
                    row1["ngay_ct"] = (object)this.txtNgay_ct.dValue.Date;
                }
                else
                {
                    row1["ngay_ct"] = DateTime.Now.Date;
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                {
                    row1["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    row1["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id);
                }
                else
                {
                    row1["ma_gd"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_gd"];
                    row1["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_qs"].ToString().Trim());
                }
                if (row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    row1["ty_giaf"] = (object)1;
                }
                else
                {
                    DataRow dataRow = row1;
                    string _ma_nt = row1["ma_nt"].ToString().Trim();
                    DateTime date = Convert.ToDateTime(row1["ngay_ct"]).Date;
                    dataRow["ty_giaf"] = StartUp.GetRates(_ma_nt, date); ;
                }
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["sua_tien"] = (object)0;
                row1["px_gia_dd"] = (object)0;
                row1["t_tien_nt"] = (object)0;
                row1["t_tien"] = (object)0;
                row1["t_so_luong"] = (object)0;
                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                row2["stt_rec"] = (object)stt_rec;
                row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)1);
                row2["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row2["ngay_ct"] = row1["ngay_ct"];
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    row2["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"];
                row2["so_luong"] = (object)0;
                row2["gia_nt"] = (object)0;
                row2["tien_nt"] = (object)0;
                row2["gia"] = (object)0;
                row2["tien"] = (object)0;
                row2["sl_dm"] = (object)-1;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                this.iRow_old = FrmInctpxd.iRow;
                FrmInctpxd.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nt"].ToString());
                this.DsVitual = (DataSet)null;
                this.IsInEditMode.Value = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtma_gd.IsFocus = true));
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
                int num = (int)ExMessageBox.Show(870, StartupBase.SasObj, "Không có dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                    return;
                FormTrans.currActionTask = ActionTask.Edit;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.IsInEditMode.Value = true;
                this.txtma_gd.IsFocus = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   this.CheckReadonlyMa_kho();
                   this.txtma_gd.IsFocus = true;
               }));
            }
        }

        private void V_Huy()
        {
            this.IsInEditMode.Value = false;
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
            {
                switch (FormTrans.currActionTask)
                {
                    case ActionTask.Add:
                    case ActionTask.Copy:
                        this.Xoa();
                        FrmInctpxd.iRow = this.iRow_old;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
                        break;
                    case ActionTask.Edit:
                        if (this.DsVitual != null)
                        {
                            string stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString());
                            if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                            {
                                foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + stt_rec + "'"))
                                    StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                            }
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow].ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                            StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                            StartUp.DataFilter(stt_rec);
                            break;
                        }
                        break;
                }
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
            this.TabInfo.SelectedIndex = 0;
            FormTrans.currActionTask = ActionTask.None;
        }

        private void Xoa()
        {
            FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUpTrans.UpdateTkSd13(1, 0);
                string format = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 11 ? string.Format(format, (object)"[DeleteVoucher]") : string.Format(format, (object)StartUpTrans.Process_Store[11]));
                sqlcmd.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)str;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString());
                StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmInctpxd.iRow);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.txtNgay_ct.Text = "";
                    FrmInctpxd.iRow = FrmInctpxd.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmInctpxd.iRow - 1 : FrmInctpxd.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["stt_rec"].ToString());
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            FormTrans.currActionTask = ActionTask.None;
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            this.Xoa();
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void V_Nhan()
        {
            try
            {
                if (!this.IsSequenceSave)
                {
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                }
                if (this.CheckValid())
                {
                    this.Sum_ALL();
                    StartUpTrans.DsTrans.Tables[1].Clone();
                    if (!this.IsSequenceSave)
                    {
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = (object)StartupBase.SasObj.GetOption("M_MA_DVCS").ToString();
                        DateTime dateTime = (DateTime)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                        object obj = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["px_gia_dd"];
                        string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                        if (dateTime >= StartUp.ngay_gia_px && obj.ToString() != "1")
                        {
                            SqlCommand sqlcmd = new SqlCommand("Ingia_px");
                            sqlcmd.CommandType = CommandType.StoredProcedure;
                            sqlcmd.Parameters.Add("@Ngay_ct", SqlDbType.SmallDateTime).Value = (object)dateTime;
                            sqlcmd.Parameters.Add("@Ma_kho", SqlDbType.VarChar).Value = (object)"";
                            sqlcmd.Parameters.Add("@Ma_vt", SqlDbType.VarChar).Value = (object)"";
                            sqlcmd.Parameters.Add("@So_luong", SqlDbType.Decimal).Value = (object)0;
                            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                            {
                                string str2 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_ton"].ToString().Trim();
                                if (str2.Equals("1") || str2.Equals("4"))
                                {
                                    string str3 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString();
                                    string str4 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString();
                                    Decimal num = (Decimal)StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"];
                                    sqlcmd.Parameters["@Ma_kho"].Value = (object)str3;
                                    sqlcmd.Parameters["@Ma_vt"].Value = (object)str4;
                                    sqlcmd.Parameters["@So_luong"].Value = (object)num;
                                    DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                                    if (dataSet != null && dataSet.Tables.Count != 0 && dataSet.Tables[0].Rows.Count != 0)
                                    {
                                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia"] = dataSet.Tables[0].Rows[0]["gia"];
                                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien"] = dataSet.Tables[0].Rows[0]["tien"];
                                        if (StartUpTrans.M_ma_nt0 == str1)
                                        {
                                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"] = dataSet.Tables[0].Rows[0]["gia"];
                                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"] = dataSet.Tables[0].Rows[0]["tien"];
                                        }
                                        else
                                        {
                                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"] = dataSet.Tables[0].Rows[0]["gia_nt"];
                                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"] = dataSet.Tables[0].Rows[0]["tien_nt"];
                                        }
                                    }
                                }
                            }
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                        }
                    }
                   
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[0].Clone();
                    LocalTable2.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable2.Rows[0]["status"] = (object)0;
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable2, "stt_rec;row_id");

                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[1].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                            dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;

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

                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable1, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num1 = (int)ExMessageBox.Show(885, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else
                    {
                        bool flag = false;
                        if (!this.IsSequenceSave && !flag)
                        {
                            this.dsCheckData = StartUp.CheckData();
                            this.dsCheckData.Tables[0].AcceptChanges();
                            if (this.dsCheckData.Tables.Count > 0)
                            {
                                string str = "";
                                foreach (DataRowView dataRowView in this.dsCheckData.Tables[0].DefaultView)
                                {
                                    if (!flag)
                                    {
                                        switch (dataRowView[0].ToString())
                                        {
                                            case "PH01":
                                                if (StartUpTrans.M_trung_so.Equals("1"))
                                                {
                                                    if (ExMessageBox.Show(890, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
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
                                                    int num2 = (int)ExMessageBox.Show(895, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    this.txtSo_ct.SelectAll();
                                                    this.txtSo_ct.Focus();
                                                    flag = true;
                                                    break;
                                                }
                                                break;
                                            case "CT01":
                                                int int16_1 = (int)Convert.ToInt16(dataRowView[1]);
                                                int num3 = (int)ExMessageBox.Show(900, StartupBase.SasObj, "Tk nợ là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = true;
                                                this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_1] as DataRecord).Cells["tk_vt"];
                                                this.GrdCt.Focus();
                                                break;
                                            case "CT02":
                                                int int16_2 = (int)Convert.ToInt16(dataRowView[1]);
                                                int num4 = (int)ExMessageBox.Show(905, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = true;
                                                this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_2] as DataRecord).Cells["ma_nx_i"];
                                                this.GrdCt.Focus();
                                                break;
                                            case "CT03":
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
                                        int num2 = (int)ExMessageBox.Show(910, StartupBase.SasObj, "Có vật tư [" + str + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag = true;
                                    }
                                    else if (StartUp.M_CHK_TON_VT.Equals("1"))
                                    {
                                        int num5 = (int)ExMessageBox.Show(915, StartupBase.SasObj, "Có vật tư [" + str + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    }
                                }
                            }
                        }
                        if (!flag)
                        {
                            string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                            new Thread((ThreadStart)(() =>
                           {
                               this.Post();
                               if (this.IsSequenceSave)
                                   return;
                               this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                 {
                                 if (!StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(_stt_rec1))
                                     return;
                                 this.UpdateTonKho();
                             }));
                           })).Start();
                            if (!this.IsSequenceSave)
                            {
                                int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                if (FrmInctpxd.iRow != pos)
                                {
                                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                    DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                    row2.ItemArray = row1.ItemArray;
                                    if (FrmInctpxd.iRow > pos)
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                    else
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    FrmInctpxd.iRow = pos;
                                }
                                FormTrans.currActionTask = ActionTask.None;
                                this.IsInEditMode.Value = false;
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[INCTPXD-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private bool CheckValid()
        {
            bool flag = true;
            if (!this.IsSequenceSave)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                int result = 0;
                int.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString(), out result);
                if (this.IsInEditMode.Value)
                {
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(920, StartupBase.SasObj, "Chưa vào loại hóa đơn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtma_gd.IsFocus = true;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(925, StartupBase.SasObj, "Chưa vào mã khách hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtMa_kh.IsFocus = true;
                    }
                    if ((this.txtNgay_ct.Value == null || this.txtNgay_ct.Value.ToString() == "") && flag)
                    {
                        int num = (int)ExMessageBox.Show(930, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtNgay_ct.Focus();
                    }
                    if (this.txtNgay_ct.Value.ToString() != "" && flag)
                    {
                        if (!this.txtNgay_ct.IsValueValid && flag)
                        {
                            int num = (int)ExMessageBox.Show(935, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                        if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(Convert.ToDateTime(this.txtNgay_ct.dValue))) && flag)
                        {
                            int num = (int)ExMessageBox.Show(940, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                        if (flag && Convert.ToDateTime(this.txtNgay_ct.dValue) < NgayTC.GetStartDate(StartUp.M_ngay_ct0))
                        {
                            int num = (int)ExMessageBox.Show(945, StartupBase.SasObj, "Ngày hạch toán phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                        int num1;
                        if (flag && StartUp.M_NGAY_BAT_DAU.HasValue)
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
                                    goto label_21;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_21:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                    }
                    if (flag && StartUpTrans.M_ngay_lct.Equals("1"))
                    {
                        if (this.txtNgay_lct.Value == null || this.txtNgay_lct.Value.ToString() == "")
                        {
                            int num = (int)ExMessageBox.Show(950, StartupBase.SasObj, "Chưa vào ngày lập px!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtNgay_lct.Focus();
                            return false;
                        }
                        if (!this.txtNgay_lct.IsValueValid)
                        {
                            int num = (int)ExMessageBox.Show(955, StartupBase.SasObj, "Ngày lập px không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtNgay_lct.Focus();
                            return false;
                        }
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(985, StartupBase.SasObj, "Chưa vào quyển c.từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtMa_qs.IsFocus = true;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(975, StartupBase.SasObj, "Chưa vào số chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                        this.txtSo_ct.Focus();
                    }
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 && flag)
                    {
                        int num = (int)ExMessageBox.Show(990, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                        this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                        this.GrdCt.ActiveCell = (this.GrdCt.Records[0] as DataRecord).Cells["ma_vt"];
                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                    }
                    var source = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().AsEnumerable().GroupBy(o =>
                    {
                        var data = new
                        {
                            mvt = o.Field<string>("ma_vt"),
                            mkho = o.Field<string>("ma_kho_i")
                        };
                        return data;
                    }).Select(g =>
                    {
                        var data = new
                        {
                            ma_vt = g.Key.mvt,
                            ma_kho = g.Key.mkho,
                            so_luong = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("so_luong")))
                        };
                        return data;
                    });
                    DataTable dataTable = StartUpTrans.DsTrans.Tables[1].Clone();
                    if (source.ToArray().Length > 0)
                    {
                        foreach (var data in source)
                        {
                            DataRow row = dataTable.NewRow();
                            row["ma_vt"] = (object)data.ma_vt;
                            row["ma_kho_i"] = (object)data.ma_kho;
                            row["so_luong"] = (object)data.so_luong;
                            dataTable.Rows.Add(row);
                        }
                    }
                    for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count && flag; ++index)
                    {
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_ct"] = (object)StartUpTrans.Ma_ct;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString()))
                        {
                            int num = (int)ExMessageBox.Show(995, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                            return false;
                        }
                        if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString()))
                        {
                            int num = (int)ExMessageBox.Show(1000, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_kho_i"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                        if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_nx_i"].ToString().Trim()))
                        {
                            int num = (int)ExMessageBox.Show(1015, StartupBase.SasObj, "Chưa vào tk nợ, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_nx_i"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                        if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_vt"].ToString().Trim()))
                        {
                            int num = (int)ExMessageBox.Show(1025, StartupBase.SasObj, "Chưa vào tk có, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_vt"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                        if (int.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_ton"].ToString()) == 3 && StartUp.M_SL0_NTXT_CK != "1" && Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString()) == new Decimal(0))
                        {
                            int num = (int)ExMessageBox.Show(1035, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["so_luong"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                            flag = false;
                        }
                    }
                }
            }
            return flag;
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            this._formcopy = new FrmCopy();
            SysFunc.LoadIcon((Window)this._formcopy);
            this.DsVitual = (DataSet)null;
            this._formcopy.Closed += new EventHandler(this._formcopy_Closed);
            this._formcopy.ShowDialog();
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!this._formcopy.isCopy)
                return;
            string stt_rec = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(stt_rec))
            {
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow].ItemArray;
                row1["stt_rec"] = (object)stt_rec;
                row1["ngay_ct"] = (object)this._formcopy.ngay_ct;
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
                this.iRow_old = FrmInctpxd.iRow;
                FrmInctpxd.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUp.DataFilter(stt_rec);
                this.IsInEditMode.Value = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtma_gd.IsFocus = true));
            }
        }

        private void V_Xem()
        {
            FormTrans.currActionTask = ActionTask.View;
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.ListFieldSum = "t_tien_nt;t_tien";
            formView.frmBrw.Title = StartUp.M_Tilte;
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "Inctpxd_5";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null || !formView.DataGrid.ActiveRecord.GetType().Name.Equals("DataRecord"))
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string stt_rec = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmInctpxd.iRow = index + 1;
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmSearchInctpxd frmSearchInctpxd = new FrmSearchInctpxd(StartupBase.SasObj, StartUpTrans.filterId, StartUp.tableList);
                SysFunc.LoadIcon((Window)frmSearchInctpxd);
                frmSearchInctpxd.Closed += new EventHandler(this._FrmTim_Closed);
                frmSearchInctpxd.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void _FrmTim_Closed(object sender, EventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void V_In()
        {
            FrmPrintInctpxd frmPrintInctpxd = new FrmPrintInctpxd()
            {
                DsPrint = StartUpTrans.DsTrans.Copy()
            };
            frmPrintInctpxd.DsPrint.Tables[0].TableName = "TablePH";
            frmPrintInctpxd.DsPrint.Tables[1].TableName = "TableCT";
            frmPrintInctpxd.DsPrint.Tables["TablePH"].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
            frmPrintInctpxd.DsPrint.Tables["TablePH"].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
            {
                DefaultValue = (object)0
            });
            string str = frmPrintInctpxd.DsPrint.Tables["TablePH"].Rows[FrmInctpxd.iRow]["stt_rec"].ToString();
            int int16 = (int)Convert.ToInt16(StartUpTrans.DmctInfo["so_dong_in"]);
            int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
            if (count < int16)
            {
                for (int index = count; index < int16; ++index)
                {
                    DataRow row = frmPrintInctpxd.DsPrint.Tables["TableCT"].NewRow();
                    row["stt_rec"] = (object)str;
                    row["stt_rec0"] = (object)"999";
                    frmPrintInctpxd.DsPrint.Tables["TableCT"].Rows.Add(row);
                }
            }
            frmPrintInctpxd.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            frmPrintInctpxd.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
            frmPrintInctpxd.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
            frmPrintInctpxd.DsPrint.Tables.Add(StartUp.GetDmnt().Copy());
            frmPrintInctpxd.DsPrint.Tables.Add(this.CreateTableInfo().Copy());
            frmPrintInctpxd.ShowDialog();
        }

        private DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "TableInfo";
            dataTable.Columns.Add(new DataColumn("M_PHONE", typeof(string))
            {
                DefaultValue = (object)StartUp.M_PHONE
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (FormTrans.currActionTask != ActionTask.Add)
                this.UpdateTonKho();
            this.IsVisibilityFieldsXamDataGridByMa_NT(ma_nt);
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            this.IsVisibilityFieldsXamDataGridByPx_gia_dd();
        }

        private void IsVisibilityFieldsXamDataGridByMa_NT(string ma_nt)
        {
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
                FieldLength? width = this.GrdCt.FieldLayouts[0].Fields["gia"].Width;
                double num1 = width.Value.Value;
                settings1.CellMaxWidth = num1;
                FieldSettings settings2 = this.GrdCt.FieldLayouts[0].Fields["tien"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["tien"].Width;
                double num2 = width.Value.Value;
                settings2.CellMaxWidth = num2;
            }
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0][nameof(ma_nt)].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0][nameof(ma_nt)].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.ChangeLanguage();
        }

        private void IsVisibilityFieldsXamDataGridBySua_Tien()
        {
            this.IsCheckedSua_tien.Value = this.ChkSua_tien.IsChecked.Value;
        }

        private void IsVisibilityFieldsXamDataGridByPx_gia_dd()
        {
            this.IsCheckedPx_gia_dd.Value = this.ChkPx_gia_dd.IsChecked.Value;
        }

        private void txtma_gd_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsInEditMode.Value || (this.txtma_gd.RowResult == null || string.IsNullOrEmpty(this.txtma_gd.Text.Trim())))
                return;
            this.txtten_gd.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtma_gd.RowResult["ten_gd"].ToString().Trim() : this.txtma_gd.RowResult["ten_gd2"].ToString().Trim();
        }

        private void txtMa_kho_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.IsInEditMode.Value)
            {
                if (this.txtMa_kho.RowResult == null || string.IsNullOrEmpty(this.txtMa_kho.Text.Trim()))
                    return;
                this.txtTen_kho.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kho.RowResult["ten_kho"].ToString().Trim() : this.txtMa_kho.RowResult["ten_kho2"].ToString().Trim();
                for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                {
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString().Trim() == "")
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"] = (object)this.txtMa_kho.Text;
                }
            }
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                this.txtso_luongSP.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsInEditMode.Value || (this.txtMa_kh.RowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()) || !this.txtMa_kh.IsDataChanged))
                return;
            string str1 = this.txtMa_kh.RowResult["ten_kh"].ToString().Trim();
            string str2 = this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim();
            this.txtTen_kh.Text = StartUpTrans.M_LAN.Equals("V") ? str1 : str2;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ten_kh"] = (object)str1;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ten_kh2"] = (object)str2;
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = (object)this.txtMa_kh.RowResult["tk"].ToString().Trim();
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
            if (this.txtNgay_ct.IsFocusWithin || !this.IsInEditMode.Value || !StartUpTrans.M_ngay_lct.Equals("0") && !string.IsNullOrEmpty(this.txtNgay_lct.Text))
                return;
            this.txtNgay_lct.Value = this.txtNgay_ct.Value;
        }

        private void txtNgay_lct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_lct.IsFocusWithin || !this.IsInEditMode.Value || !this.txtNgay_lct.IsValueValid || !(this.txtNgay_ct.Value.ToString() != this.txtNgay_lct.Value.ToString()))
                return;
            int num = (int)ExMessageBox.Show(1045, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
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

        private void txtMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null)
                return;
            this.IsVisibilityFieldsXamDataGridByMa_NT(this.txtMa_nt.Text.Trim());
            if (this.txtMa_nt.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.txtMa_nt.RowResult["loai_tg"];
                if (this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    this.txtTy_gia.Value = (object)1;
                else
                    this.txtTy_gia.Value = (object)StartUp.GetRates(this.txtMa_nt.Text.Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
            }
            this.Ty_gia_ValueChanged(true);
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtTy_gia.Value == DBNull.Value)
                this.txtTy_gia.Value = (object)0;
            if (this.txtTy_gia.OldValue == Convert.ToDecimal(this.txtTy_gia.Value))
                return;
            this.Ty_gia_ValueChanged(false);
        }

        private void Ty_gia_ValueChanged(bool IsMa_ntChanged)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            this.Ty_Gia_ValueChange.Value = this.txtTy_gia.OldValue != Convert.ToDecimal(this.txtTy_gia.Value);
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            int num2 = this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
            if (!(num1 != new Decimal(0)) || num2 != 0 && !IsMa_ntChanged)
                return;
            this.UpdateTotal("gia", "gia_nt", true);
            this.UpdateTotal("tien", "tien_nt", false);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
        }

        private void UpdateTotal(string columnname, string columnname_nt, bool isPrice)
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal num2 = SysFunc.Round(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname_nt], new Decimal(0)) * num1, !isPrice ? StartUpTrans.M_ROUND : StartUpTrans.M_ROUND_GIA);
                if (num2 != new Decimal(0))
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname] = (object)num2;
            }
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void ChkSua_tien_Click(object sender, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            bool? isChecked = this.ChkSua_tien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.UpdateTotalChkSua_tien();
            this.Ty_gia_ValueChanged(false);
        }

        private void UpdateTotalChkSua_tien()
        {
            int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
            if (count <= 0)
                return;
            for (int index = 0; index < count; ++index)
            {
                Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"], new Decimal(0));
                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"], new Decimal(0));
                if (num1 != new Decimal(0) && num2 != new Decimal(0))
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"] = (object)SysFunc.Round(num1 * num2, StartUpTrans.M_ROUND_NT);
            }
            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void GrdCt_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!this.IsInEditMode.Value || (this.GrdCt.ActiveCell == null || StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) != null))
                    return;
                Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
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
                            if (autoCompleteControl1.RowResult["tk_nvl"].ToString().Trim() != "")
                                e.Cell.Record.Cells["ma_nx_i"].Value = autoCompleteControl1.RowResult["tk_nvl"];
                            (e.Cell.Record.DataItem as DataRowView)["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                            if (e.Cell.IsDataChanged)
                            {
                                e.Cell.Record.Cells["gia_ton"].Value = autoCompleteControl1.RowResult["gia_ton"];
                                e.Cell.Record.Cells["vt_ton_kho"].Value = autoCompleteControl1.RowResult["vt_ton_kho"];
                                if (this.ParseInt(autoCompleteControl1.RowResult["vt_ton_kho"], 0) == 0)
                                {
                                    e.Cell.Record.Cells["so_luong"].Value = (object)0;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_so_luong"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                                    e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                    e.Cell.Record.Cells["gia"].Value = (object)0;
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)0;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                                    e.Cell.Record.Cells["tien"].Value = (object)0;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                                }
                                DataRowView dataItem = e.Cell.Record.DataItem as DataRowView;
                                dataItem["sua_tk_vt"] = autoCompleteControl1.RowResult["sua_tk_vt"];
                                dataItem["tk_vt_dmvt"] = (object)autoCompleteControl1.RowResult["tk_vt"].ToString();
                                dataItem["sl_min"] = autoCompleteControl1.RowResult["sl_min"];
                                if (string.IsNullOrEmpty(e.Cell.Record.Cells["tk_vt"].Value.ToString().Trim()) || autoCompleteControl1.RowResult["sua_tk_vt"].ToString().Trim().Equals("0"))
                                    e.Cell.Record.Cells["tk_vt"].Value = (object)autoCompleteControl1.RowResult["tk_vt"].ToString();
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
                                ControlFunction.RefreshSingleBinding((DependencyObject)CellValuePresenter.FromCell(e.Cell.Record.Cells["tk_vt"]), AutoCompleteTextBox.IsReadOnlyProperty);
                                if (this.ParseInt(autoCompleteControl1.RowResult["vt_ton_kho"], 0) == 1)
                                {
                                    if (!string.IsNullOrEmpty(e.Cell.Record.Cells["ma_vt"].Value.ToString()) && !string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kho_i"].Value.ToString()))
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
                                }
                                else
                                    e.Cell.Record.Cells["ton13"].Value = (object)DBNull.Value;


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
                                            dataItem["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                        }
                                        else
                                        {
                                            dataItem["he_so1"] = (decimal)0;
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
                                                dataItem["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                            }
                                            else
                                            {
                                                dataItem["he_so1"] = (decimal)0;
                                            }
                                        }
                                        else
                                        {
                                            if ((autoCompleteControlmavtdvt1.RowResult["hs_qd"] != DBNull.Value))
                                                dataItem["he_so1"] = (decimal)autoCompleteControlmavtdvt1.RowResult["hs_qd"];
                                            else
                                                dataItem["he_so1"] = (decimal)0;
                                        }
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
                        DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                        AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                        if (autoCompleteControl3.RowResult == null)
                            autoCompleteControl3.SearchInit();
                        if (autoCompleteControl3.RowResult != null && (autoCompleteControl3.RowResult["sua_tk_vt"] != DBNull.Value && Convert.ToDecimal(autoCompleteControl3.RowResult["sua_tk_vt"]) == new Decimal(0)))
                        {
                            e.Cell.Record.Cells["tk_vt"].Value = (object)autoCompleteControl3.RowResult["tk_vt"].ToString();
                            dataItem1["tk_vt_dmvt"] = (object)autoCompleteControl3.RowResult["tk_vt"].ToString();
                        }
                        AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl4 != null)
                        {
                            autoCompleteControl4.SearchInit();
                            if (autoCompleteControl4.RowResult != null && (autoCompleteControl4.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl4.RowResult["tk_dl"].ToString().Trim())))
                            {
                                if (autoCompleteControl4.RowResult["tk_dl"].ToString() != "")
                                    e.Cell.Record.Cells["tk_vt"].Value = (object)autoCompleteControl4.RowResult["tk_dl"].ToString();
                                dataItem1["tk_vt_dmvt"] = (object)autoCompleteControl4.RowResult["tk_dl"].ToString();
                            }
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
                        if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                            e.Cell.Record.Cells["so_luong"].Value = (object)0;
                        Decimal num2 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                        AutoCompleteTextBox autoCompleteControl5 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                        if (autoCompleteControl5.RowResult != null && autoCompleteControl5.RowResult["gia_ton"] != DBNull.Value && int.Parse(autoCompleteControl5.RowResult["gia_ton"].ToString()) == 3 && (StartUp.M_SL0_NTXT_CK != "1" && num2 == new Decimal(0)))
                        {
                            int num3 = (int)ExMessageBox.Show(1050, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            break;
                        }
                        if (e.Cell.IsDataChanged)
                        {
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_so_luong"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                            if (num2 == new Decimal(0))
                            {
                                e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                                e.Cell.Record.Cells["gia"].Value = (object)0;
                            }
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                            Decimal num4 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND_GIA);
                            if (num2 != new Decimal(0) && num4 != new Decimal(0))
                                e.Cell.Record.Cells["gia"].Value = (object)num4;
                            Decimal num5 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND_NT);
                            if (num5 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien_nt"].Value = (object)num5;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            }
                            Decimal num6 = SysFunc.Round(num5 * num1, StartUpTrans.M_ROUND);
                            if (num6 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien"].Value = (object)num6;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["gia"].Value = e.Cell.Record.Cells["gia_nt"].Value;
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"];
                            }
                            break;
                        }
                        break;
                    case "gia_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                            Decimal num5 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND_GIA);
                            if (num3 != new Decimal(0) && num5 != new Decimal(0))
                                e.Cell.Record.Cells["gia"].Value = (object)num5;
                            bool? isChecked = this.ChkSua_tien.IsChecked;
                            if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                            {
                                Decimal num6 = SysFunc.Round(num3 * num4, StartUpTrans.M_ROUND_NT);
                                if (num6 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["tien_nt"].Value = (object)num6;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                                }
                                Decimal num7 = SysFunc.Round(num6 * num1, StartUpTrans.M_ROUND);
                                if (num7 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["tien"].Value = (object)num7;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                                }
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["gia"].Value = e.Cell.Record.Cells["gia_nt"].Value;
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"];
                            }
                            break;
                        }
                        break;
                    case "tien_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien_nt"].Value = (object)0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt"].Value, new Decimal(0));
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                            if (num5 == new Decimal(0))
                            {
                                if (num4 != new Decimal(0))
                                    num5 = SysFunc.Round(num3 / num4, StartUpTrans.M_ROUND_GIA);
                                e.Cell.Record.Cells["gia_nt"].Value = (object)num5;
                            }
                            Decimal num6 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND);
                            if (num6 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien"].Value = (object)num6;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"];
                            }
                            break;
                        }
                        break;
                    case "tien":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien"].Value = (object)0;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            break;
                        }
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
            if (string.IsNullOrEmpty(this.txtMa_sp.Text.Trim()))
                return this.NewRowCt();
            this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
            return false;
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
                    dataRow["ma_kho_i"] = (this.GrdCt.Records[this.GrdCt.Records.Count - 1] as DataRecord).Cells["ma_kho_i"].Value;
                }
                dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)result);
                dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                dataRow["so_luong"] = (object)0;
                dataRow["gia_nt"] = (object)0;
                dataRow["tien_nt"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["tien"] = (object)0;
                dataRow["sl_dm"] = (object)-1;
                dataRow["ton13"] = (object)DBNull.Value;
                dataRow["ma_kho_i"] = (object)this.txtMa_kho.Text;
                int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                if (count > 0)
                {
                    if (this.txtMa_kho.Text.Trim() == "")
                        dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[count - 1].Row["ma_kho_i"];
                    dataRow["ma_nx_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[count - 1].Row["ma_nx_i"];
                    SqlCommand sqlCommand = new SqlCommand("select ma_dm from dmctct where ma_ct=@ma_ct and ma_dm=@ma_dm");
                    sqlCommand.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                    sqlCommand.Parameters.Add("@ma_dm", SqlDbType.Char).Value = (object)"dmvv";
                }
                else
                    dataRow["ma_nx_i"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nx"];
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
            this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && Keyboard.Modifiers == ModifierKeys.Control && string.IsNullOrEmpty(this.txtMa_sp.Text.Trim()))
            {
                this.NewRowCt();
                this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || Keyboard.Modifiers != ModifierKeys.Control)
                return;
            this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value || !string.IsNullOrEmpty(this.txtMa_sp.Text.Trim()))
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
                        CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell((this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"]);
                        if (cellValuePresenter != null && cellValuePresenter.Editor is ControlHostEditor editor)
                        {
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(editor);
                            if (string.IsNullOrEmpty(autoCompleteControl.Text.Trim()))
                            {
                                int num1 = (int)ExMessageBox.Show(1055, StartupBase.SasObj, "Chưa nhập mã vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else if (autoCompleteControl != null)
                            {
                                if (autoCompleteControl.CheckLostFocus())
                                {
                                    string ma_vt = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"].Value.ToString();
                                    string ten_vt = !StartUpTrans.M_LAN.Equals("V") ? ((this.GrdCt.ActiveRecord as DataRecord).DataItem as DataRowView)["ten_vt2"].ToString() : ((this.GrdCt.ActiveRecord as DataRecord).DataItem as DataRowView)["ten_vt"].ToString();
                                    string ma_kho = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_kho_i"].Value.ToString();
                                    object ngay_ct = StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ngay_ct"];
                                    DataTable pn = StartUp.GetPN(ma_vt, ma_kho, ngay_ct);
                                    if (pn.Rows.Count > 0)
                                    {
                                        FrmInctpxd_Pn frmInctpxdPn = new FrmInctpxd_Pn(pn, ten_vt);
                                        frmInctpxdPn.ShowDialog();
                                        int index = this.GrdCt.ActiveRecord.Index;
                                        if (index >= 0 && index < this.GrdCt.Records.Count)
                                        {
                                            DataRowView drvFrmInctpxdPn = frmInctpxdPn.drvFrmINCTPXD_PN;
                                            if (drvFrmInctpxdPn != null)
                                            {
                                                if (StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0))
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmInctpxdPn["gia"];
                                                else
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmInctpxdPn["gia_nt"];
                                                (this.GrdCt.DataSource as DataView)[index]["gia"] = drvFrmInctpxdPn["gia"];
                                                if (this.ParseInt((this.GrdCt.DataSource as DataView)[index]["gia_ton"], 0) == 1 || this.ParseInt((this.GrdCt.DataSource as DataView)[index]["gia_ton"], 0) == 4)
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["px_gia_dd"] = (object)1;
                                                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                                                this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
                                                Decimal num3 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["so_luong"], new Decimal(0));
                                                Decimal num4 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["gia_nt"], new Decimal(0));
                                                this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["gia"], new Decimal(0));
                                                Decimal num5 = SysFunc.Round(num3 * num4, StartUpTrans.M_ROUND_NT);
                                                if (num5 != new Decimal(0))
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["tien_nt"] = (object)num5;
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                                                }
                                                Decimal num6 = SysFunc.Round(num5 * num2, StartUpTrans.M_ROUND);
                                                if (num6 != new Decimal(0))
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["tien"] = (object)num6;
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int num7 = (int)ExMessageBox.Show(1060, StartupBase.SasObj, "Không có phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    }
                                }
                                else
                                {
                                    int num8 = (int)ExMessageBox.Show(1065, StartupBase.SasObj, "Không có phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                            }
                        }
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(1070, StartupBase.SasObj, "Có xoá dòng ghi hiện thời?", "Xoá dòng", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord2))
                        break;
                    int num = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                    int index1 = activeRecord2.Index;
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                    if (num >= 0)
                    {
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[index1].Row);
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_so_luong"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                        StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien_nt"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                        StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                        if (index1 == 0 && this.GrdCt.Records.Count == 0)
                            this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                        if (this.GrdCt.Records.Count > 0)
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[index1 > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : index1];
                    }
                    break;
            }
        }

        private Decimal SumFunction(DataTable datatable, string columnname, int km)
        {
            Decimal result = new Decimal(0);
            Decimal? nullable = datatable.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>(columnname)));
            if (nullable.HasValue)
                Decimal.TryParse(nullable.ToString(), out result);
            return result;
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void ChkPx_gia_dd_Click(object sender, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridByPx_gia_dd();
        }

        private void txtt_tien_nt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            this.UpdateMoney_NT();
        }

        private void UpdateMoney_NT()
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
            Decimal num2 = SysFunc.Round(this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"], new Decimal(0)) * num1, StartUpTrans.M_ROUND);
            if (!(num2 != new Decimal(0)))
                return;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["t_tien"] = (object)num2;
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
                        sqlcmd.Parameters.Add("@Ma_kho", SqlDbType.VarChar).Value = (object)str4;
                        if (this.ParseInt(StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows[0][0], 0) == 0)
                            str6 = "";
                        DataRow[] dataRowArray = StartUp.M_QL_LO_CK.ToString() != "1" ? listTon13.Select("ma_kho LIKE '" + str4 + "' AND ma_vt LIKE '" + str5 + "' AND ma_vv LIKE '" + str6 + "'") : listTon13.Select("ma_kho LIKE '" + str4 + "' AND ma_vt LIKE '" + str5 + "' AND ma_vv LIKE '" + str6 + "' AND ma_lo LIKE '" + str7 + "'");
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = dataRowArray.Length <= 0 ? (object)0 : dataRowArray[0]["ton13"];
                    }
                }
            }
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        private void FormMain_Closed(object sender, EventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.None || this.IsInEditMode == null || !this.IsInEditMode.Value)
                return;
            this.V_Huy();
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            if (menuItemName.Equals("btnSave"))
                return;
            this.UpdateTonKho();
        }

        private void btnDM_Click(object sender, RoutedEventArgs e)
        {
            if (!this.IsEditMode)
                return;
            codmnvlLoc codmnvlLoc = new codmnvlLoc();
            codmnvlLoc.ShowDialog();
            if (codmnvlLoc.isClose)
                this.GrdCt.Focus();
            else
                this.txtso_luongSP.Focus();
            this.CheckReadonlyMa_kho();
        }

        private void txtso_luongSP_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!this.IsEditMode || this.txtso_luongSP.IsFocusWithin)
                return;
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal result = new Decimal(0);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["sl_dm"].ToString(), out result);
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"] = (object)SysFunc.Round(result * this.txtso_luongSP.nValue, StartUp.M_ROUND_SL);
                Cell cell = (this.GrdCt.Records[0] as DataRecord).Cells["so_luong"];
                this.GrdCt_PreviewEditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(cell, CellValuePresenter.FromCell(cell).Editor, true));
            }
        }

        private void Sum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter);
        }

        private void CheckReadonlyMa_kho()
        {
            if (StartUpTrans.DsTrans.Tables[1].DefaultView.Cast<DataRowView>().Any<DataRowView>((Func<DataRowView, bool>)(x => x["ma_kho_i"].ToString().Trim() == "")))
            {
                this.txtMa_kho.IsReadOnly = false;
                this.txtMa_kho.IsTabStop = true;
            }
            else
            {
                this.txtMa_kho.IsReadOnly = true;
                this.txtMa_kho.IsTabStop = false;
            }
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
    }
}
