using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
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
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace SODMHDB
{
    public partial class FrmPoctpna : FormTrans
    {
        public static int iRow = 0;
        public static int OldiRow = 0;
        public static string[] FieldCk = new string[8]
        {
      "tk_ck",
      "tl_ck",
      "ck",
      "ck_nt",
      "t_ck_nt",
      "t_ck",
      "t_tien_sau_ck_nt",
      "t_tien_sau_ck"
        };

        public static string[] FieldKm = new string[12]
        {
      "khuyen_mai",
          "ma_ctkm",
          "ma_nh_km",
          "pt_km",
          "tien_km",
          "tien_km_nt",
          "chiet_khau",
          "chiet_khau_nt",
          "chi_phi",
          "chi_phi_nt",
          "tyle_chietkhau",
          "so_luong_km"
        };

        public static string[] FieldGiaBan = new string[12]
       {
      "gia2",
      "gia_nt2",
      "tien2",
      "tien_nt2",
      "t_tien2",
      "t_tien_nt2",
      "t_tt_nt",
      "t_tt",
      "ma_thue",
      "thue_suat",
      "thue",
      "thue_nt"
       };

        public string Old_ma_kho = string.Empty;
        private bool txtDiaChiFocusable = true;
        private string ma_hd;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private DataSet DsVitual;
        private DataSet dsCheckData;
        int newstt_rec0 = 0;
        private string ct81_ma_vt;
        DataRow[] rkhuyenmaihth;
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
                StartUp.M_AR_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_CK"));
                StartUp.M_KM_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_KM_CK"));
                StartUp.M_AR_TT = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_TT"));
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtso_du_kh.Visibility = this.tblso_du_kh.Visibility = Visibility.Hidden;
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmPoctpna.IsInEditMode,
                    Mode = BindingMode.OneWay
                });
                this.M_LAN = StartUpTrans.M_LAN;
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                if (StartUp.M_BP_BH == "1")
                    this.txtMa_bp.IsTabStop = true;
                this.LanguageProvider.Language = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                }
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    this.Old_ma_kho = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"].ToString();
                if (StartUp.M_AR_CK == 0)
                {
                    this.txtHan_tt.IsReadOnly = true;
                    this.GrdLayout20.RowDefinitions[1].Height = new GridLength(0.0);
                    this.GrdLayout20.RowDefinitions[2].Height = new GridLength(0.0);
                    this.GrdLayout21.RowDefinitions[1].Height = new GridLength(0.0);
                    this.GrdLayout21.RowDefinitions[2].Height = new GridLength(0.0);
                    this.GrdLayout00.RowDefinitions[3].Height = new GridLength(114.0);
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmPoctpna.FieldCk).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
                }
                if (StartUp.M_KM_CK == 0)
                {
                    this.ChkAutoKM.Visibility = Visibility.Hidden;
                    this.tabItemNote1.Visibility = Visibility.Hidden;
                    this.GrdLayoutkm.Visibility = Visibility.Hidden;
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmPoctpna.FieldKm).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
                }
                if (StartUp.M_AR_TT == 0)
                {
                    this.GrdLayout22.RowDefinitions[0].Height = new GridLength(0.0);
                    this.GrdLayout22.RowDefinitions[1].Height = new GridLength(0.0);
                }
                
                //Gia ban
                if(this.stt_mau_temlate == 426)
                {
                    this.ChkSuaTien.Visibility = Visibility.Hidden;
                    this.gbTotal.Visibility = Visibility.Hidden;
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmPoctpna.FieldGiaBan).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
                }    
                this.SetFocusToolbar();
                this.ReSum_ALL();
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
            this.GrdLayout20.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout21.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout22.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayoutkm.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;

            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;

            this.GrdCtKM.DataSource = StartUp.CTKMTable.DefaultView;
            this.GrdVTKM.DataSource = StartUp.VattuKMHTHTable.DefaultView;
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                this.ChkAutoKM.IsChecked = true;
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
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                this.ChkAutoKM.IsChecked = true;
            else
                this.ChkAutoKM.IsChecked = false;
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
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                this.ChkAutoKM.IsChecked = true;
            else
                this.ChkAutoKM.IsChecked = false;
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
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                this.ChkAutoKM.IsChecked = true;
            else
                this.ChkAutoKM.IsChecked = false;
        }

        private void V_Cuoi()
        {
            FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                this.ChkAutoKM.IsChecked = true;
            else
                this.ChkAutoKM.IsChecked = false;
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
                row["ngay_ct"] = !SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct.dValue)) ? (object)DateTime.Now.Date : (object)this.txtNgay_ct.dValue.Date;
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
                row["t_ck_nt"] = (object)0;
                row["t_ck"] = (object)0;
                row["t_tien"] = (object)0;
                row["t_tien_nt"] = (object)0;
                row["t_tien2"] = (object)0;
                row["t_tien_nt2"] = (object)0;
                row["t_thue_nt"] = (object)0;
                row["t_thue"] = (object)0;
                row["t_tt_nt"] = (object)0;
                row["t_tt"] = (object)0;
                row["t_so_luong"] = (object)0;
                row["t_sau_ck_nt"] = (object)0;
                row["t_sau_ck"] = (object)0;
                row["tien_km"] = (object)0;
                row["tien_km_nt"] = (object)0;
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
                this.ChkAutoKM.IsChecked = false;
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
            SqlCommand sqlcmd = new SqlCommand("select so_ct from ct00 where ma_hd = @ma_hd");
            sqlcmd.Parameters.Add("@ma_hd", SqlDbType.Char).Value = (object)this.ma_hd;
            if (StartupBase.SasObj.ExcuteScalar(sqlcmd) != null)
                return true;
            sqlcmd.CommandText = "select so_ct from ct70 where ma_hd = @ma_hd";
            if (StartupBase.SasObj.ExcuteScalar(sqlcmd) != null)
                return true;
            sqlcmd.CommandText = "select so_ct from cttt20 where ma_hd = @ma_hd";
            return StartupBase.SasObj.ExcuteScalar(sqlcmd) != null;
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
            StartUp.CTKMTable.Rows.Clear();
            StartUp.VattuKMHTHTable.Rows.Clear();
            this.GrdCtKM.DataSource = StartUp.CTKMTable.DefaultView;
            this.GrdVTKM.DataSource = StartUp.VattuKMHTHTable.DefaultView;
            this.txtSobaogia.Text = string.Empty;
        }

        protected override bool CheckCanDelete()
        {
            this.ma_hd = this.txtSo_ct.Text.Trim();
            if (string.IsNullOrEmpty(this.ma_hd) || !this.KiemTraCoPhatSinh())
                return true;
            int num = (int)ExMessageBox.Show(2370, StartupBase.SasObj, "Đơn hàng đã có phát sinh, không thể xóa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
            if (StartUp.M_KM_CK == 0 && !StartUp.HiddenFieldIsSetted)
            {
                StartUp.stringBrowse1 = StartUp.EditKmFields(StartUp.stringBrowse1);
                StartUp.stringBrowse2 = StartUp.EditKmFields(StartUp.stringBrowse2);
            }
            if (!StartUp.HiddenFieldIsSetted)
            {
                switch(this.stt_mau_temlate)
                {
                    case 426:
                        StartUp.stringBrowse1 = StartUp.EditGiaBanFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditGiaBanFields(StartUp.stringBrowse2);
                        break;
                }
            }
            StartUp.HiddenFieldIsSetted = true;
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "SODMHDB_4";
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
                dataRow["gia_nt2"] = (object)0;
                dataRow["tien_nt2"] = (object)0;
                dataRow["gia_nt"] = (object)0;
                dataRow["tien_nt"] = (object)0;
                dataRow["tien2"] = (object)0;
                dataRow["gia2"] = (object)0;
                dataRow["tien"] = (object)0;
                dataRow["gia"] = (object)0;
                dataRow["ck_nt"] = (object)0;
                dataRow["ck"] = (object)0;
                dataRow["ton13"] = (object)0;
                dataRow["tien_km"] = (object)0;
                dataRow["tien_km_nt"] = (object)0;
                dataRow["khuyen_mai"] = (object)false;
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
                string nh_kh3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["nh_kh3"].ToString();
                Decimal num1;
                Decimal num2;
                Decimal num3;
                Decimal num4;
                Decimal num5;
                Decimal num6;
                Decimal num7;
                Decimal num8;
                Decimal num9;
                Decimal num10;
                Decimal num11;
                Decimal num12;
                if (this.GrdCt.ActiveCell != null && StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) == null)
                {
                    switch (e.Cell.Field.Name)
                    {
                        case "ma_vt":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null)
                            {
                                DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                                CellCollection cells = e.Cell.Record.Cells;
                                e.Cell.Record.Cells["ten_vt"].Value = autoCompleteControl1.RowResult["ten_vt"];
                                e.Cell.Record.Cells["ten_vt2"].Value = autoCompleteControl1.RowResult["ten_vt2"];
                                e.Cell.Record.Cells["dvt"].Value = autoCompleteControl1.RowResult["dvt"];
                                if (string.IsNullOrEmpty((e.Cell.Record.DataItem as DataRowView)["tk_vt"].ToString().Trim()))
                                    (e.Cell.Record.DataItem as DataRowView)["tk_vt"] = autoCompleteControl1.RowResult["tk_vt"];
                                ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                                if (this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0)) == new Decimal(0) && this.ParseDecimal(e.Cell.Record.Cells["gia2"].Value, new Decimal(0)) == new Decimal(0) && this.txtNgay_ct.dValue != new DateTime())
                                {
                                    DataRow dataRow = StartUpTrans.Getdmgia2(e.Editor.Value.ToString(), string.Format("{0:yyyyMMdd}", (object)this.txtNgay_ct.dValue), nh_kh3);
                                    if (dataRow != null)
                                    {
                                        if (this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                            e.Cell.Record.Cells["gia_nt2"].Value = dataRow["gia2"];
                                        else
                                            e.Cell.Record.Cells["gia_nt2"].Value = dataRow["gia_nt2"];
                                        e.Cell.Record.Cells["gia2"].Value = dataRow["gia2"];
                                    }
                                }

                                //Dvt1
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
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl2 != null)
                            {
                                autoCompleteControl2.SearchInit();
                                if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                                    (e.Cell.Record.DataItem as DataRowView)["tk_vt"] = autoCompleteControl2.RowResult["tk_dl"];
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
                                    Decimal num13 = new Decimal(0);
                                    Decimal result1 = new Decimal(0);
                                    Decimal result2 = new Decimal(0);
                                    Decimal result3 = new Decimal(0);
                                    num1 = new Decimal(0);
                                    num2 = new Decimal(0);
                                    num3 = new Decimal(0);
                                    num4 = new Decimal(0);
                                    Decimal result4 = new Decimal(0);
                                    num5 = new Decimal(0);
                                    Decimal result5 = new Decimal(0);
                                    num6 = new Decimal(0);
                                    Decimal nValue = (e.Editor as NumericTextBox).nValue;
                                    Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result3);
                                    Decimal.TryParse(e.Cell.Record.Cells["gia_nt2"].Value.ToString(), out result1);
                                    Decimal.TryParse(e.Cell.Record.Cells["gia2"].Value.ToString(), out result2);
                                    Decimal.TryParse(e.Cell.Record.Cells["gia_nt"].Value.ToString(), out result5);
                                    Decimal.TryParse(e.Cell.Record.Cells["gia"].Value.ToString(), out result4);
                                    Decimal result6;
                                    Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result6);
                                    if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    {
                                        if (result1 * nValue != new Decimal(0))
                                        {
                                            Decimal num14 = SysFunc.Round(result1 * nValue, StartUpTrans.M_ROUND_NT);
                                            Decimal num15 = SysFunc.Round(result5 * nValue, StartUpTrans.M_ROUND_NT);
                                            Decimal num16 = num14;
                                            Decimal num17 = num15;
                                            e.Cell.Record.Cells["tien_nt2"].Value = (object)num14;
                                            e.Cell.Record.Cells["tien2"].Value = (object)num16;
                                            e.Cell.Record.Cells["tien_nt"].Value = (object)num15;
                                            e.Cell.Record.Cells["tien"].Value = (object)num17;
                                            Decimal num18 = SysFunc.Round(num14 * result6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                            Decimal num19 = num18;
                                            e.Cell.Record.Cells["ck_nt"].Value = (object)num18;
                                            e.Cell.Record.Cells["ck"].Value = (object)num19;
                                            e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num14 - num18) * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                            e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num14 - num18) * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        }
                                    }
                                    else
                                    {
                                        if (result1 * nValue != new Decimal(0))
                                        {
                                            Decimal num14 = SysFunc.Round(result1 * nValue, StartUpTrans.M_ROUND_NT);
                                            e.Cell.Record.Cells["tien_nt2"].Value = (object)num14;
                                            Decimal num15 = SysFunc.Round(result5 * nValue, StartUpTrans.M_ROUND_NT);
                                            e.Cell.Record.Cells["tien_nt"].Value = (object)num15;
                                            Decimal num16 = SysFunc.Round(num14 * result6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                            e.Cell.Record.Cells["ck_nt"].Value = (object)num16;
                                            e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num14 - num16) * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        }
                                        if (result2 * nValue != new Decimal(0))
                                        {
                                            Decimal num14 = SysFunc.Round(result2 * nValue, StartUpTrans.M_ROUND);
                                            e.Cell.Record.Cells["tien2"].Value = (object)num14;
                                            Decimal num15 = SysFunc.Round(num14 * result6 / new Decimal(100), StartUpTrans.M_ROUND);
                                            e.Cell.Record.Cells["ck"].Value = (object)num15;
                                            Decimal num16 = SysFunc.Round(result4 * nValue, StartUpTrans.M_ROUND);
                                            e.Cell.Record.Cells["tien"].Value = (object)num16;
                                            e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num14 - num15) * result3 / new Decimal(100), StartUpTrans.M_ROUND);
                                        }
                                    }
                                    this.Sum_ALL();
                                    Tinhkhuyenmai();
                                    break;
                                }
                                break;
                            }
                            catch (Exception ex)
                            {
                                ErrorLog.CatchMessage(ex);
                                break;
                            }
                        case "gia_nt2":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["gia_nt2"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result1 = new Decimal(0);
                                num7 = new Decimal(0);
                                Decimal num13 = new Decimal(0);
                                num2 = new Decimal(0);
                                num8 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                num9 = new Decimal(0);
                                num5 = new Decimal(0);
                                num10 = new Decimal(0);
                                num6 = new Decimal(0);
                                Decimal result3 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                                Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result1);
                                Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result3);
                                Decimal nValue2 = this.txtTy_gia.nValue;
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    if (nValue1 * result1 != new Decimal(0))
                                    {
                                        bool? isChecked = this.ChkSuaTien.IsChecked;
                                        if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                                        {
                                            num13 = SysFunc.Round(result1 * nValue1, StartUpTrans.M_ROUND_NT);
                                            e.Cell.Record.Cells["tien_nt2"].Value = (object)num13;
                                        }
                                        Decimal num14 = SysFunc.Round(num13 * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["ck_nt"].Value = (object)num14;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num13 - num14) * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["gia2"].Value = (object)nValue1;
                                        e.Cell.Record.Cells["tien2"].Value = (object)num13;
                                        e.Cell.Record.Cells["ck"].Value = (object)num14;
                                        e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num13 - num14) * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    }
                                    e.Cell.Record.Cells["gia2"].Value = (object)nValue1;
                                }
                                else
                                {
                                    if (nValue1 * result1 != new Decimal(0))
                                    {
                                        num13 = SysFunc.Round(result1 * nValue1, StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["tien_nt2"].Value = (object)num13;
                                        Decimal num14 = SysFunc.Round(num13 * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["ck_nt"].Value = (object)num14;
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num13 - num14) * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    }
                                    if (nValue1 * nValue2 != new Decimal(0))
                                        e.Cell.Record.Cells["gia2"].Value = (object)SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND_GIA);
                                    if (num13 * nValue2 != new Decimal(0))
                                    {
                                        Decimal num14 = SysFunc.Round(num13 * nValue2, StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["tien2"].Value = (object)num14;
                                        Decimal num15 = SysFunc.Round(result3 * num14 / new Decimal(100), StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["ck"].Value = (object)num15;
                                        e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num14 - num15) * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    }
                                }
                                this.Sum_ALL();
                                Tinhkhuyenmai();
                                break;
                            }
                            break;
                        case "gia_nt":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["gia_nt"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result = new Decimal(0);
                                num7 = new Decimal(0);
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                num8 = new Decimal(0);
                                num11 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                num9 = new Decimal(0);
                                num5 = new Decimal(0);
                                num10 = new Decimal(0);
                                num6 = new Decimal(0);
                                num12 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result);
                                Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                Decimal nValue2 = this.txtTy_gia.nValue;
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    if (nValue1 * result != new Decimal(0))
                                    {
                                        Decimal num13 = SysFunc.Round(result * nValue1, StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["tien_nt"].Value = (object)num13;
                                        e.Cell.Record.Cells["gia"].Value = (object)nValue1;
                                        e.Cell.Record.Cells["tien"].Value = (object)num13;
                                    }
                                }
                                else
                                {
                                    if (nValue1 * result != new Decimal(0))
                                    {
                                        Decimal num13 = SysFunc.Round(result * nValue1, StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["tien_nt"].Value = (object)num13;
                                    }
                                    if (nValue1 * nValue2 != new Decimal(0))
                                    {
                                        Decimal num13 = SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND_GIA);
                                        e.Cell.Record.Cells["gia"].Value = (object)num13;
                                        e.Cell.Record.Cells["tien"].Value = (object)SysFunc.Round(num13 * result, StartUpTrans.M_ROUND);
                                    }
                                }
                                this.Sum_ALL();
                                Tinhkhuyenmai();
                                break;
                            }
                            break;
                        case "tien_nt2":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["tien_nt2"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                num8 = new Decimal(0);
                                num1 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal result1 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                Decimal nValue2 = this.txtTy_gia.nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result1);
                                Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result2);
                                Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                                if (num14 == new Decimal(0))
                                {
                                    if (num13 != new Decimal(0))
                                        num14 = SysFunc.Round(nValue1 / num13, StartUpTrans.M_ROUND_GIA);
                                    e.Cell.Record.Cells["gia_nt2"].Value = (object)num14;
                                }
                                if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    Decimal num15 = SysFunc.Round(nValue1 * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["ck_nt"].Value = (object)num15;
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((nValue1 - num15) * result1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["tien2"].Value = (object)nValue1;
                                    e.Cell.Record.Cells["ck"].Value = (object)num15;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((nValue1 - num15) * result1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                }
                                else if (nValue1 * nValue2 != new Decimal(0))
                                {
                                    Decimal num15 = SysFunc.Round(nValue1 * result2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["ck_nt"].Value = (object)num15;
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((nValue1 - num15) * result1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                    Decimal num16 = SysFunc.Round(nValue1 * nValue2, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien2"].Value = (object)num16;
                                    Decimal num17 = SysFunc.Round(num16 * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["ck"].Value = (object)num17;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num16 - num17) * result1 / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "gia2":
                            if (e.Editor.Value == DBNull.Value)
                                e.Cell.Record.Cells["gia2"].Value = (object)0;
                            if (e.Cell.IsDataChanged)
                            {
                                Decimal result1 = new Decimal(0);
                                Decimal num13 = new Decimal(0);
                                num2 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                Decimal result3 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result2);
                                Decimal nValue = (e.Editor as NumericTextBox).nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["so_luong"].Value.ToString(), out result1);
                                Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result3);
                                if (nValue * result1 != new Decimal(0))
                                {
                                    Decimal num14 = SysFunc.Round(nValue * result1, StartUpTrans.M_ROUND);
                                    Decimal num15 = SysFunc.Round(num14 * result3 / new Decimal(100), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien2"].Value = (object)num14;
                                    e.Cell.Record.Cells["ck"].Value = (object)num15;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num14 - num15) * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "tien2":
                            if (e.Cell.IsDataChanged)
                            {
                                if (!this.IsCheckedSua_tien.Value)
                                {
                                    num2 = new Decimal(0);
                                    Decimal result1 = new Decimal(0);
                                    Decimal result2 = new Decimal(0);
                                    num4 = new Decimal(0);
                                    Decimal nValue = (e.Editor as NumericTextBox).nValue;
                                    Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result2);
                                    Decimal num13 = SysFunc.Round(nValue * result2 / new Decimal(100), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["ck"].Value = (object)num13;
                                    Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result1);
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((nValue - num13) * result1 / new Decimal(100), StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "tl_ck":
                            if (e.Cell.IsDataChanged)
                            {
                                num12 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal result1 = new Decimal(0);
                                num3 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                Decimal result3 = new Decimal(0);
                                Decimal result4 = (e.Editor as NumericTextBox).nValue;
                                Decimal.TryParse(e.Cell.Record.Cells["tien_nt2"].Value.ToString(), out result1);
                                Decimal.TryParse(e.Cell.Record.Cells["tien2"].Value.ToString(), out result2);
                                Decimal.TryParse(e.Cell.Record.Cells["thue_suat"].Value.ToString(), out result3);
                                Decimal.TryParse(e.Cell.Record.Cells["tl_ck"].Value.ToString(), out result4);
                                Decimal num13 = result1 * result4 / new Decimal(100);
                                Decimal num14 = result2 * result4 / new Decimal(100);
                                e.Cell.Record.Cells["ck_nt"].Value = (object)SysFunc.Round(num13, StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["ck"].Value = (object)SysFunc.Round(num14, StartUpTrans.M_ROUND);
                                e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((result1 - num13) * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((result2 - num14) * result3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "ma_thue":
                            AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl3.IsDataChanged)
                            {
                                Decimal num13 = new Decimal(0);
                                Decimal num14 = new Decimal(0);
                                num11 = new Decimal(0);
                                Decimal result1 = new Decimal(0);
                                Decimal result2 = new Decimal(0);
                                if (autoCompleteControl3.RowResult != null)
                                {
                                    e.Cell.Record.Cells["thue_suat"].Value = autoCompleteControl3.RowResult["thue_suat"];
                                    Decimal num15 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                                    Decimal num16 = this.ParseDecimal(e.Cell.Record.Cells["tien2"].Value, new Decimal(0));
                                    Decimal num17 = this.ParseDecimal(e.Cell.Record.Cells["thue_suat"].Value, new Decimal(0));
                                    Decimal.TryParse(e.Cell.Record.Cells["ck_nt"].Value.ToString(), out result2);
                                    Decimal.TryParse(e.Cell.Record.Cells["ck"].Value.ToString(), out result1);
                                    if (this.cbMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    {
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num15 - result2) * num17 / new Decimal(100), StartUpTrans.M_ROUND);
                                        e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                    }
                                    else
                                    {
                                        e.Cell.Record.Cells["thue_nt"].Value = (object)SysFunc.Round((num15 - result2) * num17 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                        e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num16 - result1) * num17 / new Decimal(100), StartUpTrans.M_ROUND);
                                    }
                                }
                                else
                                {
                                    e.Cell.Record.Cells["thue_suat"].Value = (object)0;
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)0;
                                    e.Cell.Record.Cells["thue"].Value = (object)0;
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "ck_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                }
                                else
                                {
                                    Decimal nValue = this.txtTy_gia.nValue;
                                    e.Cell.Record.Cells["ck"].Value = (object)SysFunc.Round(Convert.ToDecimal(e.Cell.Record.Cells["ck_nt"].Value) * nValue, StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
                                break;
                            }
                            break;
                        case "ck":
                            this.Sum_ALL();
                            break;
                        case "thue":
                            this.Sum_ALL();
                            break;
                        case "thue_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.cbMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                }
                                else
                                {
                                    Decimal nValue = this.txtTy_gia.nValue;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(Convert.ToDecimal(e.Cell.Record.Cells["thue_nt"].Value) * nValue, StartUpTrans.M_ROUND);
                                }
                                this.Sum_ALL();
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

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            if(this.gbTotal.Visibility == Visibility.Visible)
            {
                this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtma_thck.IsFocus = true));
            }    
            else
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
            }    
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmPoctpna.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    if (!(this.GrdCt.ActiveRecord is DataRecord activeRecord) || activeRecord.Cells["ma_vt"].Value == null || activeRecord.Cells["ma_vt"].Value.ToString() == "")
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
                case Key.F8:
                    if (ExMessageBox.Show(2225, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No || !(this.GrdCt.ActiveRecord is DataRecord activeRecord1))
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
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result3);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"].ToString(), out result2);
                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"].ToString(), out result4);
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
                        int num = (int)ExMessageBox.Show(2230, StartupBase.SasObj, "Chưa vào mã khách hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag = true;
                    }
                    else if (string.IsNullOrEmpty(this.txtNgay_ct.Text.ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2245, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                        else if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString()))
                        {
                            int num2 = (int)ExMessageBox.Show(2250, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCt.Focus();
                            flag = true;
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(2260, StartupBase.SasObj, "Chưa vào số đơn hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Focus();
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            flag = true;
                        }
                        else if (result1 != result2 || result3 != result4)
                        {
                            int num2 = (int)ExMessageBox.Show(2265, StartupBase.SasObj, "Tổng chiết khấu khác với chiết khấu tổng cộng của các vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D2);
                            flag = true;
                        }
                    }
                    if (!flag && StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    {
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(2270, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt"];
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
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_ct"] = (object)StartUpTrans.Ma_ct;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_hd"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"];
                            if (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tl_ck"], new Decimal(0)) == new Decimal(0))
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_ck"] = (object)string.Empty;
                        }
                        this.Sum_ALL();
                        Decimal result1 = new Decimal(0);
                        Decimal result2 = new Decimal(0);
                        Decimal num = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result2);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = (object)result1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)result2;
                        bool? isChecked = this.ChkSuaTien.IsChecked;
                        if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && num != new Decimal(0))
                            this.CanBangTien();
                        this.PhanBoThueInCT();
                        StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                        StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_hd"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_ct"] = StartUpTrans.DmctInfo["ct_nxt"];
                    if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("search"))
                    {
                        DataTable table = StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable();
                        SysFunc.SetStrSearch(StartupBase.SasObj, "dmhd", ref table);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["search"] = (object)table.Rows[0]["search"].ToString().Trim();
                    }
                    LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable1.Rows[0]["status"] = (object)0;
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                    {
                        if (!this.IsSequenceSave)
                            dataRowView.Row["so_ct"] = (object)this.txtSo_ct.Text;
                        LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                    }
                    for (int index = 0; index < LocalTable2.Rows.Count; ++index)
                    {
                        if (LocalTable2.Rows[index]["dvt1"].ToString().Trim() == LocalTable2.Rows[index]["dvt"].ToString().Trim())
                        {
                            LocalTable2.Rows[index]["he_so1"] = 0;
                            LocalTable2.Rows[index]["so_luong1"] = 0;
                        }
                        if (String.IsNullOrEmpty(LocalTable2.Rows[index]["he_so1"].ToString()))
                            LocalTable2.Rows[index]["he_so1"] = (object)0;

                        if ((Decimal)LocalTable2.Rows[index]["he_so1"] != new Decimal(0))
                        {
                            LocalTable2.Rows[index]["so_luong1"] = (Decimal)(Convert.ToDecimal(LocalTable2.Rows[index]["so_luong"]) * Convert.ToDecimal(LocalTable2.Rows[index]["he_so1"]));
                        }
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num1 = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else
                    {
                        string sqlbaogia =  string.Format("Update ph101 set status={0}, status2={0} where ma_hd = '{1}'","3", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_baogia"].ToString());
                        int result = StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(sqlbaogia));                      
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
                                            case "PH02":
                                                int num3 = (int)ExMessageBox.Show(2300, StartupBase.SasObj, "Tk có là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = true;
                                                this.txtMa_nx.IsFocus = true;
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

        public new string GetLastSoct(SasObject SasObj, string ma_qs)
        {
            try
            {
                string cmdText = "SELECT MAX(so_ct) FROM dmhd WHERE ma_qs='" + ma_qs.Trim() + "';";
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

        private void cbMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.cbMa_nt.IsDataChanged)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.cbMa_nt.RowResult["loai_tg"];
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (this.cbMa_nt.RowResult != null)
            {
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
            if (!string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_kh.RowResult["dia_chi"].ToString().Trim();
            if (StartUpTrans.M_LAN.Equals("V"))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"] = (object)this.txtMa_kh.RowResult["ten_kh"].ToString().Trim();
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh2"] = (object)this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim();
            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"] == DBNull.Value || !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"].ToString()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["han_tt"] = (object)this.txtMa_kh.RowResult["han_tt"];
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["nh_kh3"] = (object)this.txtMa_kh.RowResult["nh_kh3"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = (object)this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim();
            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"].ToString().Trim() == "")
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_thck"] = (object)this.txtMa_kh.RowResult["ma_thck"].ToString().Trim();
            this.txtma_thck.SearchInit();
            this.txtma_thck_PreviewLostFocus((object)this.txtma_thck, (KeyboardFocusChangedEventArgs)null);
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ong_ba"] = (object)this.txtMa_kh.RowResult["doi_tac"].ToString().Trim();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"] = string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()) ? (object)this.txtMa_kh.RowResult["tk"].ToString().Trim() : (object)this.txtMa_nx.Text.Trim();
            this.txtDiaChiFocusable = string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim());

            this.txtSobaogia.Filter = "ma_kh = '" +this.txtMa_kh.Text.Trim()+ "' and status = '2'";
        }

        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            this.txtDia_chi.IsTabStop = false;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_nx.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = (object)"";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = (object)"";
            }
        }

        private void txtMa_nguoi_ban_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_bp.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp"] = (object)this.txtMa_bp.RowResult["ten_bp"].ToString();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp2"] = (object)this.txtMa_bp.RowResult["ten_bp2"].ToString();
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp"] = (object)"";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp2"] = (object)"";
            }
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
                this.txtTy_gia.Value = this.cbMa_nt.RowResult["ma_nt"] == (object)StartUpTrans.M_ma_nt0 ? (object)1 : this.txtTy_gia.Value;
            if (string.IsNullOrEmpty(this.txtTy_gia.Text.ToString()))
                this.txtTy_gia.Value = (object)0;
            try
            {
                if (FormTrans.currActionTask == ActionTask.Delete || !FrmPoctpna.IsInEditMode.Value || (this.txtTy_gia.Value == null || this.txtTy_gia.Value == DBNull.Value || !(this.ParseDecimal(this.txtTy_gia.Value, new Decimal(0)) != new Decimal(0))))
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
                Decimal num10 = new Decimal(0);
                Decimal num11 = new Decimal(0);
                Decimal num12 = new Decimal(0);
                Decimal nValue = this.txtTy_gia.nValue;
                num6 = this.txtT_Tien_nt.Value == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(this.txtT_Tien_nt.Value);

                Decimal num23 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"], new Decimal(0));
                if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)SysFunc.Round((num23 / nValue), StartUpTrans.M_ROUND_NT);
                }
                else
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)SysFunc.Round(num23, StartUpTrans.M_ROUND_NT);
                }
                if (this.GrdCt.Records.Count > 0 && (this.GrdCt.DataSource as DataView).Table.DefaultView[0]["ma_vt"] != DBNull.Value)
                {
                    for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                    {
                        if ((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt2"].Value != DBNull.Value)
                        {
                            Decimal num13 = (this.GrdCt.DataSource as DataView)[index]["so_luong"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["so_luong"].Value);
                            Decimal num14 = (this.GrdCt.DataSource as DataView)[index]["gia_nt2"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["gia_nt2"].Value);
                            Decimal num15 = (this.GrdCt.DataSource as DataView)[index]["tien_nt2"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt2"].Value);
                            Decimal num16 = (this.GrdCt.DataSource as DataView)[index]["thue_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["thue_nt"].Value);
                            Decimal num17 = (this.GrdCt.DataSource as DataView)[index]["ck_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["ck_nt"].Value);
                            Decimal num18 = (this.GrdCt.DataSource as DataView)[index]["gia_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["gia_nt"].Value);
                            Decimal num19 = (this.GrdCt.DataSource as DataView)[index]["tien_nt"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal((this.GrdCt.Records[index] as DataRecord).Cells["tien_nt"].Value);
                            if (num13 * num14 != new Decimal(0))
                            {
                                num15 = SysFunc.Round(num13 * num14, StartUpTrans.M_ROUND_NT);
                                (this.GrdCt.DataSource as DataView)[index]["tien_nt2"] = (object)num15;
                            }
                            if (nValue * num14 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["gia2"] = (object)SysFunc.Round(nValue * num14, StartUpTrans.M_ROUND_GIA);
                            if (nValue * num15 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["tien2"] = (object)SysFunc.Round(nValue * num15, StartUpTrans.M_ROUND);
                            if (nValue * num16 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["thue"] = (object)SysFunc.Round(nValue * num16, StartUpTrans.M_ROUND);
                            if (nValue * num17 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["ck"] = (object)SysFunc.Round(nValue * num17, StartUpTrans.M_ROUND);
                            if (nValue * num18 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["gia"] = (object)SysFunc.Round(nValue * num18, StartUpTrans.M_ROUND);
                            if (nValue * num19 != new Decimal(0))
                                (this.GrdCt.DataSource as DataView)[index]["tien"] = (object)SysFunc.Round(nValue * num19, StartUpTrans.M_ROUND);
                        }
                    }
                    this.UpdateTotalKM("tien_km", "tien_km_nt", false);
                    this.Sum_ALL();
                    this.ReSum_ALL();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        private void UpdateTotalKM(string columnname, string columnname_nt, bool isPrice)
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname], new Decimal(0));
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname_nt] = num2;
                }
                else
                {
                    Decimal num3 = SysFunc.Round(num2 / num1, !isPrice ? StartUpTrans.M_ROUND : StartUpTrans.M_ROUND_GIA_NT);
                    if (num3 != new Decimal(0))
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname_nt] = num3;
                }
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
        private void ReSum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            Decimal num10;
            Decimal num11;
            Decimal num12;
            Decimal num13;
            Decimal num14;
            Decimal num15;
            Decimal num16;
            Decimal num17;
            Decimal num18;
            Decimal num19;
            Decimal num20;
            Decimal num21;
            Decimal kmhd = new Decimal(0);
            Decimal kmhd_nt = new Decimal(0);

            kmhd = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"]);
            kmhd_nt = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"]);
            if (this.cbMa_nt.Text.Equals(StartUpTrans.M_ma_nt0))
            {
                num10 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num11 = num10;
                num12 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num13 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num18 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_km)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num19 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_km_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num20 = (num12 + kmhd + num18);
                num21 = (num13 + kmhd_nt + num19);
                num14 = num11 - num20;
                num15 = num10 - num21;
                num16 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num17 = num16;
            }
            else
            {
                num11 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num10 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num12 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num13 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num18 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_km)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num19 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_km_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num20 = (num12 + kmhd + num18);
                num21 = (num13 + kmhd_nt + num19);
                num14 = num11 - num20;
                num15 = num10 - num21;
                num17 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num16 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
            }

            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = (object)num11;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"] = (object)num10;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"] = (object)num20;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"] = (object)num21;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck"] = (object)num14;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck_nt"] = (object)num15;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)num17;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num16;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(num14 + num17);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(num15 + num16);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));

        }
        private void Sum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            Decimal num10;
            Decimal num11;
            Decimal num12;
            Decimal num13;
            Decimal num14;
            Decimal num15;
            Decimal num16;
            Decimal num17;

            if (this.cbMa_nt.Text.Equals(StartUpTrans.M_ma_nt0))
            {
                num10 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num11 = num10;
                num12 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num13 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num14 = num11 - num12;
                num15 = num10 - num13;
                num16 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num17 = num16;
            }
            else
            {
                num11 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num10 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num12 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num13 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
                num14 = num11 - num12;
                num15 = num10 - num13;
                num17 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND);
                num16 = SysFunc.Round(this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0)), StartUpTrans.M_ROUND_NT);
            }

            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = (object)num11;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"] = (object)num10;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"] = (object)num12;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"] = (object)num13;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck"] = (object)num14;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck_nt"] = (object)num15;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)num17;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)num16;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(num14 + num17);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(num15 + num16);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));
        }
        private void Tinhkhuyenmai()
        {
            ResetKhuyenmai();
            StartUp.CTKMTable.Clear();
            if (StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable().Rows.Count > 0)
            {
                Decimal tienhang = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"]);
                String ma_kh = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString();
                if (!string.IsNullOrEmpty(ma_kh) && tienhang > 0)
                {
                    string sql = "Exec checkkmtheohoadon @tien, @ma_kh,@ngay";
                    SqlCommand kmcmd = new SqlCommand(sql);
                    kmcmd.Parameters.Add("@tien", SqlDbType.Decimal).Value = tienhang;
                    kmcmd.Parameters.Add("@ma_kh", SqlDbType.Char).Value = (object)ma_kh.Trim();
                    kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"]).ToString("yyyy-MM-dd");

                    DataTable tblhd = StartupBase.SasObj.ExcuteReader(kmcmd).Tables[0];
                    if (tblhd.Rows.Count > 0)
                    {
                        foreach (DataRow nr in tblhd.Rows)
                        {
                            DataRow r1 = StartUp.CTKMTable.NewRow();
                            r1["stt_rec"] = nr["stt_rec"];
                            r1["ischoose"] = (object)true;
                            r1["pt_ck"] = nr["pt_ck"];
                            r1["tien_ck"] = nr["tien_ck"];
                            r1["loai"] = 1;
                            r1["loai_km"] = nr["loai_km"];
                            r1["ten_km"] = (object)"KMHD: " + nr["ten_km"].ToString().Trim();
                            r1["ngay_bd"] = nr["ngay_bd"];
                            r1["ngay_kt"] = nr["ngay_kt"];
                            StartUp.CTKMTable.Rows.Add(r1);
                        }
                    }
                }
                if (!string.IsNullOrEmpty(ma_kh) && StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Rows)
                    {
                        if (!Convert.ToBoolean(r["khuyen_mai"]))
                        {
                            String ma_vt = r["ma_vt"].ToString();
                            string sql = "Exec checkkmtheohanghoa @ma_vt, @ma_kh,@ngay";
                            SqlCommand kmcmd = new SqlCommand(sql);
                            kmcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)ma_vt.Trim();
                            kmcmd.Parameters.Add("@ma_kh", SqlDbType.Char).Value = (object)ma_kh.Trim();
                            kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"]).ToString("yyyy-MM-dd");

                            DataTable tblhh = StartupBase.SasObj.ExcuteReader(kmcmd).Tables[0];
                            if (tblhh.Rows.Count > 0)
                            {
                                foreach (DataRow nr in tblhh.Rows)
                                {
                                    DataRow r1 = StartUp.CTKMTable.NewRow();
                                    r1["stt_rec"] = nr["stt_rec"];
                                    r1["ischoose"] = (object)true;
                                    r1["ma_vt"] = nr["ma_vt"];
                                    r1["ten_vt"] = nr["ten_vt"];
                                    r1["pt_ck"] = nr["pt_ck"];
                                    r1["tien_ck"] = nr["tien_ck"];
                                    r1["loai"] = 2;
                                    r1["ma_vt_mua"] = ma_vt.Trim();
                                    r1["loai_km"] = nr["loai_km"];
                                    r1["ten_km"] = (object)"KMHH: " + nr["ten_km"].ToString().Trim() + ": " + r["ma_vt"].ToString().Trim();
                                    r1["ngay_bd"] = nr["ngay_bd"];
                                    r1["ngay_kt"] = nr["ngay_kt"];
                                    StartUp.CTKMTable.Rows.Add(r1);
                                }
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(ma_kh) && StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Rows.Count > 0)
                {
                    foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Rows)
                    {
                        if (!Convert.ToBoolean(r["khuyen_mai"]))
                        {
                            String ma_vt = r["ma_vt"].ToString();
                            Double so_Luong = Convert.ToDouble(r["so_luong"]);
                            string sql = "Exec checkkmhangtanghang @ma_vt, @ma_kh, @so_luong,@ngay";
                            SqlCommand kmcmd = new SqlCommand(sql);
                            kmcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)ma_vt.Trim();
                            kmcmd.Parameters.Add("@ma_kh", SqlDbType.Char).Value = (object)ma_kh.Trim();
                            kmcmd.Parameters.Add("@so_luong", SqlDbType.Decimal).Value = (object)so_Luong;
                            kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"]).ToString("yyyy-MM-dd");

                            DataTable tblhth = StartupBase.SasObj.ExcuteReader(kmcmd).Tables[0];
                            if (tblhth.Rows.Count > 0)
                            {
                                foreach (DataRow nr in tblhth.Rows)
                                {
                                    DataRow r1 = StartUp.CTKMTable.NewRow();
                                    r1["stt_rec"] = nr["stt_rec"];
                                    r1["ischoose"] = (object)true;
                                    r1["loai"] = 3;
                                    r1["ma_vt_mua"] = ma_vt.Trim();
                                    r1["so_luong_mua"] = nr["so_luong_mua"];
                                    r1["nhom_km"] = nr["nhom_km"];
                                    r1["loai_km"] = nr["loai_km"];
                                    r1["boi_so"] = nr["boi_so"];
                                    string message = "";
                                    if (Convert.ToBoolean(nr["boi_so"]))
                                        message = " (bội số)";

                                    r1["ten_km"] = (object)nr["ten_km"].ToString().Trim() + ": " + ma_vt.Trim() + " - " + message;
                                    r1["ngay_bd"] = nr["ngay_bd"];
                                    r1["ngay_kt"] = nr["ngay_kt"];
                                    r1["dvt1"] = nr["dvt1"];
                                    StartUp.CTKMTable.Rows.Add(r1);
                                }
                            }
                        }
                    }
                }
                StartUp.CTKMTable.DefaultView.Sort = "so_luong_mua desc";
                StartUp.CTKMTable = StartUp.CTKMTable.DefaultView.ToTable();
                if (StartUp.CTKMTable.Rows.Count > 0)
                    LaysanphamkhuyenmaiHTH();

            }
        }
        private void LaysanphamkhuyenmaiHTH()
        {
            StartUp.VattuKMHTHTable.Clear();
            rkhuyenmaihth = StartUp.CTKMTable.Select("ischoose = 1 and loai='3'");
            if (rkhuyenmaihth.Length > 0)
            {
                foreach (DataRow r in rkhuyenmaihth)
                {
                    string sql = "Exec laysanphamkhuyenmai @stt_rec, @ma_nh";
                    SqlCommand kmcmd = new SqlCommand(sql);
                    kmcmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = (object)r["stt_rec"].ToString().Trim();
                    kmcmd.Parameters.Add("@ma_nh", SqlDbType.Char).Value = (object)r["nhom_km"].ToString().Trim();
                    DataTable tblvattu = StartupBase.SasObj.ExcuteReader(kmcmd).Tables[0];
                    if (tblvattu.Rows.Count > 0)
                    {
                        foreach (DataRow nr in tblvattu.Rows)
                        {
                            DataRow rvt = StartUp.VattuKMHTHTable.NewRow();
                            rvt["ischoose"] = true;
                            rvt["stt_rec"] = nr["stt_rec"];
                            rvt["ma_vt"] = nr["ma_vt"];
                            rvt["ten_vt"] = nr["ten_vt"];
                            rvt["so_luong"] = nr["so_luong"];
                            rvt["boi_so"] = (object)r["boi_so"];
                            rvt["so_luong_mua"] = (object)r["so_luong_mua"];
                            rvt["ma_vt_mua"] = (object)r["ma_vt_mua"];
                            rvt["nhom_km"] = (object)r["nhom_km"];
                            rvt["ma_nh_km"] = (object)r["ma_nh_km"];
                            rvt["dvt1"] = r["dvt1"];
                            StartUp.VattuKMHTHTable.Rows.Add(rvt);
                        }
                    }
                }
            }
            if (this.ChkAutoKM.IsChecked == true)
                ThemKhuyenmai();
        }
        private void ThemKhuyenmai()
        {
            Decimal ty_gia = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            foreach (DataRow r in StartUp.CTKMTable.Rows)
            {
                if (Convert.ToBoolean(r["ischoose"]) && r["loai"].ToString().Trim().Equals("1"))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                    Decimal result3 = Convert.ToDecimal(r["pt_ck"]);
                    Decimal result4 = Convert.ToDecimal(r["tien_ck"]);
                    if (result3 > 0)
                    {
                        Decimal result1 = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"]);
                        if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                        {
                            Decimal num6 = SysFunc.Round(((result1 * result3) / 100), StartUpTrans.M_ROUND);
                            Decimal num7 = SysFunc.Round(((result1 * result3) / 100), StartUpTrans.M_ROUND_NT);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)num6;
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)num7;
                        }
                        else
                        {
                            Decimal num6 = SysFunc.Round(((result1 * result3) / 100), StartUpTrans.M_ROUND);
                            Decimal num7 = SysFunc.Round((num6 / ty_gia), StartUpTrans.M_ROUND_NT);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)num6;
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)num7;
                        }
                    }
                    else if (result4 > 0)
                    {
                        if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND_NT);
                        }
                        else
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)SysFunc.Round((result4 / ty_gia), StartUpTrans.M_ROUND_NT);
                        }
                    }
                }
                else if (Convert.ToBoolean(r["ischoose"]) && r["loai"].ToString().Trim().Equals("2"))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                    foreach (DataRow rdt in StartUpTrans.DsTrans.Tables[1].Rows)
                    {
                        if (!string.IsNullOrEmpty(rdt["ma_vt"].ToString()) && !string.IsNullOrEmpty(r["ma_vt_mua"].ToString()) && !string.IsNullOrEmpty(rdt["khuyen_mai"].ToString()))
                        {
                            string mvt = rdt["ma_vt"].ToString().Trim();
                            string mvtm = r["ma_vt_mua"].ToString().Trim();
                            bool vtkm = Convert.ToBoolean(rdt["khuyen_mai"]);
                            if (mvt.Equals(mvtm) && !vtkm)
                            {
                                rdt["ma_ctkm"] = r["stt_rec"];
                                Decimal result3 = Convert.ToDecimal(r["pt_ck"]);
                                Decimal result4 = Convert.ToDecimal(r["tien_ck"]);
                                if (result3 > 0)
                                {
                                    Decimal result1 = Convert.ToDecimal(rdt["tien_nt2"]);
                                    if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                    {
                                        Decimal num1 = (result3 * result1) / 100;
                                        rdt["tien_km"] = SysFunc.Round(num1, StartUpTrans.M_ROUND);
                                        rdt["tien_km_nt"] = SysFunc.Round(num1, StartUpTrans.M_ROUND_NT);
                                        rdt["pt_km"] = result3;
                                    }
                                    else
                                    {
                                        Decimal num1 = (result3 * result1) / 100;
                                        Decimal num2 = (((result3 * result1) / 100) / ty_gia);
                                        rdt["tien_km"] = SysFunc.Round(num1, StartUpTrans.M_ROUND);
                                        rdt["tien_km_nt"] = SysFunc.Round(num2, StartUpTrans.M_ROUND_NT);
                                        rdt["pt_km"] = result3;
                                    }
                                }
                                else if (result4 > 0)
                                {
                                    if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                    {
                                        rdt["tien_km"] = SysFunc.Round(result4, StartUpTrans.M_ROUND);
                                        rdt["tien_km_nt"] = SysFunc.Round(result4, StartUpTrans.M_ROUND_NT);
                                    }
                                    else
                                    {
                                        rdt["tien_km"] = result4;
                                        rdt["tien_km_nt"] = SysFunc.Round((result4 / ty_gia), StartUpTrans.M_ROUND_NT);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            DataTable tmpTable = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable();
            tmpTable.DefaultView.Sort = "stt_rec0";
            tmpTable = tmpTable.DefaultView.ToTable();
            DataTable CT81New = tmpTable.Copy();
            CT81New.Rows.Clear();
            newstt_rec0 = 0;

            for (int x = 0; x < tmpTable.Rows.Count; x++)
            {
                // DataRow rct81 = tmpTable.Rows[x];
                newstt_rec0++;
                ct81_ma_vt = tmpTable.Rows[x]["ma_vt"].ToString().Trim();
                tmpTable.Rows[x]["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)newstt_rec0);
                int soluong = Convert.ToInt32(tmpTable.Rows[x]["so_luong"]);
                if (rkhuyenmaihth.Length > 0)
                {
                    foreach (DataRow r in rkhuyenmaihth)
                    {
                        string vt_ctkm = r["ma_vt_mua"].ToString();
                        int soluong_cttkm = Convert.ToInt32(r["so_luong_mua"]);
                        if (soluong >= soluong_cttkm)
                        {
                            if (ct81_ma_vt.Equals(vt_ctkm.Trim()))
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                                bool boiso = Convert.ToBoolean(r["boi_so"]);
                                int heso = 1;
                                if (boiso)
                                    heso = soluong / soluong_cttkm;

                                string stt_recctkm = r["stt_rec"].ToString().Trim();
                                string nhom_kmctkm = r["nhom_km"].ToString().Trim();
                                DataRow[] rvt = StartUp.VattuKMHTHTable.Select("stt_rec = '" + stt_recctkm + "' and nhom_km = '" + nhom_kmctkm + "' and ischoose = 1");
                                if (rvt.Length > 0)
                                {
                                    for (int y = 0; y < rvt.Length; y++)
                                    {
                                        //checked trùng cộng dồn số lượng
                                        bool addnew = true;
                                        foreach (DataRow r81 in CT81New.Rows)
                                        {
                                            string ma_vt = r81["ma_vt"].ToString().Trim();
                                            string stt_rec = r81["stt_rec"].ToString().Trim();
                                            string ma_ctkm = r81["ma_ctkm"].ToString().Trim();
                                            if (ma_vt.Equals(rvt[y]["ma_vt"].ToString().Trim()) && ma_ctkm.Equals(rvt[y]["stt_rec"].ToString().Trim()))
                                            {
                                                r81["so_luong"] = Convert.ToInt32(r81["so_luong"]) + Convert.ToInt32(rvt[y]["so_luong"]) * heso;
                                                r81["ma_nh_km"] = r81["ma_nh_km"].ToString().Trim() + "+" + rvt[y]["nhom_km"].ToString().Trim();
                                                addnew = false;
                                                break;
                                            }
                                        }

                                        if (addnew)
                                        {
                                            DataRow rnew = CT81New.NewRow();
                                            newstt_rec0++;
                                            rnew["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)newstt_rec0);
                                            rnew["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                            rnew["ma_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"];
                                            rnew["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                                            rnew["ma_kho_i"] = tmpTable.Rows[x]["ma_kho_i"];
                                            rnew["gia2"] = 0;
                                            rnew["gia_nt2"] = 0;
                                            rnew["pt_km"] = 0;
                                            rnew["tien_km"] = 0;
                                            rnew["tien_km_nt"] = 0;
                                            rnew["ma_nh_km"] = rvt[y]["nhom_km"];
                                            rnew["ma_ctkm"] = rvt[y]["stt_rec"];
                                            rnew["ma_vt"] = rvt[y]["ma_vt"];
                                            rnew["ten_vt"] = rvt[y]["ten_vt"];
                                            rnew["dvt1"] = rvt[y]["dvt1"];
                                            rnew["dvt"] = rvt[y]["dvt1"];
                                            rnew["khuyen_mai"] = (object)true;
                                            rnew["so_luong"] = Convert.ToInt32(rvt[y]["so_luong"]) * heso;
                                            tmpTable.Rows[x]["ma_ctkm"] = rvt[y]["stt_rec"];
                                            tmpTable.Rows[x]["ma_nh_km"] = rvt[y]["nhom_km"];
                                            //string sql_tkdt = "Select gia_ton,tk_vt,tk_dt,tk_gv from dmvt where ma_vt = '" + rvt[y]["ma_vt"].ToString().Trim() + "'";
                                            //DataTable tbltk = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql_tkdt)).Tables[0];
                                            //if (tbltk.Rows.Count > 0)
                                            //{
                                            //    rnew["gia_ton"] = tbltk.Rows[0]["gia_ton"];
                                            //    rnew["tk_vt"] = tbltk.Rows[0]["tk_vt"];
                                            //    rnew["tk_dt"] = tbltk.Rows[0]["tk_dt"];
                                            //    rnew["tk_gv"] = tbltk.Rows[0]["tk_gv"];
                                            //}
                                            CT81New.Rows.Add(rnew);
                                        }
                                    }
                                }
                                soluong = soluong - (heso * soluong_cttkm);
                            }
                        }
                    }
                }
            }
            StartUpTrans.DsTrans.Tables[1].Rows.Clear();
            tmpTable.Merge(CT81New);
            StartUpTrans.DsTrans.Tables[1].Merge(tmpTable);
            ReSum_ALL();
        }
        private void ThemKhuyenmai1()
        {
            foreach (DataRow r in StartUp.CTKMTable.Rows)
            {
                if (Convert.ToBoolean(r["ischoose"]) && r["loai"].ToString().Trim().Equals("1"))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                    Decimal result3 = Convert.ToDecimal(r["pt_ck"]);
                    Decimal result4 = Convert.ToDecimal(r["tien_ck"]);
                    if (result3 > 0)
                    {
                        Decimal result1 = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"]);
                        Decimal result2 = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"]);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)((result1 * result3) / 100);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)((result2 * result3) / 100);
                    }
                    else if (result4 > 0)
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = (object)result4;
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object)result4;
                    }
                }
                else if (Convert.ToBoolean(r["ischoose"]) && r["loai"].ToString().Trim().Equals("2"))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                    foreach (DataRow rdt in StartUpTrans.DsTrans.Tables[1].Rows)
                    {
                        if (!string.IsNullOrEmpty(rdt["ma_vt"].ToString()) && !string.IsNullOrEmpty(r["ma_vt_mua"].ToString()) && !string.IsNullOrEmpty(rdt["khuyen_mai"].ToString()))
                        {
                            string mvt = rdt["ma_vt"].ToString().Trim();
                            string mvtm = r["ma_vt_mua"].ToString().Trim();
                            bool vtkm = Convert.ToBoolean(rdt["khuyen_mai"]);
                            if (mvt.Equals(mvtm) && !vtkm)
                            {
                                rdt["ma_ctkm"] = r["stt_rec"];
                                Decimal result3 = Convert.ToDecimal(r["pt_ck"]);
                                Decimal result4 = Convert.ToDecimal(r["tien_ck"]);
                                if (result3 > 0)
                                {
                                    Decimal result1 = Convert.ToDecimal(rdt["tien2"]);
                                    Decimal result2 = Convert.ToDecimal(rdt["tien_nt2"]);

                                    rdt["tien_km"] = (result3 * result1) / 100;
                                    rdt["tien_km_nt"] = (result3 * result2) / 100;
                                    rdt["ck"] = (result3 * result1) / 100;
                                    rdt["ck_nt"] = (result3 * result2) / 100;
                                    rdt["pt_km"] = result3;
                                }
                                else if (result4 > 0)
                                {
                                    rdt["tien_km"] = rdt["tien_km_nt"] = rdt["ck"] = rdt["ck_nt"] = result4;
                                }
                            }
                        }
                    }
                }
                else if (Convert.ToBoolean(r["ischoose"]) && r["loai"].ToString().Trim().Equals("3"))
                {
                    int manhkm = 0;
                    string sql = "select isnull(max(ma_nh_km),1) as ma_nh_km from dmhdct";
                    DataTable tblManhomkm = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql)).Tables[0];
                    if (tblManhomkm.Rows.Count > 0)
                    {
                        manhkm = Convert.ToInt32(tblManhomkm.Rows[0]["ma_nh_km"].ToString()) + 1;
                    }
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = r["stt_rec"];
                    DataTable tmpTable = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable();
                    for (int x = 0; x < tmpTable.Rows.Count; x++)
                    {
                        string vt1 = tmpTable.Rows[x]["ma_vt"].ToString();
                        bool vtkm = Convert.ToBoolean(tmpTable.Rows[x]["khuyen_mai"]);
                        string vt2 = r["ma_vt_mua"].ToString();
                        if (!string.IsNullOrEmpty(vt1) && !string.IsNullOrEmpty(vt2) && !vtkm)
                        {
                            if (vt1.Trim().Equals(vt2.Trim()))
                            {
                                StartUpTrans.DsTrans.Tables[1].DefaultView[x]["ma_nh_km"] = manhkm;
                                StartUpTrans.DsTrans.Tables[1].DefaultView[x]["stt_rec0"] = GetMaxSttrec0();
                                if (StartUp.VattuKMHTHTable.Rows.Count > 0)
                                {
                                    foreach (DataRow rkm in StartUp.VattuKMHTHTable.Rows)
                                    {
                                        string mvt = rkm["ma_vt_mua"].ToString().Trim();
                                        string mvtm = r["ma_vt_mua"].ToString().Trim();
                                        if (mvt.Equals(mvtm))
                                        {
                                            int boiso = 1;
                                            if (Convert.ToBoolean(rkm["boi_so"]))
                                                boiso = Convert.ToInt32(tmpTable.Rows[x]["so_luong"]) / Convert.ToInt32(rkm["so_luong_mua"]);
                                            DataRow rnew = StartUpTrans.DsTrans.Tables[1].NewRow();
                                            rnew["stt_rec0"] = GetMaxSttrec0();
                                            rnew["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                            rnew["ma_hd"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_hd"];
                                            rnew["ma_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"];
                                            rnew["ma_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"];
                                            rnew["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                                            rnew["gia2"] = 0;
                                            rnew["gia_nt2"] = 0;
                                            rnew["ma_nh_km"] = manhkm;
                                            rnew["ma_ctkm"] = rkm["stt_rec"];
                                            rnew["ma_vt"] = rkm["ma_vt"];
                                            rnew["ten_vt"] = rkm["ten_vt"];
                                            rnew["khuyen_mai"] = (object)true;
                                            rnew["so_luong"] = Convert.ToInt32(rkm["so_luong"]) * boiso;
                                            StartUpTrans.DsTrans.Tables[1].Rows.InsertAt(rnew, x + 1);
                                        }
                                    }
                                    manhkm++;
                                }
                            }
                        }
                    }
                }
            }
        }
        void ResetKhuyenmai()
        {
            foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].Rows)
            {
                r["Ma_nh_km"] = "";
                r["Ma_ctkm"] = "";
                r["Pt_km"] = 0;
                r["Tien_km"] = 0;
                r["Tien_km_nt"] = 0;
            }
            DataRow[] rdel = StartUpTrans.DsTrans.Tables[1].Select("khuyen_mai=1");
            foreach (DataRow r in rdel)
            {
                StartUpTrans.DsTrans.Tables[1].Rows.Remove(r);
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = "";
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = 0;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = 0;
        }
        string GetMaxSttrec0()
        {
            int result = 0;
            int num1 = 0;
            if (this.GrdCt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result);
            }
            int num2 = (result >= num1 ? result : num1) + 1;
            return string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num2);
        }
        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (StartUp.M_AR_CK == 0)
            {
                this.GrdCt.FieldLayouts[0].Fields["ck"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["ck"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["tl_ck"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tl_ck"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["ck_nt"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["ck_nt"].Settings.CellMaxWidth = 0.0;
            }
            
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
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", "stt_rec= '" + str + "'").ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", "stt_rec= '" + str + "'").ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(ck_nt)", "stt_rec= '" + str + "'").ToString(), out result5);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(ck)", "stt_rec= '" + str + "'").ToString(), out result6);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt2"].ToString(), out result9);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien2"].ToString(), out result10);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck_nt"].ToString(), out result7);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck"].ToString(), out result8);
                Decimal num5;
                Decimal num6;
                if (this.cbMa_nt.Text != StartUpTrans.M_ma_nt0)
                {
                    num5 = !(result9 == new Decimal(0)) ? (result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result9 - result7) / (result3 - result5) * result1, StartUpTrans.M_ROUND_NT)) : new Decimal(0);
                    num6 = !(result10 == new Decimal(0)) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result10 - result8) / (result4 - result6) * result2, StartUpTrans.M_ROUND)) : new Decimal(0);
                }
                else
                {
                    num5 = result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result9 - result7) / (result3 - result5) * result1, StartUpTrans.M_ROUND_NT);
                    num6 = result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round((result10 - result8) / (result4 - result6) * result2, StartUpTrans.M_ROUND);
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
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", "stt_rec= '" + str + "'").ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", "stt_rec= '" + str + "'").ToString(), out result4);
            Decimal result6 = new Decimal(0);
            Decimal result7 = new Decimal(0);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien2"].ToString(), out result6);
                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt2"].ToString(), out result7);
                Decimal num6 = !(this.cbMa_nt.Text != StartUpTrans.M_ma_nt0) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result7 / result4 * result2, StartUpTrans.M_ROUND_NT)) : (!(result7 == new Decimal(0)) ? (result4 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result7 / result4 * result2, StartUpTrans.M_ROUND_NT)) : (result3 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(result6 / result3 * result2, StartUpTrans.M_ROUND_NT)));
                Decimal num7 = !(num6 != new Decimal(0)) ? (!(result4 != new Decimal(0)) ? (!(result3 != new Decimal(0)) ? new Decimal(0) : SysFunc.Round(result6 / result3 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(result7 / result4 * result1, StartUpTrans.M_ROUND)) : SysFunc.Round(num6 * result5, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck"] = (object)num7;
                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck_nt"] = (object)num6;
                num2 += num7;
                num3 += num6;
            }
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ck"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ck"].ToString()) + (result1 - num2));
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ck_nt"] = (object)(Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ck_nt"].ToString()) + (result2 - num3));
        }

        private void txttong_cp_nt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
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
            int num = (int)ExMessageBox.Show(2310, StartupBase.SasObj, "Đã thực hiện xong phân bổ chi phí!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                            e.Cell.Record.Cells["cp"].Value = (object)(nValue2 * nValue1);
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
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"].ToString(), out result1);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"].ToString(), out result2);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), out result4);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = (object)SysFunc.Round(result1 * result5, StartUpTrans.M_ROUND);
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                if (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt2"], new Decimal(0)) != new Decimal(0))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien2"] = (object)(this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien2"], new Decimal(0)) + (this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"], new Decimal(0)) - result4));
                    break;
                }
            }
            Decimal result6 = new Decimal(0);
            Decimal result7 = new Decimal(0);
            Decimal result8 = new Decimal(0);
            Decimal result9 = new Decimal(0);
            Decimal result10 = new Decimal(0);
            Decimal result11 = new Decimal(0);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"].ToString(), out result6);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"].ToString(), out result7);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"].ToString(), out result8);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"].ToString(), out result9);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"].ToString(), out result10);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"].ToString(), out result11);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)(result6 - result8 + result10);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)(result7 - result9 + result11);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck"] = (object)(result7 - result9);
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_sau_ck_nt"] = (object)(result6 - result8);
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void txtSo_ct_me_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
        }

        private void txtMa_bp_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!(StartUp.M_BP_BH != "1"))
                return;
            this.txtNgay_ct.Focus();
        }

        private void txtHan_tt_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtHan_tt.SelectAll();
        }

        private void txtHan_tt_TextChanged(object sender, RoutedPropertyChangedEventArgs<string> e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
          {
              if (this.txtHan_tt.Text.IndexOf('-') < 0)
                  return;
              this.txtHan_tt.Value = (object)0;
          }));
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value != DBNull.Value)
                return;
            this.txtNgay_ct.Value = (object)DateTime.Now;
        }

        private void txtSo_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Edit || !(this.txtSo_ct.Text.Trim() != this.ma_hd) || !this.KiemTraCoPhatSinh())
                return;
            int num = (int)ExMessageBox.Show(2377, StartupBase.SasObj, "Đơn hàng đã có phát sinh, không thể sửa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtSo_ct.Text = this.ma_hd;
        }

        private void txtma_thck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtma_thck.RowResult != null)
            {
                if (this.txtHan_tt.Value == DBNull.Value || this.txtHan_tt.nValue == new Decimal(0))
                    this.txtHan_tt.Value = this.txtma_thck.RowResult["han_tt"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck"] = this.txtma_thck.RowResult["ten_thck"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck2"] = this.txtma_thck.RowResult["ten_thck2"];
            }
            else
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck"] = (object)"";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck2"] = (object)"";
            }
        }

        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[SODMHDB-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
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

        private void GrdCtKM_EditModeEnded(object sender, EditModeEndedEventArgs e)
        {

        }

        private void BtnHuyKM_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.CTKMTable.Rows.Count <= 0)
            {
                MessageBox.Show("Không có CTKM nào!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            ResetKhuyenmai();
            ReSum_ALL();
            foreach (DataRow r in StartUp.CTKMTable.Rows)
            {
                r["ischoose"] = (object)false;
            }
            this.ChkAutoKM.IsChecked = false;
        }

        private void BtnApdungKM_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.CTKMTable.Rows.Count <= 0)
            {
                MessageBox.Show("Không có CTKM nào!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            ResetKhuyenmai();
            ThemKhuyenmai();
            ReSum_ALL();
            this.ChkAutoKM.IsChecked = true;
            this.tiHT.IsSelected = true;
        }
        private void BtnCheckKM_Click(object sender, RoutedEventArgs e)
        {
            Tinhkhuyenmai();
        }
        private void txtSobaogia_GotFocus(object sender, RoutedEventArgs e)
        {

        }

        private void txtSobaogia_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtSobaogia.RowResult != null)
            {               
               string ma_baogia = this.txtSobaogia.RowResult["ma_hd"].ToString().Trim();
                string sql = "Select v_ct101.* from ph101 left join  v_ct101 on ph101.stt_rec = v_ct101.stt_rec where ph101.status2 <> '2' and ph101.status = '2' and ph101.ma_hd = '" + ma_baogia + "'";              
                DataTable tblvt = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql)).Tables[0];
                if(tblvt.Rows.Count>0)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_bp"] = txtSobaogia.RowResult["ma_bp"].ToString();                 
                    Themvattutubaogia(tblvt);
                }    
            }              
        }
        void Themvattutubaogia(DataTable tblvattu)
        {
            try
            {
                StartUpTrans.DsTrans.Tables[1].Rows.Clear();
                foreach (DataRow r in tblvattu.Rows)
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
                    dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];                   
                    dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
                    dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                    dataRow["ma_vt"] = (object)r["ma_vt"];
                    dataRow["ten_vt"] = (object)r["ten_vt"];
                    dataRow["ten_vt2"] = (object)r["ten_vt2"];
                    dataRow["dvt"] = (object)r["dvt"];
                    dataRow["dvt1"] = (object)r["dvt1"];
                    dataRow["so_luong"] = (object)r["so_luong"];
                    dataRow["gia_nt2"] = (object)r["gia_nt2"];
                    dataRow["tien_nt2"] = (object)r["tien_nt2"];
                    dataRow["gia_nt"] = (object)0;
                    dataRow["tien_nt"] = (object)0;
                    dataRow["tien2"] = (object)r["tien2"];
                    dataRow["gia2"] = (object)r["gia2"];
                    dataRow["tien"] = (object)0;
                    dataRow["gia"] = (object)0;
                    dataRow["ck_nt"] = (object)(object)r["ck_nt"];
                    dataRow["ck"] = (object)(object)r["ck"];
                    dataRow["thue_nt"] = (object)(object)r["thue_nt"];
                    dataRow["thue"] = (object)(object)r["thue"];
                    dataRow["han_gh_i"] = (object)r["han_giaohang"];                  
                    dataRow["ton13"] = (object)0;
                    dataRow["tien_km"] = (object)0;
                    dataRow["tien_km_nt"] = (object)0;
                    dataRow["khuyen_mai"] = (object)false;
                    FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
                    StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
                }
                Sum_ALL();
                Tinhkhuyenmai();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
    }
}
