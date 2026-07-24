using ArapLib;
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
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace Socthda
{
    public partial class FrmSocthda : FormTrans
    {
        public static int iRow = 0;
        public static string[] FieldKm = new string[17]
        {
      "km_ck",
      "tk_km",
      "tk_km_i",
      "t_sl_km",
      "t_tien_km_nt",
      "t_tien_km",
      "t_thue_km_nt",
      "t_thue_km",
      "tien_tc_nt",
      "tien_tc",
      "t_tt_km_nt",
      "t_tt_km",
      "tien_km",
       "tien_km_nt",
        "pt_km",
        "khuyen_mai",
         "ma_nh_km"
        };
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
        private int iRow_old = 0;
        public int so_ct_pt_length = 12;
        private FrmCopy _formcopy = (FrmCopy)null;
        private bool txtDiaChiFocusable = true;
        private bool allowKhuyenmai = true;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject IsCheckedPx_gia_dd;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject IsCheckedCo_km;
        private CodeValueBindingObject M_Ngay_lct;
        public static string hinhthuc_tt;
        public DataSet DsVitual;
        private DataSet dsCheckData;
        public int so_ct_px_length = 12;
        DataRow Vnpayrow = null;
        private string ct81_ma_vt;
        DataRow[] rkhuyenmaihth;
        public FrmSocthda()
        {
            this.InitializeComponent();
            this.LanguageProvider.Language = StartUpTrans.M_LAN;
            this.BindingSasObj = StartupBase.SasObj;
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_So_ct = this.txtSo_ct;
            this.C_Ma_nt = this.txtMa_nt;
            allowKhuyenmai = true;
        }

        private void FrmSocthda_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                StartUp.M_KM_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_KM_CK"));
                StartUp.M_AR_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_CK"));
                switch (this.stt_mau_temlate)
                {
                    case 120:
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);
                        this.GridKM1.Children.Remove((UIElement)this.tblt_tienkmhd);
                        this.GridKM1.Children.Remove((UIElement)this.txtTienkmhd_nt);
                        this.GridKM2.Children.Remove((UIElement)this.txtTienkmhd);
                        tabItemKM.Visibility = Visibility.Hidden;
                        allowKhuyenmai = false;

                        break;
                    case 121:
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);

                        break;
                    case 122:
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)2);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);
                        this.GridKM1.Children.Remove((UIElement)this.tblt_tienkmhd);
                        this.GridKM1.Children.Remove((UIElement)this.txtTienkmhd_nt);
                        this.GridKM2.Children.Remove((UIElement)this.txtTienkmhd);
                        tabItemKM.Visibility = Visibility.Hidden;
                        allowKhuyenmai = false;
                        break;
                    case 123:
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);
                        this.GridKM1.Children.Remove((UIElement)this.tblt_tienkmhd);
                        this.GridKM1.Children.Remove((UIElement)this.txtTienkmhd_nt);
                        this.GridKM2.Children.Remove((UIElement)this.txtTienkmhd);
                        tabItemKM.Visibility = Visibility.Hidden;
                        allowKhuyenmai = false;
                        break;


                    //Them gia von gia ban
                    case 225: //Gia von
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Settings.CellMaxWidth = 0.0;

                        break;
                    case 226: //Gia ban
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);
                        this.GridKM1.Children.Remove((UIElement)this.tblt_tienkmhd);
                        this.GridKM1.Children.Remove((UIElement)this.txtTienkmhd_nt);
                        this.GridKM2.Children.Remove((UIElement)this.txtTienkmhd);
                        tabItemKM.Visibility = Visibility.Hidden;
                        allowKhuyenmai = false;
                        
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Settings.CellMaxWidth = 0.0;

                        break;

                    case 227: //Gia ban - Gia von
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.txtHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHt_tt);
                        this.txtHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.txtHt_tt.SetValue(Grid.ColumnProperty, (object)1);
                        this.GrdLayoutHan_ck.Children.Remove((UIElement)this.tblHt_tt);
                        this.GrdLayoutHan_ck.Children.Add((UIElement)this.tblHt_tt);
                        this.tblHt_tt.SetValue(Grid.RowProperty, (object)0);
                        this.tblHt_tt.SetValue(Grid.ColumnProperty, (object)0);
                        this.GridKM1.Children.Remove((UIElement)this.tblt_tienkmhd);
                        this.GridKM1.Children.Remove((UIElement)this.txtTienkmhd_nt);
                        this.GridKM2.Children.Remove((UIElement)this.txtTienkmhd);
                        tabItemKM.Visibility = Visibility.Hidden;
                        allowKhuyenmai = false;

                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Settings.CellMaxWidth = 0.0;

                        this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Settings.CellMaxWidth = 0.0;

                        break;
                }

                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmSocthda.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                this.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.IsCheckedCo_km = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedCo_km");
                this.IsCheckedPx_gia_dd = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedPx_gia_dd");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                this.M_Ngay_lct = (CodeValueBindingObject)this.FormMain.FindResource((object)"M_Ngay_lct");
                this.M_Ngay_lct.Value = StartUpTrans.M_ngay_lct.Equals("1");
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtSoDuKH.Visibility = this.tblSoDuKH.Visibility = Visibility.Hidden;
                if (StartUp.M_SD_HDDT.Equals("1"))
                    this.tabHDDT.Visibility = Visibility.Visible;
                else
                    this.tabHDDT.Visibility = Visibility.Hidden;
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)this.IsInEditMode,
                    Mode = BindingMode.TwoWay
                });
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                this.M_LAN = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
                    this.LoadData();
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.IsCheckedPx_gia_dd.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["px_gia_dd"].ToString() == "1";
                    this.IsCheckedCo_km.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = false;
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                    this.so_ct_pt_length = StartupBase.SasObj.GetDatabaseFieldLength("so_ct");
                }
                if (StartUp.M_KM_CK == 0)
                {
                    this.ChkCo_km.Visibility = Visibility.Collapsed;
                    // this.GrdLayoutThue.RowDefinitions[2].Height = new GridLength(0.0);
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmSocthda.FieldKm).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;

                        }
                    }
                    using (IEnumerator<Field> enumerator = this.GrdTT.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmSocthda.FieldKm).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
                }
                if (StartUp.M_AR_CK == 0)
                {
                    this.GrdLayoutTong_NT.RowDefinitions[1].Height = new GridLength(0.0);
                    this.GrdLayoutTong_NT.RowDefinitions[2].Height = new GridLength(0.0);
                    this.GrdLayoutTong.RowDefinitions[1].Height = new GridLength(0.0);
                    this.GrdLayoutTong.RowDefinitions[2].Height = new GridLength(0.0);
                    this.GrdLayout00.RowDefinitions[4].Height = new GridLength(114.0);
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmSocthda.FieldCk).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
                }

                this.SetFocusToolbar();
                if (StartupBase.SasObj.GetOption("M_TON_KHO13").ToString() != "1" && this.GrdCt.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ton13")))
                    this.GrdCt.FieldLayouts[0].Fields["ton13"].Visibility = Visibility.Collapsed;

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

                //an thong tin theo mau temlate
                switch (this.stt_mau_temlate)
                {
                    case 225: //Gia von
                        this.GrdCt.FieldLayouts[0].Fields["tk_vt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_vt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tk_gv"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_gv"].Settings.CellMaxWidth = 0.0;
                        this.ChkPx_gia_dd.Visibility = Visibility.Hidden;
                        //this.btnChonHDB.Visibility = Visibility.Hidden;
                        //this.tblF5.Visibility = Visibility.Hidden;

                        break;

                    case 226: //Gia ban
                        this.GrdCt.FieldLayouts[0].Fields["tk_dt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_dt"].Settings.CellMaxWidth = 0.0;
                        this.tblF5.Visibility = Visibility.Hidden;
                        //this.btnChonHDB.Visibility = Visibility.Hidden;

                        this.panel40.Visibility = Visibility.Hidden;
                        this.GroupTotal.Visibility = Visibility.Hidden;
                        this.tabHDDT.Visibility = Visibility.Hidden;
                        this.tabItemTTvc.Visibility = Visibility.Hidden;
                        this.tabItemTTgc.Visibility = Visibility.Hidden;
                        this.tabItemTT.Visibility = Visibility.Hidden;
                        
                        break;

                    case 227: //Gia ban Gia von
                        //Gia von
                        this.GrdCt.FieldLayouts[0].Fields["tk_vt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_vt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tk_gv"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_gv"].Settings.CellMaxWidth = 0.0;
                        this.ChkPx_gia_dd.Visibility = Visibility.Hidden;
                        this.tblF5.Visibility = Visibility.Hidden;

                        //Gia ban
                        this.GrdCt.FieldLayouts[0].Fields["tk_dt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tk_dt"].Settings.CellMaxWidth = 0.0;
                        this.ChkSua_tien.Visibility = Visibility.Hidden;
                        //this.btnChonHDB.Visibility = Visibility.Hidden;

                        this.panel40.Visibility = Visibility.Hidden;
                        this.GroupTotal.Visibility = Visibility.Hidden;
                        this.tabHDDT.Visibility = Visibility.Hidden;
                        this.tabItemTTvc.Visibility = Visibility.Hidden;
                        this.tabItemTTgc.Visibility = Visibility.Hidden;
                        this.tabItemTT.Visibility = Visibility.Hidden;
                        break;
                }
                string str = StartupBase.SasObj.GetOption("M_DC_THUE_CK").ToString();
                if (str == "2" || str == "3")
                    return;
                this.lbNhom_hh.Visibility = Visibility.Collapsed;
                this.txtNhom_hh.Visibility = Visibility.Collapsed;
                this.gridlayout50.RowDefinitions[0].Height = new GridLength(0.0);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void LoadData()
        {
            this.GrdLayout00.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            this.GrdCtKM.DataSource = StartUp.CTKMTable.DefaultView;
            this.GrdVTKM.DataSource = StartUp.VattuKMHTHTable.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count == 1)
                this.txtStatus.IsEnabled = false;
            this.UpdateKM();
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;

        }

        private void V_Sau()
        {
            if (FrmSocthda.iRow < StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                ++FrmSocthda.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;
        }

        private void V_Truoc()
        {
            if (FrmSocthda.iRow > 1)
                --FrmSocthda.iRow;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;
        }

        private void V_Dau()
        {
            FrmSocthda.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count > 1 ? 1 : 0;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;
        }

        private void V_Cuoi()
        {
            FrmSocthda.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            if (!string.IsNullOrEmpty((StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"].ToString())))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;
        }

        private void V_Moi()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.Add;
                string stt_rec = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
                this.txtsd_hddt_yn.IsReadOnly = false;
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
                    row1["ma_gd"] = "1";
                    row1["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_nt"];
                    row1["so_seri"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["so_seri"];
                    row1["ma_thue"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_thue"];
                    row1["thue_suat"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["thue_suat"];
                    row1["tk_thue_co"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tk_thue_co"];
                    row1["loai_tk_co"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["loai_tk_co"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_qs"].ToString().Trim());
                }
                if (row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    row1["ty_giaf"] = 1;
                }
                else
                {
                    row1["ty_giaf"] = StartUp.GetRates(row1["ma_nt"].ToString().Trim(), Convert.ToDateTime(row1["ngay_ct"]).Date);
                }
                row1["sd_hddt_yn"] = 0;
                row1["tinh_trang_hddt"] = 0;
                row1["mau_hddt"] = DBNull.Value;
                row1["so_seri_hddt"] = DBNull.Value;
                row1["so_ct_hddt"] = DBNull.Value;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["sua_tien"] = 0;
                row1["px_gia_dd"] = 0;
                row1["sua_tkthue"] = 0;
                row1["sua_thue"] = 0;
                row1["tinh_ck"] = 0;
                row1["t_tien"] = 0;
                row1["t_tien_nt"] = 0;
                row1["t_tien2"] = 0;
                row1["t_tien_nt2"] = 0;
                row1["t_tien_sau_ck"] = 0;
                row1["t_tien_sau_ck_nt"] = 0;
                row1["t_thue"] = 0;
                row1["t_thue_nt"] = 0;
                row1["t_ck"] = 0;
                row1["t_ck_nt"] = 0;
                row1["han_tt"] = 0;
                row1["sl_in"] = 0;
                row1["ma_ctkm"] = "";
                row1["tien_km"] = (object)0;
                row1["tien_km_nt"] = (object)0;
                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                row2["stt_rec"] = stt_rec;
                row2["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), 1);
                row2["ma_ct"] = StartUpTrans.Ma_ct;
                row2["ngay_ct"] = row1["ngay_ct"];
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    row2["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_kho_i"];
                row2["so_luong"] = 0;
                row2["gia_nt"] = 0;
                row2["tien_nt"] = 0;
                row2["gia"] = 0;
                row2["tien"] = 0;
                row2["tl_ck"] = 0;
                row2["ck"] = 0;
                row2["ck_nt"] = 0;
                row2["gia_nt2"] = 0;
                row2["tien_nt2"] = 0;
                row2["gia2"] = 0;
                row2["tien2"] = 0;
                row2["km_ck"] = 0;
                row2["ma_ctkm"] = "";
                row2["pt_km"] = (object)0;
                row2["tien_km"] = (object)0;
                row2["tien_km_nt"] = (object)0;
                row2["khuyen_mai"] = (object)false;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                this.iRow_old = FrmSocthda.iRow;
                FrmSocthda.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                this.txtSoDuKH.Value = 0;
                if (row1["ma_thue"] == null || row1["ma_thue"] == DBNull.Value || row1["ma_thue"].ToString().Trim() == "")
                    row1["ma_thue"] = "";
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_nt"].ToString());
                this.DsVitual = (DataSet)null;
                this.IsInEditMode.Value = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   this.txtMa_kh.SelectAllOnFocus = true;
                   this.txtMa_kh.IsFocus = true;
               }));
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
                int num = (int)ExMessageBox.Show(445, StartupBase.SasObj, "Không có dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1 || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim() != "" && ExMessageBox.Show(150, StartupBase.SasObj, "Đã tạo phiếu thu tự động cho hóa đơn. Có sửa không?", "", MessageBoxButton.YesNo, MessageBoxImage.Asterisk, MessageBoxResult.No) == MessageBoxResult.No)
                    return;
                FormTrans.currActionTask = ActionTask.Edit;
                if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("0"))
                    this.txtsd_hddt_yn.IsReadOnly = true;
                else
                    this.txtsd_hddt_yn.IsReadOnly = false;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                Vnpayrow = StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable().Rows[0];
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                this.IsInEditMode.Value = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   this.txtMa_kh.IsFocus = true;
                   this.txtThue_suat_ValueChanged((object)null, (RoutedPropertyChangedEventArgs<object>)null);
               }));
                this.txtMa_kh.SearchInit();
                this.txtMa_kh_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
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
                        FrmSocthda.iRow = this.iRow_old;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
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
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow].ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
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
            this.txtsd_hddt_yn.IsReadOnly = true;
            StartUp.CTKMTable.Rows.Clear();
            StartUp.VattuKMHTHTable.Rows.Clear();
            this.GrdCtKM.DataSource = StartUp.CTKMTable.DefaultView;
            this.GrdVTKM.DataSource = StartUp.VattuKMHTHTable.DefaultView;
        }

        private void Xoa()
        {
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count == 1)
                return;
            if (FormTrans.currActionTask == ActionTask.None || FormTrans.currActionTask == ActionTask.View)
                FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                StartUpTrans.UpdateTkSd13(1, 0);
                StartUp.DeleteVoucher(_stt_rec, this.txtMa_qs.Text, FormTrans.currActionTask, this.IsNd51);
                StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString());
                StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmSocthda.iRow);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    this.txtNgay_ct.Text = "";
                    FrmSocthda.iRow = FrmSocthda.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmSocthda.iRow - 1 : FrmSocthda.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"].ToString());
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
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()) && ExMessageBox.Show(391, StartupBase.SasObj, "Hóa đơn đã được thanh toán, có muốn xóa phiếu thanh toán hay không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                StartUp.DeletePT(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim());
            this.Xoa();
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
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
                            int num1 = (int)ExMessageBox.Show(460, FormTrans.SasO, "Ngày bắt đầu sử dụng quyển sổ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = true;
                            break;
                        case 2:
                            int num2 = (int)ExMessageBox.Show(465, FormTrans.SasO, "Quyền sử dụng quyển sổ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                            int num3 = (int)ExMessageBox.Show(470, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            return;
                        }
                        new Thread((ThreadStart)(() => this.Post(0))).Start();
                    }
                }
                this.txtsd_hddt_yn.IsReadOnly = true;
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
                this.Sum_all();
                this.UpdateKM();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = "1";
                if (this.CheckValid())
                {
                    if (!this.IsSequenceSave)
                    {
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = StartupBase.SasObj.GetOption("M_MA_DVCS").ToString();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_act"] = StartUpTrans.tbStatus.Select("ma_post =" + this.txtStatus.SelectedIndex.ToString())[0]["ten_act"];
                    }
                    DateTime dateTime = (DateTime)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    object obj = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["px_gia_dd"];
                    string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                    if (dateTime >= StartUp.ngay_gia_px && obj.ToString() != "1")
                    {
                        SqlCommand sqlcmd = new SqlCommand("Ingia_px");
                        sqlcmd.CommandType = CommandType.StoredProcedure;
                        sqlcmd.Parameters.Add("@Ngay_ct", SqlDbType.SmallDateTime).Value = dateTime;
                        sqlcmd.Parameters.Add("@Ma_kho", SqlDbType.VarChar).Value = "";
                        sqlcmd.Parameters.Add("@Ma_vt", SqlDbType.VarChar).Value = "";
                        sqlcmd.Parameters.Add("@So_luong", SqlDbType.Decimal).Value = 0;
                        for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
                        {
                            string str2 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_ton"].ToString().Trim();
                            if (str2.Equals("1") || str2.Equals("4"))
                            {
                                string str3 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString();
                                string str4 = StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString();
                                Decimal num = (Decimal)StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"];
                                sqlcmd.Parameters["@Ma_kho"].Value = str3;
                                sqlcmd.Parameters["@Ma_vt"].Value = str4;
                                sqlcmd.Parameters["@So_luong"].Value = num;
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
                    }
                    if (StartupBase.M_MA_NT0 == str1)
                    {
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            dataRowView["gia"] = dataRowView["gia_nt"];
                            dataRowView["tien"] = dataRowView["tien_nt"];
                            dataRowView["gia2"] = dataRowView["gia_nt2"];
                            dataRowView["tien2"] = dataRowView["tien_nt2"];
                            dataRowView["ck"] = dataRowView["ck_nt"];
                        }
                        DataRowView dataRowView1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
                        dataRowView1["t_tien2"] = dataRowView1["t_tien_nt2"];
                        dataRowView1["t_ck"] = dataRowView1["t_ck_nt"];
                        dataRowView1["t_tien_sau_ck"] = dataRowView1["t_tien_sau_ck_nt"];
                        dataRowView1["t_thue"] = dataRowView1["t_thue_nt"];
                        dataRowView1["t_tt"] = dataRowView1["t_tt_nt"];
                    }
                    DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[1].Clone();
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                    {
                        if (!this.IsSequenceSave)
                            this.PhanBoThueInCT_CBTien();
                        LocalTable1 = StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Copy();
                    }
                    DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[0].Clone();
                    LocalTable2.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                    if (!this.IsSequenceSave)
                        LocalTable2.Rows[0]["status"] = 0;
                    if (LocalTable2.Rows[0]["ten_kh_thue"].ToString().Trim() != "")
                        LocalTable2.Rows[0]["ten_kh"] = LocalTable2.Rows[0]["ten_kh_thue"];

                    LocalTable2.Rows[0]["order_id"] = DateTime.Now.Ticks;
                    Vnpayrow = LocalTable2.Rows[0];
                    DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable2, "stt_rec;row_id");
                    for (int index = 0; index < LocalTable1.Rows.Count; ++index)
                    {
                        if (LocalTable1.Rows[index]["dvt1"].ToString().Trim() == LocalTable1.Rows[index]["dvt"].ToString().Trim())
                        {
                            LocalTable1.Rows[index]["he_so1"] = 0;
                            LocalTable1.Rows[index]["so_luong1"] = 0;
                        }
                        if (String.IsNullOrEmpty(LocalTable1.Rows[index]["he_so1"].ToString()))
                            LocalTable1.Rows[index]["he_so1"] = new Decimal(0);
                        if ((Decimal)LocalTable1.Rows[index]["he_so1"] != new Decimal(0))
                        {
                            LocalTable1.Rows[index]["so_luong1"] = (Decimal)(Convert.ToDecimal(LocalTable1.Rows[index]["so_luong"]) * Convert.ToDecimal(LocalTable1.Rows[index]["he_so1"]));
                        }
                    }
                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable1, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num1 = (int)ExMessageBox.Show(475, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    else
                    {
                        bool flag1 = false;
                        if (!this.IsSequenceSave)
                        {
                            if (!flag1)
                            {
                                this.dsCheckData = StartUp.CheckData(FormTrans.currActionTask == ActionTask.Edit ? 0 : 1, this.stt_mau_temlate);
                                this.dsCheckData.Tables[0].AcceptChanges();
                                if (this.dsCheckData.Tables.Count > 0)
                                {
                                    string str2 = "";
                                    foreach (DataRowView dataRowView in this.dsCheckData.Tables[0].DefaultView)
                                    {
                                        if (!flag1)
                                        {
                                            switch (dataRowView[0].ToString())
                                            {
                                                case "PH01":
                                                    if (StartUpTrans.M_trung_so.Equals("1"))
                                                    {
                                                        if (ExMessageBox.Show(480, StartupBase.SasObj, "Có chứng từ trùng số. Số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
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
                                                        int num2 = (int)ExMessageBox.Show(485, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag1 = true;
                                                        break;
                                                    }
                                                    break;
                                                case "PH02":
                                                    if (FormTrans.currActionTask != ActionTask.Edit && this.dsCheckData.Tables.Count > 1)
                                                    {
                                                        int int32 = Convert.ToInt32(this.dsCheckData.Tables[1].Rows[0][0]);
                                                        this.dsCheckData.Tables[2].Rows[0][0].ToString().Trim();
                                                        Convert.ToInt32(this.dsCheckData.Tables[3].Rows[0][0]);
                                                        switch (int32)
                                                        {
                                                            case 1:
                                                                int num2 = (int)ExMessageBox.Show(490, FormTrans.SasO, "Ngày bắt đầu sử dụng ký hiệu hóa đơn không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtNgay_ct.Focus();
                                                                flag1 = true;
                                                                break;
                                                            case 2:
                                                                int num3 = (int)ExMessageBox.Show(495, FormTrans.SasO, "Quyền sử dụng ký hiệu hóa đơn không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtMa_qs.IsFocus = true;
                                                                flag1 = true;
                                                                break;
                                                            case 3:
                                                                string lower = this.dsCheckData.Tables[4].Rows[0]["ten_tthd"].ToString().Trim().ToLower();
                                                                int num4 = (int)ExMessageBox.Show(500, FormTrans.SasO, "Số hóa đơn của ký hiệu [" + this.txtMa_qs.Text.Trim() + "] đã [" + lower + "]!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                                                                flag1 = true;
                                                                break;
                                                        }
                                                        break;
                                                    }
                                                    break;
                                                case "PH03":
                                                    int num5 = (int)ExMessageBox.Show(505, StartupBase.SasObj, "Số hóa đơn không liên tục!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                                                    break;
                                                case "PH04":
                                                    int num6 = (int)ExMessageBox.Show(510, StartupBase.SasObj, "Mã nx là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtTk_thue_no.IsFocus = true;
                                                    break;
                                                case "PH05":
                                                    int num7 = (int)ExMessageBox.Show(621, StartupBase.SasObj, "Tài khoản thuế đối ứng là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtTk_thue_co.IsFocus = true;
                                                    break;
                                                case "PH06":
                                                    int num8 = (int)ExMessageBox.Show(622, StartupBase.SasObj, "Tài khoản thuế là tk tổng hợp!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtMa_nx.IsFocus = true;
                                                    break;
                                                case "CT01":
                                                    int int16_1 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num9 = (int)ExMessageBox.Show(515, StartupBase.SasObj, "Tk dt là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_1] as DataRecord).Cells["tk_dt"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                                    break;
                                                case "CT02":
                                                    int int16_2 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num10 = (int)ExMessageBox.Show(520, StartupBase.SasObj, "Tk kho là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_2] as DataRecord).Cells["tk_vt"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                                    return;
                                                case "CT03":
                                                    StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ton13"] = InFuncLib.GetTon13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_kho_i"].ToString(), StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString(), StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vv_i"].ToString());
                                                    if (this.ParseInt(StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["vt_ton_kho"], 0) == 1 && !str2.Contains(StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString().Trim()))
                                                    {
                                                        str2 = str2 + StartUpTrans.DsTrans.Tables[1].DefaultView[(int)Convert.ToInt16(dataRowView[1])]["ma_vt"].ToString().Trim() + ", ";
                                                        break;
                                                    }
                                                    break;
                                                case "CT04":
                                                    int int16_3 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num11 = (int)ExMessageBox.Show(535, StartupBase.SasObj, "Tk gv là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_3] as DataRecord).Cells["tk_gv"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                                    break;
                                                case "CT05":
                                                    if (StartUp.M_AR_CK == 1)
                                                    {
                                                        int int16_4 = (int)Convert.ToInt16(dataRowView[1]);
                                                        int num2 = (int)ExMessageBox.Show(540, StartupBase.SasObj, "Tk c.khấu là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        flag1 = true;
                                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_4] as DataRecord).Cells["tk_ck"];
                                                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                                        break;
                                                    }
                                                    break;
                                                case "CT06":
                                                    int int16_5 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num12 = (int)ExMessageBox.Show(545, StartupBase.SasObj, "Tk cp km là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_5] as DataRecord).Cells["tk_km_i"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                                                    break;
                                            }
                                            this.dsCheckData.Tables[0].Rows.Remove(dataRowView.Row);
                                        }
                                        else
                                            break;
                                    }
                                    if (!string.IsNullOrEmpty(str2))
                                    {
                                        if (StartUp.M_CHK_TON_VT.Equals("2"))
                                        {
                                            int num2 = (int)ExMessageBox.Show(525, StartupBase.SasObj, "Có vật tư [" + str2.Substring(0, str2.Length - 2) + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            flag1 = true;
                                        }
                                        else if (StartUp.M_CHK_TON_VT.Equals("1"))
                                        {
                                            int num13 = (int)ExMessageBox.Show(530, StartupBase.SasObj, "Có vật tư [" + str2.Substring(0, str2.Length - 2) + "] xuất âm hoặc tồn kho nhỏ hơn tồn tối thiếu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        }
                                    }
                                }
                            }
                            if (!flag1 && this.IsNd51)
                            {
                                if (this.txtMa_qs.RowResult == null)
                                    this.txtMa_qs.SearchInit();
                                this.UpdateNewSoCt(FormTrans.SasO, this.txtMa_qs.Text);
                                if (FormTrans.currActionTask == ActionTask.Edit && this.txtMa_qs.RowResult != null)
                                    this.UpdateNewNgayCt(FormTrans.SasO, this.txtMa_qs.Text, this.GetCurrentSo_ct(this.txtMa_qs.RowResult["transform"].ToString(), this.txtSo_ct.Text.Trim()));
                            }
                        }
                        if (!flag1)
                        {
                            string format = "EXEC  {0} '" + this.txtMa_qs.Text.Trim() + "', '" + this.txtSo_ct.Text.Trim() + "'";
                            this.BindingSasObj.ExcuteNonQuery(new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 2 ? string.Format(format, "SetSoct") : string.Format(format, StartUpTrans.Process_Store[2])));
                            bool _createPT1 = false;
                            string newstt_recPt1 = "";
                            DataTable dt = new DataTable();
                            dt.Columns.Add("ma_ct", typeof(string));
                            dt.Columns.Add("stt_rec", typeof(string));
                            dt.Columns.Add("stt_recPT", typeof(string));
                            dt.Columns.Add("ma_qs", typeof(string));
                            dt.Columns.Add("so_ct", typeof(string));
                            dt.Columns.Add("ma_nt", typeof(string));
                            dt.Columns.Add("ty_gia", typeof(Decimal));
                            dt.Columns.Add("ty_giaf", typeof(Decimal));
                            dt.Columns.Add("nguoinop", typeof(string));
                            dt.Columns.Add("lydonop", typeof(string));
                            dt.Columns.Add("ma_gd", typeof(string));
                            if (SysFunc.CheckPermission(this.BindingSasObj, ActionTask.Add, this.BindingSasObj.ExcuteScalar(new SqlCommand("Select top 1 menu_id From command Where ma_ct like 'PT1'")).ToString().Trim()) && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("2"))
                            {
                                bool flag2 = false;
                                string str2 = this.BindingSasObj.GetOption("M_TK_TK_VT").ToString();
                                char[] chArray = new char[1] { ',' };
                                foreach (string str3 in str2.Split(chArray))
                                {
                                    if (!string.IsNullOrEmpty(str3.ToString().Trim()) && this.txtMa_nx.Text.Trim().StartsWith(str3.ToString().Trim()))
                                        flag2 = true;
                                }
                                string sma_ct_pt = string.Empty;
                                if (!flag2 && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()))
                                {
                                    newstt_recPt1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim();
                                    sma_ct_pt = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim();
                                    this.DeleteVoucherPT(newstt_recPt1, sma_ct_pt);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = "";
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                                }
                                if (flag2)
                                {
                                    DataTable dataTable = (DataTable)null;
                                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString()))
                                    {
                                        switch (this.BindingSasObj.GetOption("M_TAO_PT_TM").ToString())
                                        {
                                            case "1":
                                                if (ExMessageBox.Show(696, StartupBase.SasObj, "Có tạo phiếu thu tiền ngay cho hóa đơn bán hàng?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                                                {
                                                    _createPT1 = true;
                                                    break;
                                                }
                                                break;
                                            case "2":
                                                _createPT1 = true;
                                                break;
                                        }

                                        if (_createPT1)
                                        {
                                            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString()))
                                            {
                                                SqlCommand sqlcmd = new SqlCommand();
                                                sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai FROM {0} WHERE stt_rec LIKE '{1}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Equals("PT1") ? "ph41" : "ph51", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString());
                                                dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];
                                            }
                                            int num2;
                                            if (this.txtT_tt.nValue > Convert.ToDecimal(this.BindingSasObj.GetSysvar("M_MUC_TIEN_PT1")))
                                            {
                                                num2 = 1;
                                            }
                                            else
                                            {
                                                SqlCommand sqlcmd = new SqlCommand();
                                                sqlcmd.CommandText = "SELECT COUNT(1) FROM dmtknh WHERE tk LIKE @tk";
                                                sqlcmd.Parameters.Add(new SqlParameter("@tk", SqlDbType.VarChar)).Value = this.txtMa_nx.Text.Trim();
                                                num2 = (int)this.BindingSasObj.ExcuteScalar(sqlcmd);
                                            }
                                            DataRow row = dt.NewRow();
                                            dt.Rows.Add(row);
                                            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim() == "")
                                            {
                                                FrmTaoPT frmTaoPt = new FrmTaoPT();
                                                frmTaoPt.tbInfoPT = dataTable;
                                                frmTaoPt.DataContext = dt.DefaultView;
                                                frmTaoPt.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"].ToString();
                                                frmTaoPt.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"].ToString().Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                                frmTaoPt.txtnguoi_nop.Text = this.txtOng_ba.Text;
                                                frmTaoPt.kind = num2 > 0 ? 2 : 1;
                                                frmTaoPt.Ma_nt_ht = this.txtMa_nt.Text;
                                                frmTaoPt.so_hd = this.txtSo_ct.Text.Trim();
                                                frmTaoPt.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
                                                frmTaoPt.filterma_qs = this.txtMa_qs.Filter;
                                                frmTaoPt.ShowDialog();
                                                if (!frmTaoPt.isOk)
                                                {
                                                    _createPT1 = false;
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = "";
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                                                }
                                                else
                                                {
                                                    dt = dt.Copy();
                                                    newstt_recPt1 = !frmTaoPt.txtKind.Text.Equals("1") ? DataProvider.NewTrans(StartupBase.SasObj, "BC1", StartUpTrans.Ws_Id) : DataProvider.NewTrans(StartupBase.SasObj, "PT1", StartUpTrans.Ws_Id);
                                                    dt.Rows[0]["ma_ct"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                                    dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                                    dt.Rows[0]["stt_recPT"] = newstt_recPt1;
                                                    dt.Rows[0]["ma_qs"] = frmTaoPt.txtMa_qs_pt.Text;
                                                    dt.Rows[0]["so_ct"] = frmTaoPt.txtso_ct_pt.Text.PadLeft(frmTaoPt.txtso_ct_pt.MaxLength, ' ');
                                                    dt.Rows[0]["ma_nt"] = frmTaoPt.txtMa_nt.Text;
                                                    dt.Rows[0]["ty_gia"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.Rate;
                                                    dt.Rows[0]["ty_giaf"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.RateF;
                                                    dt.Rows[0]["nguoinop"] = frmTaoPt.txtnguoi_nop.Text;
                                                    dt.Rows[0]["lydonop"] = frmTaoPt.txtlydo_nop.Text;
                                                    dt.Rows[0]["ma_gd"] = frmTaoPt.txtMa_gd.Text;
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = frmTaoPt.txtMa_qs_pt.Text;
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = frmTaoPt.txtso_ct_pt.Text.Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = newstt_recPt1;
                                                }
                                            }
                                        }

                                    }
                                    else
                                    {
                                        newstt_recPt1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim();
                                        sma_ct_pt = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim();
                                        SqlCommand sqlcmd = new SqlCommand();
                                        sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai,ty_gia,ty_giaf FROM {0} WHERE stt_rec LIKE '{1}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Equals("PT1") ? "ph41" : "ph51", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString());
                                        dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];

                                        DataRow row = dt.NewRow();
                                        dt.Rows.Add(row);
                                        if (dataTable != null && dataTable.Rows.Count > 0)
                                        {
                                            _createPT1 = true;
                                            dt.Rows[0]["ma_ct"] = dataTable.Rows[0]["ma_ct"].ToString().Trim();
                                            dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                            dt.Rows[0]["stt_recPT"] = dataTable.Rows[0]["stt_rec"].ToString().Trim();
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
                                            this.DeleteVoucherPT(newstt_recPt1, sma_ct_pt);

                                            switch (this.BindingSasObj.GetOption("M_TAO_PT_TM").ToString())
                                            {
                                                case "1":
                                                    if (ExMessageBox.Show(696, StartupBase.SasObj, "Có tạo phiếu thu tiền ngay cho hóa đơn bán hàng?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                                                    {
                                                        _createPT1 = true;
                                                        break;
                                                    }
                                                    break;
                                                case "2":
                                                    _createPT1 = true;
                                                    break;
                                            }

                                            if (_createPT1)
                                            {

                                                SqlCommand sqlcmd1 = new SqlCommand();
                                                sqlcmd1.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai FROM {0} WHERE stt_rec LIKE '{1}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Equals("PT1") ? "ph41" : "ph51", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString());
                                                dataTable = this.BindingSasObj.ExcuteReader(sqlcmd1).Tables[0];

                                                int num2;
                                                if (this.txtT_tt.nValue > Convert.ToDecimal(this.BindingSasObj.GetSysvar("M_MUC_TIEN_PT1")))
                                                {
                                                    num2 = 1;
                                                }
                                                else
                                                {
                                                    SqlCommand sqlcmd2 = new SqlCommand();
                                                    sqlcmd2.CommandText = "SELECT COUNT(1) FROM dmtknh WHERE tk LIKE @tk";
                                                    sqlcmd2.Parameters.Add(new SqlParameter("@tk", SqlDbType.VarChar)).Value = this.txtMa_nx.Text.Trim();
                                                    num2 = (int)this.BindingSasObj.ExcuteScalar(sqlcmd2);
                                                }
                                                DataRow row2 = dt.NewRow();
                                                dt.Rows.Add(row2);
                                                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim() == "")
                                                {
                                                    FrmTaoPT frmTaoPt = new FrmTaoPT();
                                                    frmTaoPt.tbInfoPT = dataTable;
                                                    frmTaoPt.DataContext = dt.DefaultView;
                                                    frmTaoPt.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"].ToString();
                                                    frmTaoPt.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"].ToString().Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                                    frmTaoPt.txtnguoi_nop.Text = this.txtOng_ba.Text;
                                                    frmTaoPt.kind = num2 > 0 ? 2 : 1;
                                                    frmTaoPt.Ma_nt_ht = this.txtMa_nt.Text;
                                                    frmTaoPt.so_hd = this.txtSo_ct.Text.Trim();
                                                    frmTaoPt.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
                                                    frmTaoPt.filterma_qs = this.txtMa_qs.Filter;
                                                    frmTaoPt.ShowDialog();
                                                    if (!frmTaoPt.isOk)
                                                    {
                                                        _createPT1 = false;
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = "";
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                                                    }
                                                    else
                                                    {
                                                        dt = dt.Copy();
                                                        newstt_recPt1 = !frmTaoPt.txtKind.Text.Equals("1") ? DataProvider.NewTrans(StartupBase.SasObj, "BC1", StartUpTrans.Ws_Id) : DataProvider.NewTrans(StartupBase.SasObj, "PT1", StartUpTrans.Ws_Id);
                                                        dt.Rows[0]["ma_ct"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                                        dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                                        dt.Rows[0]["stt_recPT"] = newstt_recPt1;
                                                        dt.Rows[0]["ma_qs"] = frmTaoPt.txtMa_qs_pt.Text;
                                                        dt.Rows[0]["so_ct"] = frmTaoPt.txtso_ct_pt.Text.PadLeft(frmTaoPt.txtso_ct_pt.MaxLength, ' ');
                                                        dt.Rows[0]["ma_nt"] = frmTaoPt.txtMa_nt.Text;
                                                        dt.Rows[0]["ty_gia"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.Rate;
                                                        dt.Rows[0]["ty_giaf"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.RateF;
                                                        dt.Rows[0]["nguoinop"] = frmTaoPt.txtnguoi_nop.Text;
                                                        dt.Rows[0]["lydonop"] = frmTaoPt.txtlydo_nop.Text;
                                                        dt.Rows[0]["ma_gd"] = frmTaoPt.txtMa_gd.Text;
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = frmTaoPt.txtMa_qs_pt.Text;
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = frmTaoPt.txtso_ct_pt.Text.Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = newstt_recPt1;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                                //Cách củ
                                //if (flag2 && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim() == "")
                                //{
                                //    switch (this.BindingSasObj.GetOption("M_TAO_PT_TM").ToString())
                                //    {
                                //        case "1":
                                //            if (ExMessageBox.Show(696, StartupBase.SasObj, "Có tạo phiếu thu tiền ngay cho hóa đơn bán hàng?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                                //            {
                                //                _createPT1 = true;
                                //                break;
                                //            }
                                //            break;
                                //        case "2":
                                //            _createPT1 = true;
                                //            break;
                                //    }
                                //    if (_createPT1)
                                //    {
                                //        DataTable dataTable = (DataTable)null;
                                //        if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString()))
                                //        {
                                //            SqlCommand sqlcmd = new SqlCommand();
                                //            sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai FROM {0} WHERE stt_rec LIKE '{1}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Equals("PT1") ? "ph41" : "ph51", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString());
                                //            dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];
                                //        }
                                //        int num2;
                                //        if (this.txtT_tt.nValue > Convert.ToDecimal(this.BindingSasObj.GetSysvar("M_MUC_TIEN_PT1")))
                                //        {
                                //            num2 = 1;
                                //        }
                                //        else
                                //        {
                                //            SqlCommand sqlcmd = new SqlCommand();
                                //            sqlcmd.CommandText = "SELECT COUNT(1) FROM dmtknh WHERE tk LIKE @tk";
                                //            sqlcmd.Parameters.Add(new SqlParameter("@tk", SqlDbType.VarChar)).Value = this.txtMa_nx.Text.Trim();
                                //            num2 = (int)this.BindingSasObj.ExcuteScalar(sqlcmd);
                                //        }
                                //        DataRow row = dt.NewRow();
                                //        dt.Rows.Add(row);
                                //        if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim() == "")
                                //        {
                                //            FrmTaoPT frmTaoPt = new FrmTaoPT();
                                //            frmTaoPt.tbInfoPT = dataTable;
                                //            frmTaoPt.DataContext = dt.DefaultView;
                                //            frmTaoPt.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"].ToString();
                                //            frmTaoPt.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"].ToString().Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                //            frmTaoPt.txtnguoi_nop.Text = this.txtOng_ba.Text;
                                //            frmTaoPt.kind = num2 > 0 ? 2 : 1;
                                //            frmTaoPt.Ma_nt_ht = this.txtMa_nt.Text;
                                //            frmTaoPt.so_hd = this.txtSo_ct.Text.Trim();
                                //            frmTaoPt.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
                                //            frmTaoPt.filterma_qs = this.txtMa_qs.Filter;
                                //            frmTaoPt.ShowDialog();
                                //            if (!frmTaoPt.isOk)
                                //            {
                                //                _createPT1 = false;
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = "";
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                                //            }
                                //            else
                                //            {
                                //                dt = dt.Copy();
                                //                newstt_recPt1 = !frmTaoPt.txtKind.Text.Equals("1") ? DataProvider.NewTrans(StartupBase.SasObj, "BC1", StartUpTrans.Ws_Id) : DataProvider.NewTrans(StartupBase.SasObj, "PT1", StartUpTrans.Ws_Id);
                                //                dt.Rows[0]["ma_ct"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                //                dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                //                dt.Rows[0]["stt_recPT"] = newstt_recPt1;
                                //                dt.Rows[0]["ma_qs"] = frmTaoPt.txtMa_qs_pt.Text;
                                //                dt.Rows[0]["so_ct"] = frmTaoPt.txtso_ct_pt.Text.PadLeft(frmTaoPt.txtso_ct_pt.MaxLength, ' ');
                                //                dt.Rows[0]["ma_nt"] = frmTaoPt.txtMa_nt.Text;
                                //                dt.Rows[0]["ty_gia"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.Rate;
                                //                dt.Rows[0]["ty_giaf"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? 1 : this.txtTy_gia.RateF;
                                //                dt.Rows[0]["nguoinop"] = frmTaoPt.txtnguoi_nop.Text;
                                //                dt.Rows[0]["lydonop"] = frmTaoPt.txtlydo_nop.Text;
                                //                dt.Rows[0]["ma_gd"] = frmTaoPt.txtMa_gd.Text;
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = frmTaoPt.txtMa_qs_pt.Text;
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = frmTaoPt.txtso_ct_pt.Text.Trim().PadLeft(frmTaoPt.txtso_ct_pt.MaxLength);
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = frmTaoPt.txtKind.Text.Equals("1") ? "PT1" : "BC1";
                                //                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = newstt_recPt1;
                                //            }
                                //        }
                                //    }
                                //}
                            }
                            string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                            new Thread((ThreadStart)(() =>
                           {
                               this.Post(_createPT1 ? 1 : 0);
                               if (_createPT1)
                                   this.CreatePT1(dt);
                               if (this.IsSequenceSave)
                                   return;
                               this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate)new Action(() =>
                 {
                     if (StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(_stt_rec1))
                     {
                         this.UpdateTonKho();
                         this.LoadDataDu13();
                     }
                     if (!_createPT1 || string.IsNullOrEmpty(newstt_recPt1))
                         return;
                     DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + _stt_rec1 + "'");
                     if (dataRowArray.Length == 1)
                     {
                         dataRowArray[0]["stt_rec_pt"] = newstt_recPt1;
                         dataRowArray[0]["so_ct_pt"] = dt.Rows[0]["so_ct"].ToString().Trim().PadLeft(this.so_ct_pt_length);
                         dataRowArray[0]["ma_ct_pt"] = dt.Rows[0]["ma_ct"];
                     }
                 }));
                           })).Start();
                            if (!this.IsSequenceSave)
                            {
                                int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                if (FrmSocthda.iRow != pos)
                                {
                                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                    DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                    row2.ItemArray = row1.ItemArray;
                                    if (FrmSocthda.iRow > pos)
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                    else
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    FrmSocthda.iRow = pos;
                                }
                                FormTrans.currActionTask = ActionTask.None;
                                this.IsInEditMode.Value = false;

                            }
                            this.txtsd_hddt_yn.IsReadOnly = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void CreatePT1(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[SOCTHDA-CREATEPT1] @Stt_rec, @Stt_recPT, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct");
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = dt.Rows[0]["stt_rec"];
                sqlcmd.Parameters.Add("@Stt_recPT", SqlDbType.VarChar).Value = dt.Rows[0]["stt_recPT"];
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

        private void Post(int ispostck)
        {
            string format = "exec [dbo].{0} @stt_rec,0,@IsHasPT1";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, "[SOCTHDA-Post]") : string.Format(format, StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@IsHasPT1", SqlDbType.Int).Value = ispostck;
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
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(555, StartupBase.SasObj, "Chưa vào mã khách hàng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtMa_kh.IsFocus = true;
                    }
                    if (StartUp.M_SD_HDDT.Equals("1") && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sd_hddt_yn"].ToString() == "1" && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() != string.Empty && flag)
                    {
                        SqlCommand sqlcmd = new SqlCommand();
                        sqlcmd.CommandText = "SELECT * FROM khhddt WHERE ma_kh='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() + "' AND status in (1,2)";
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count == 0)
                        {
                            int num = (int)ExMessageBox.Show(999, StartupBase.SasObj, "Khách hàng [" + (StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kh.RowResult["ten_kh"].ToString().Trim() : this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim()) + "] chưa cập nhật khách hàng hóa đơn điện tử!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtMa_kh.IsFocus = true;
                        }
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(560, StartupBase.SasObj, "Chưa vào mã nx!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtMa_nx.IsFocus = true;
                    }
                    if ((this.txtNgay_ct.Value == null || this.txtNgay_ct.Value.ToString() == "") && flag)
                    {
                        int num = (int)ExMessageBox.Show(570, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtNgay_ct.Focus();
                    }
                    if (this.txtNgay_ct.Value.ToString() != "" && flag)
                    {
                        if (!this.txtNgay_ct.IsValueValid && flag)
                        {
                            int num = (int)ExMessageBox.Show(575, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                        if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(Convert.ToDateTime(this.txtNgay_ct.dValue))) && flag)
                        {
                            int num = (int)ExMessageBox.Show(580, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtNgay_ct.Focus();
                        }
                        if (flag && Convert.ToDateTime(this.txtNgay_ct.dValue) < NgayTC.GetStartDate(StartUp.M_ngay_ct0))
                        {
                            int num = (int)ExMessageBox.Show(585, StartupBase.SasObj, "Ngày hạch toán phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                    goto label_25;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_25:
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
                            int num = (int)ExMessageBox.Show(590, StartupBase.SasObj, "Chưa vào ngày lập px!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtNgay_lct.Focus();
                            return false;
                        }
                        if (!this.txtNgay_lct.IsValueValid)
                        {
                            int num = (int)ExMessageBox.Show(595, StartupBase.SasObj, "Ngày lập px không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtNgay_lct.Focus();
                            return false;
                        }
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()) && flag)
                    {
                        int num = (int)ExMessageBox.Show(1223, StartupBase.SasObj, "Chưa vào quyển c.từ", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        this.txtMa_qs.IsFocus = true;
                    }
                    if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0 && flag)
                    {
                        int num = (int)ExMessageBox.Show(620, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        flag = false;
                        SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                        this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                        this.GrdCt.ActiveCell = (this.GrdCt.Records[0] as DataRecord).Cells["ma_vt"];
                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                    }
                    for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count && flag; ++index)
                    {
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_ct"] = StartUpTrans.Ma_ct;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_vt"].ToString().Trim()))
                        {
                            int num = (int)ExMessageBox.Show(625, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_vt"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                        if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ma_kho_i"].ToString().Trim()))
                        {
                            int num = (int)ExMessageBox.Show(630, StartupBase.SasObj, "Chưa vào chi tiết vật tư, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["ma_kho_i"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                        if (flag && StartUpTrans.DsTrans.Tables[1].DefaultView[index]["km_ck"].ToString().Trim() != "1" && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_dt"].ToString().Trim()))
                        {
                            if (this.stt_mau_temlate != 226 && this.stt_mau_temlate != 227)
                            {
                                int num = (int)ExMessageBox.Show(645, StartupBase.SasObj, "Chưa vào tk doanh thu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag = false;
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_dt"];
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                            }  
                        }
                        if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_gv"].ToString().Trim()))
                        {
                            if(this.stt_mau_temlate != 225 && this.stt_mau_temlate != 227)
                            {
                                int num = (int)ExMessageBox.Show(665, StartupBase.SasObj, "Chưa vào tk giá vốn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag = false;
                                this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_gv"];
                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                            }  
                        }
                        //if (flag && string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tk_km_i"].ToString().Trim()) && (StartUp.M_KM_CK == 1 && StartUpTrans.DsTrans.Tables[1].DefaultView[index]["km_ck"].ToString().Trim() == "1"))
                        //{
                        //    int num = (int)ExMessageBox.Show(685, StartupBase.SasObj, "Chưa vào tk cp km!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        //    flag = false;
                        //    this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_km_i"];
                        //    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        //}
                        if (flag && int.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_ton"].ToString()) == 3 && Decimal.Parse(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"].ToString()) == new Decimal(0))
                        {
                            int num = (int)ExMessageBox.Show(695, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["so_luong"];
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.Focus()));
                        }
                    }
                    this.txtMa_kh.SearchInit();
                    if (flag)
                    {
                        if (this.txtMa_kh.RowResult == null)
                            this.txtMa_kh.SearchInit();
                        if (this.txtMa_kh.RowResult != null && (string.IsNullOrEmpty(this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim()) || string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()) || string.IsNullOrEmpty(this.txtMa_kh.RowResult["ten_kh"].ToString().Trim())) && flag)
                        {
                            FrmKHInfo frmKhInfo = new FrmKHInfo();
                            frmKhInfo.ShowDialog();
                            string mst = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"].ToString().Trim();
                            if (frmKhInfo.isError && mst != "" && !SysFunc.CheckSumMaSoThue(mst))
                            {
                                if (StartUpTrans.M_MST_CHECK.Equals("1"))
                                {
                                    int num1 = (int)ExMessageBox.Show(705, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                                else
                                {
                                    int num2 = (int)ExMessageBox.Show(710, StartupBase.SasObj, "Mã số thuế không hợp lệ, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    flag = false;
                                    this.txtma_so_thue.Focus();
                                }
                            }
                            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0][StartUpTrans.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"].ToString().Trim() == "")
                            {
                                this.txtMa_kh.SearchInit();
                                this.txtMa_kh_PreviewLostFocus((object)this.txtMa_kh, (KeyboardFocusChangedEventArgs)null);
                            }
                        }
                    }
                    if (flag && this.txtTk_thue_no.Text == "")
                    {
                        if(this.stt_mau_temlate != 226 && this.stt_mau_temlate != 227)
                        {
                            int num = (int)ExMessageBox.Show(725, StartupBase.SasObj, "Chưa vào tk thuế!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtTk_thue_no.IsFocus = true;
                        }
                    }
                    if (flag && !this.txtTk_thue_no.CheckLostFocus())
                    {
                        if(this.stt_mau_temlate != 226 && this.stt_mau_temlate != 227)
                        {
                            int num = (int)ExMessageBox.Show(730, StartupBase.SasObj, "Tk thuế không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtTk_thue_no.IsFocus = true;
                        }
                    }
                    if (flag && this.txtTk_thue_co.Text == "")
                    {
                        if(this.stt_mau_temlate != 226 && this.stt_mau_temlate != 227)
                        {
                            int num = (int)ExMessageBox.Show(735, StartupBase.SasObj, "Chưa vào tk thuế!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtMa_thue.IsFocus = true;
                        }
                    }
                    if (flag && !this.txtTk_thue_co.CheckLostFocus())
                    {
                        if(this.stt_mau_temlate != 226 && this.stt_mau_temlate != 227)
                        {
                            int num = (int)ExMessageBox.Show(740, StartupBase.SasObj, "Tk thuế không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtTk_thue_co.IsFocus = true;
                        }
                    }
                    if (!this.IsNd51)
                    {
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()) && flag)
                        {
                            int num = (int)ExMessageBox.Show(745, StartupBase.SasObj, "Chưa vào số chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                            this.txtSo_ct.Focus();
                        }
                    }
                    else
                    {
                        if (flag)
                            this.txtMa_qs.SearchInit();
                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString()) && flag)
                        {
                            int num = (int)ExMessageBox.Show(770, StartupBase.SasObj, "Chưa vào số hóa đơn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag = false;
                            this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                        }
                        if (flag && !this.CheckSo_ct(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["transform"].ToString(), this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["so_ct1"], new Decimal(0)), this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["so_ct2"], new Decimal(0)), this.txtSo_ct.Text.Trim()))
                        {
                            int num = (int)ExMessageBox.Show(775, StartupBase.SasObj, "Số hóa đơn không thuộc ký hiệu hiện hành!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
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
            this.txtsd_hddt_yn.IsReadOnly = false;
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
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow].ItemArray;
                if (StartUp.M_SD_HDDT.Equals("0"))
                    row1["sd_hddt_yn"] = 0;
                row1["tinh_trang_hddt"] = 0;
                row1["mau_hddt"] = DBNull.Value;
                row1["so_seri_hddt"] = DBNull.Value;
                row1["so_ct_hddt"] = DBNull.Value;
                row1["sd_hddt_yn"] = 0;
                row1["stt_rec"] = stt_rec;
                row1["stt_rec_pt"] = "";
                row1["so_ct_pt"] = "";
                row1["ma_ct_pt"] = "";
                row1["ma_qs_pt"] = "";

                row1["stt_rec_px"] = "";
                row1["so_ct_px"] = "";
                row1["ma_ct_px"] = "";
                row1["ma_qs_px"] = "";
                row1["loai_xnvl"] = 0;

                row1["ma_gd"] = "1";
                row1["ten_kh_thue"] = "";
                row1["ngay_ct"] = this._formcopy.ngay_ct;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["ten_post"] = StartUpTrans.tbStatus.Select("ma_post =" + StartUpTrans.DmctInfo["ma_post"].ToString())[0]["ten_post"];
                if (StartUpTrans.M_ngay_lct.Trim().Equals("0"))
                    row1["ngay_lct"] = this._formcopy.ngay_ct;
                row1["ma_qs"] = this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? "" : this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];
                row1["sl_in"] = 0;
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
                this.iRow_old = FrmSocthda.iRow;
                FrmSocthda.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUp.DataFilter(stt_rec);
                this.IsInEditMode.Value = true;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_kh.IsFocus = true));
            }
        }

        private void V_Xem()
        {
            FormTrans.currActionTask = ActionTask.View;
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            if (StartUp.M_AR_CK == 0 && !StartUp.HiddenFieldIsSetted)
            {
                StartUp.stringBrowse1 = StartUp.EditCkFields(StartUp.stringBrowse1);
                StartUp.stringBrowse2 = StartUp.EditCkFields(StartUp.stringBrowse2);
                StartUp.stringBrowse3 = StartUp.EditCkFields(StartUp.stringBrowse3);
                StartUp.stringBrowse4 = StartUp.EditCkFields(StartUp.stringBrowse4);
            }
            if (StartUp.M_KM_CK == 0 && !StartUp.HiddenFieldIsSetted)
            {
                StartUp.stringBrowse1 = StartUp.EditKmFields(StartUp.stringBrowse1);
                StartUp.stringBrowse2 = StartUp.EditKmFields(StartUp.stringBrowse2);
                StartUp.stringBrowse3 = StartUp.EditKmFields(StartUp.stringBrowse3);
                StartUp.stringBrowse4 = StartUp.EditKmFields(StartUp.stringBrowse4);
            }

            //Gia ban - gia von
            switch (this.stt_mau_temlate)
            {
                case 225:
                    if (!StartUp.HiddenFieldIsSetted)
                    {
                        StartUp.stringBrowse1 = StartUp.EditGiaVonFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditGiaVonFields(StartUp.stringBrowse2);
                        StartUp.stringBrowse3 = StartUp.EditGiaVonFields(StartUp.stringBrowse3);
                        StartUp.stringBrowse4 = StartUp.EditGiaVonFields(StartUp.stringBrowse4);
                    }
                    break;
                case 226:
                    if (!StartUp.HiddenFieldIsSetted)
                    {
                        StartUp.stringBrowse1 = StartUp.EditGiaBanFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditGiaBanFields(StartUp.stringBrowse2);
                        StartUp.stringBrowse3 = StartUp.EditGiaBanFields(StartUp.stringBrowse3);
                        StartUp.stringBrowse4 = StartUp.EditGiaBanFields(StartUp.stringBrowse4);
                    }
                    break;
                case 227:
                    if (!StartUp.HiddenFieldIsSetted)
                    {
                        StartUp.stringBrowse1 = StartUp.EditGiaVonFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditGiaVonFields(StartUp.stringBrowse2);
                        StartUp.stringBrowse3 = StartUp.EditGiaVonFields(StartUp.stringBrowse3);
                        StartUp.stringBrowse4 = StartUp.EditGiaVonFields(StartUp.stringBrowse4);

                        StartUp.stringBrowse1 = StartUp.EditGiaBanFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditGiaBanFields(StartUp.stringBrowse2);
                        StartUp.stringBrowse3 = StartUp.EditGiaBanFields(StartUp.stringBrowse3);
                        StartUp.stringBrowse4 = StartUp.EditGiaBanFields(StartUp.stringBrowse4);
                    }
                    break;
            }

            StartUp.HiddenFieldIsSetted = true;
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.TongCongLabel = "Tổng thanh toán";
            formView.frmBrw.Title = StartUp.M_Tilte;
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "Socthda_9";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string stt_rec = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmSocthda.iRow = index + 1;
                StartUp.DataFilter(stt_rec);
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
            }
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmSearchSocthda frmSearchSocthda = new FrmSearchSocthda(StartupBase.SasObj, StartUpTrans.filterId, StartUpTrans.filterView);
                frmSearchSocthda.txtMa_qs.Filter = "ma_cts like '%HDA%' and status =1 AND (" + this.StrFilterQS + ")";
                SysFunc.LoadIcon((Window)frmSearchSocthda);
                frmSearchSocthda.Closed += new EventHandler(this._FrmTim_Closed);
                frmSearchSocthda.ShowDialog();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
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
            FrmSocthda.hinhthuc_tt = this.txtHt_tt.Text.Trim();
            try
            {
                FrmPrintSocthda frmPrintSocthda = new FrmPrintSocthda(this.IsNd51)
                {
                    DsPrint = StartUpTrans.DsTrans.Copy()
                };
                frmPrintSocthda.DsPrint.Tables[0].TableName = "TablePH";
                frmPrintSocthda.DsPrint.Tables[1].TableName = "TableCT";
                foreach (DataRow dataRow in frmPrintSocthda.DsPrint.Tables[1].AsEnumerable())
                {
                    dataRow["dvt_qd"] = dataRow["dvt_qd"];
                    dataRow["sl_qd"] = dataRow["so_luong"];
                    dataRow["gia_qd_nt"] = dataRow["gia_nt"];
                    dataRow["gia_qd"] = dataRow["gia"];
                    dataRow["gia_qd_nt2"] = dataRow["gia_nt2"];
                    dataRow["gia_qd2"] = dataRow["gia2"];
                }
                string str = frmPrintSocthda.DsPrint.Tables["TablePH"].Rows[FrmSocthda.iRow]["stt_rec"].ToString();
                SqlCommand sqlcmd = new SqlCommand(string.Format("EXEC dbo.[SOCTHDA-Getslqd] '{0}'", str));
                DataSet dataSet = FormTrans.SasO.ExcuteReader(sqlcmd);
                if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    Decimal num1 = new Decimal(0);
                    Decimal num2 = new Decimal(0);
                    IEnumerator enumerator = dataSet.Tables[0].Rows.GetEnumerator();
                    try
                    {
                        while (enumerator.MoveNext())
                        {
                            DataRow r = (DataRow)enumerator.Current;
                            foreach (DataRow dataRow in frmPrintSocthda.DsPrint.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(c => c.Field<string>("stt_rec0") == r["stt_rec0"].ToString())))
                            {
                                dataRow["dvt_qd"] = r["dvt_qd"];
                                dataRow["sl_qd"] = (num1 = (Decimal)r["sl_qd"]);
                                dataRow["gia_qd_nt"] = r["gia_qd_nt"];
                                dataRow["gia_qd"] = r["gia_qd"];
                                dataRow["gia_qd_nt2"] = r["gia_qd_nt2"];
                                dataRow["gia_qd2"] = r["gia_qd2"];
                            }
                        }
                    }
                    finally
                    {
                        if (enumerator is IDisposable disposable)
                            disposable.Dispose();
                    }
                }
                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_sl_qd"] = this.SumFunction(frmPrintSocthda.DsPrint.Tables[1], "sl_qd", 0);
                frmPrintSocthda.DsPrint.Tables["TablePH"].Columns.Add(new DataColumn("so_lien", typeof(int))
                {
                    DefaultValue = 1
                });
                frmPrintSocthda.DsPrint.Tables["TablePH"].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
                {
                    DefaultValue = 0
                });
                frmPrintSocthda.DsPrint.Tables["TablePH"].Columns.Add(new DataColumn("ban_sao", typeof(string))
                {
                    DefaultValue = ""
                });
                frmPrintSocthda.DsPrint.Tables["TableCT"].Columns.Add(new DataColumn("tag", typeof(int))
                {
                    DefaultValue = 0
                });
                frmPrintSocthda.DsPrint.Tables["TableCT"].Columns.Add(new DataColumn("stt", typeof(string))
                {
                    DefaultValue = ""
                });
                frmPrintSocthda.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                frmPrintSocthda.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                frmPrintSocthda.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
                frmPrintSocthda.DsPrint.Tables.Add(StartUp.GetDmnt().Copy());
                frmPrintSocthda.DsPrint.Tables.Add(this.CreateTableInfo().Copy());
                frmPrintSocthda.DsPrint.Tables.Add(FrmSocthda.CreateTableMST(StartUp.M_MA_THUE, "TableMST_NB").Copy());
                frmPrintSocthda.DsPrint.Tables.Add(FrmSocthda.CreateTableMST(frmPrintSocthda.DsPrint.Tables["TablePH"].Rows[FrmSocthda.iRow]["ma_so_thue"].ToString().TrimEnd(), "TableMST_NM").Copy());
                frmPrintSocthda.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            this.UpdateTonKho();
        }

        private DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "TableInfo";
            dataTable.Columns.Add(new DataColumn("M_PHONE", typeof(string))
            {
                DefaultValue = StartUp.M_PHONE
            });
            dataTable.Columns.Add(new DataColumn("M_MST", typeof(string))
            {
                DefaultValue = StartUp.M_MA_THUE
            });
            dataTable.Columns.Add(new DataColumn("M_Ten_CTY", typeof(string))
            {
                DefaultValue = StartupBase.SasObj.GetSysvar("M_Ten_CTY").ToString().ToUpper()
            });
            dataTable.Columns.Add(new DataColumn("M_DIA_CHI", typeof(string))
            {
                DefaultValue = StartupBase.SasObj.GetSysvar("M_DIA_CHI").ToString()
            });
            dataTable.Columns.Add(new DataColumn("M_TK_NH", typeof(string))
            {
                DefaultValue = StartupBase.SasObj.GetOption("M_TK_NH").ToString()
            });
            dataTable.Columns.Add(new DataColumn("M_CUC_THUE", typeof(string))
            {
                DefaultValue = StartupBase.SasObj.GetOption("M_CUC_THUE").ToString().ToUpper()
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        public static DataTable CreateTableMST(string ma_so_thue, string tablename)
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = tablename;
            int length = ma_so_thue.Length;
            int startIndex;
            int num;
            for (startIndex = 0; startIndex < length; ++startIndex)
            {
                num = startIndex + 1;
                dataTable.Columns.Add(new DataColumn("m" + num.ToString(), typeof(string))
                {
                    DefaultValue = ma_so_thue.Substring(startIndex, 1)
                });
            }
            for (; startIndex < 14; ++startIndex)
            {
                num = startIndex + 1;
                dataTable.Columns.Add(new DataColumn("m" + num.ToString(), typeof(string))
                {
                    DefaultValue = ""
                });
            }
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            if (FormTrans.currActionTask != ActionTask.Add)
            {
                this.LoadDataDu13();
                this.UpdateTonKho();
                this.UpdateKM();
            }
            this.IsVisibilityFieldsXamDataGridByMa_NT(ma_nt);
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            this.IsVisibilityFieldsXamDataGridByPx_gia_dd();
            if (StartUp.M_KM_CK != 0)
                ;
            if (StartUp.M_AR_CK != 0)
                return;
            this.ChkTinh_ck.IsEnabled = false;
        }

        private void IsVisibilityFieldsXamDataGridByMa_NT(string ma_nt)
        {
            if (ma_nt == StartUpTrans.M_ma_nt0)
            {
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["ck"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["ck"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_tien2"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_tien2"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_ck"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_ck"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_tien_sau_ck"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_tien_sau_ck"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_thue"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_thue"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_tien_km"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_tien_km"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["tien_tc"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["tien_tc"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_tt_km"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_tt_km"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_thue_km"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_thue_km"].Settings.CellMaxWidth = 0.0;
                this.GrdTT.FieldLayouts[0].Fields["t_tt"].Visibility = Visibility.Hidden;
                this.GrdTT.FieldLayouts[0].Fields["t_tt"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Visible;
                FieldSettings settings1 = this.GrdCt.FieldLayouts[0].Fields["gia"].Settings;
                FieldLength? width = this.GrdCt.FieldLayouts[0].Fields["gia"].Width;
                double num1 = width.Value.Value;
                settings1.CellMaxWidth = num1;
                FieldSettings settings2 = this.GrdCt.FieldLayouts[0].Fields["tien"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["tien"].Width;
                double num2 = width.Value.Value;
                settings2.CellMaxWidth = num2;
                FieldSettings settings3 = this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["gia2"].Width;
                double num3 = width.Value.Value;
                settings3.CellMaxWidth = num3;
                FieldSettings settings4 = this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["tien2"].Width;
                double num4 = width.Value.Value;
                settings4.CellMaxWidth = num4;
                this.GrdTT.FieldLayouts[0].Fields["t_tien2"].Visibility = Visibility.Visible;
                FieldSettings settings5 = this.GrdTT.FieldLayouts[0].Fields["t_tien2"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_tien2"].Width;
                double num5 = width.Value.Value;
                settings5.CellMaxWidth = num5;
                this.GrdTT.FieldLayouts[0].Fields["t_thue"].Visibility = Visibility.Visible;
                FieldSettings settings6 = this.GrdTT.FieldLayouts[0].Fields["t_thue"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_thue"].Width;
                double num6 = width.Value.Value;
                settings6.CellMaxWidth = num6;
                this.GrdTT.FieldLayouts[0].Fields["t_tien_km"].Visibility = Visibility.Visible;
                FieldSettings settings7 = this.GrdTT.FieldLayouts[0].Fields["t_tien_km"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_tien_km"].Width;
                double num7 = width.Value.Value;
                settings7.CellMaxWidth = num7;
                this.GrdTT.FieldLayouts[0].Fields["tien_tc"].Visibility = Visibility.Visible;
                FieldSettings settings8 = this.GrdTT.FieldLayouts[0].Fields["tien_tc"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["tien_tc"].Width;
                double num8 = width.Value.Value;
                settings8.CellMaxWidth = num8;
                this.GrdTT.FieldLayouts[0].Fields["t_tt_km"].Visibility = Visibility.Visible;
                FieldSettings settings9 = this.GrdTT.FieldLayouts[0].Fields["t_tt_km"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_tt_km"].Width;
                FieldLength fieldLength = width.Value;
                double num9 = fieldLength.Value;
                settings9.CellMaxWidth = num9;
                this.GrdTT.FieldLayouts[0].Fields["t_thue_km"].Visibility = Visibility.Visible;
                FieldSettings settings10 = this.GrdTT.FieldLayouts[0].Fields["t_thue_km"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_thue_km"].Width;
                fieldLength = width.Value;
                double num10 = fieldLength.Value;
                settings10.CellMaxWidth = num10;
                this.GrdTT.FieldLayouts[0].Fields["t_tt"].Visibility = Visibility.Visible;
                FieldSettings settings11 = this.GrdTT.FieldLayouts[0].Fields["t_tt"].Settings;
                width = this.GrdTT.FieldLayouts[0].Fields["t_tt"].Width;
                fieldLength = width.Value;
                double num11 = fieldLength.Value;
                settings11.CellMaxWidth = num11;
                if (StartUp.M_AR_CK != 0)
                {
                    this.GrdCt.FieldLayouts[0].Fields["ck"].Visibility = Visibility.Visible;
                    FieldSettings settings12 = this.GrdCt.FieldLayouts[0].Fields["ck"].Settings;
                    width = this.GrdCt.FieldLayouts[0].Fields["ck"].Width;
                    fieldLength = width.Value;
                    double num12 = fieldLength.Value;
                    settings12.CellMaxWidth = num12;
                    this.GrdTT.FieldLayouts[0].Fields["t_ck"].Visibility = Visibility.Visible;
                    FieldSettings settings13 = this.GrdTT.FieldLayouts[0].Fields["t_ck"].Settings;
                    width = this.GrdTT.FieldLayouts[0].Fields["t_ck"].Width;
                    fieldLength = width.Value;
                    double num13 = fieldLength.Value;
                    settings13.CellMaxWidth = num13;
                    this.GrdTT.FieldLayouts[0].Fields["t_tien_sau_ck"].Visibility = Visibility.Visible;
                    FieldSettings settings14 = this.GrdTT.FieldLayouts[0].Fields["t_tien_sau_ck"].Settings;
                    width = this.GrdTT.FieldLayouts[0].Fields["t_tien_sau_ck"].Width;
                    fieldLength = width.Value;
                    double num14 = fieldLength.Value;
                    settings14.CellMaxWidth = num14;
                }
                switch(this.stt_mau_temlate)
                {
                    case 225: //Gia von
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Settings.CellMaxWidth = 0.0;

                        break;
                    case 226: //Gia ban
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Settings.CellMaxWidth = 0.0;

                        break;
                    case 227: //Gia ban -Gia von
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt"].Settings.CellMaxWidth = 0.0;

                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Visibility = Visibility.Hidden;
                        this.GrdCt.FieldLayouts[0].Fields["gia_nt2"].Settings.CellMaxWidth = 0.0;
                        this.GrdCt.FieldLayouts[0].Fields["tien_nt2"].Settings.CellMaxWidth = 0.0;

                        break;
                }
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
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_gd"] = !this.M_LAN.ToUpper().Equals("V") ? (object)this.txtma_gd.RowResult["ten_gd2"].ToString().Trim() : (object)this.txtma_gd.RowResult["ten_gd"].ToString().Trim();
        }

        private void LoadDataDu13()
        {
            this.txtSoDuKH.Value = ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString());
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            DataRow rowResult = this.txtMa_kh.RowResult;
            if (rowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()))
                return;
            if (this.txtMa_kh.IsDataChanged)
            {
                dataRowView["ten_kh_thue"] = rowResult["ten_kh"].ToString().Trim();
                dataRowView["ten_kh"] = rowResult["ten_kh"].ToString().Trim();
                dataRowView["ten_kh2"] = rowResult["ten_kh2"].ToString().Trim();
                if (e != null)
                {
                    dataRowView["ma_so_thue_dmkh"] = rowResult["ma_so_thue"].ToString().Trim();
                    dataRowView["ma_so_thue"] = rowResult["ma_so_thue"].ToString().Trim();
                }
                if (string.IsNullOrEmpty(dataRowView["ong_ba"].ToString().Trim()))
                    dataRowView["ong_ba"] = rowResult["doi_tac"].ToString().Trim();
                if (dataRowView["ma_nx"].ToString().Trim() == "")
                {
                    this.txtMa_nx.Text = rowResult["tk"].ToString().Trim();
                    this.txtMa_nx.SearchInit();
                    if (this.txtMa_nx.RowResult != null)
                    {
                        dataRowView["ten_nx"] = this.txtMa_nx.RowResult["ten_nx"].ToString();
                        dataRowView["ten_nx2"] = this.txtMa_nx.RowResult["ten_nx2"].ToString();
                    }
                    else
                    {
                        dataRowView["ten_nx"] = "";
                        dataRowView["ten_nx2"] = "";
                    }
                }
                if (string.IsNullOrEmpty(dataRowView["ma_thck"].ToString().Trim()))
                    dataRowView["ma_thck"] = rowResult["ma_thck"];
                if (this.ParseInt(dataRowView["han_tt"].ToString(), 0) == 0)
                    dataRowView["han_tt"] = this.ParseInt(rowResult["han_tt"], 0);
            }
            if (string.IsNullOrEmpty(rowResult["dia_chi"].ToString().Trim()))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                dataRowView["dia_chi"] = rowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
            if (!string.IsNullOrEmpty(this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_so_thue"] = this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim();
            this.LoadDataDu13();
            dataRowView["tk_nh"] = rowResult["tk_nh"].ToString();
            dataRowView["nh_kh3"] = rowResult["nh_kh3"].ToString();
            if (StartUp.M_SD_HDDT.Equals("1") && this.txtMa_kh.IsDataChanged)
                dataRowView["sd_hddt_yn"] = SysFunc.suDungHDDT(dataRowView["ma_kh"].ToString());
        }

        private void txtDia_chi_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtDiaChiFocusable)
                return;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtMa_nx_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_nx.RowResult == null)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = "";
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = "";
            if (!string.IsNullOrEmpty(this.txtMa_nx.Text.Trim()))
            {
                if (StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["sua_tkthue"].ToString() == "0")
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_no"] = this.txtMa_nx.RowResult["ma_nx"].ToString().Trim();
                if (this.M_LAN.ToUpper().Equals("V"))
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = this.txtMa_nx.RowResult["ten_nx"].ToString();
                else
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = this.txtMa_nx.RowResult["ten_nx2"].ToString();
            }
            this.LoadDataDu13();
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
            if (this.txtNgay_lct.IsFocusWithin || !this.IsInEditMode.Value || !this.txtNgay_lct.IsValueValid || (this.txtNgay_lct.Value == null || !(this.txtNgay_ct.Value.ToString() != this.txtNgay_lct.Value.ToString())))
                return;
            int num = (int)ExMessageBox.Show(795, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            DataRowView ph = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            DataRow qs = this.txtMa_qs.RowResult;
            if (!this.IsInEditMode.Value)
                return;
            if (!string.IsNullOrEmpty(qs["so_seri"].ToString()))
                ph["so_seri"] = qs["so_seri"].ToString().Trim();
            if (!string.IsNullOrEmpty(qs["so_ct1"].ToString()))
                ph["so_ct1"] = qs["so_ct1"].ToString();
            if (!string.IsNullOrEmpty(qs["so_ct2"].ToString()))
                ph["so_ct2"] = qs["so_ct2"].ToString();
            if (!string.IsNullOrEmpty(qs["transform"].ToString()))
                ph["transform"] = qs["transform"].ToString();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (!this.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(ph["ma_qs"].ToString()))
                   return;
               if (string.IsNullOrEmpty(ph["so_ct"].ToString().Trim()) || this.IsNd51 && this.txtMa_qs.IsDataChanged)
               {
                   if (string.IsNullOrEmpty(ph["so_cttmp"].ToString().Trim()) || !ph["ma_qs"].ToString().Trim().Equals(ph["ma_qstmp"].ToString().Trim()) || this.IsNd51)
                   {
                       this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                       ph["so_cttmp"] = this.txtSo_ct.Text;
                       ph["ma_qstmp"] = this.txtMa_qs.Text;
                   }
                   else
                       this.txtSo_ct.Text = ph["so_cttmp"].ToString().Trim();
               }
               if (this.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtSo_ct.Text, ph["stt_rec"].ToString()))
               {
                   this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                   ph["so_cttmp"] = this.txtSo_ct.Text;
                   ph["ma_qstmp"] = this.txtMa_qs.Text;
               }
               if (!string.IsNullOrEmpty(qs["so_lien_hd"].ToString()))
                   ph["so_lien_hd"] = qs["so_lien_hd"];
               if (!string.IsNullOrEmpty(qs["ten_lien1"].ToString()))
                   ph["ten_lien1"] = qs["ten_lien1"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien2"].ToString()))
                   ph["ten_lien2"] = qs["ten_lien2"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien3"].ToString()))
                   ph["ten_lien3"] = qs["ten_lien3"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien4"].ToString()))
                   ph["ten_lien4"] = qs["ten_lien4"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien5"].ToString()))
                   ph["ten_lien5"] = qs["ten_lien5"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien6"].ToString()))
                   ph["ten_lien6"] = qs["ten_lien6"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien7"].ToString()))
                   ph["ten_lien7"] = qs["ten_lien7"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien8"].ToString()))
                   ph["ten_lien8"] = qs["ten_lien8"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_lien9"].ToString()))
                   ph["ten_lien9"] = qs["ten_lien9"].ToString();
               if (!string.IsNullOrEmpty(qs["ten_dn_in"].ToString()))
                   ph["ten_dn_in"] = qs["ten_dn_in"].ToString();
               if (!string.IsNullOrEmpty(qs["mst_dn_in"].ToString()))
                   ph["mst_dn_in"] = qs["mst_dn_in"].ToString();
               if (!string.IsNullOrEmpty(qs["mau_hd"].ToString()))
                   ph["mau_hd"] = qs["mau_hd"].ToString();
               if (!string.IsNullOrEmpty(qs["ma_file"].ToString()))
                   ph["ma_file"] = qs["ma_file"].ToString();
           }));
        }

        private void txtMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.txtMa_nt.IsDataChanged)
                return;
            this.IsVisibilityFieldsXamDataGridByMa_NT(this.txtMa_nt.Text.Trim());
            if (this.txtMa_nt.RowResult != null)
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.txtMa_nt.RowResult["loai_tg"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_giaf"] = !this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? StartUp.GetRates(this.txtMa_nt.Text.Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date) : 1;
            }
            this.Ty_gia_ValueChanged(true);
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtTy_gia.Value == DBNull.Value)
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_giaf"] = 0;
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
            Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            if (!(num1 != new Decimal(0)) || num2 != 0 && !IsMa_ntChanged)
                return;
            this.UpdateTotal("ck", "ck_nt", false);
            this.UpdateTotal("gia2", "gia_nt2", true);
            this.UpdateTotal("tien2", "tien_nt2", false);
            this.UpdateTotal("gia", "gia_nt", true);
            this.UpdateTotal("tien", "tien_nt", false);
            this.UpdateTotalKM("tien_km", "tien_km_nt", false);

            Decimal num4 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 1);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km"] = num4;
            if (StartUp.M_THUE_KM_CK == 1)
                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km"] = SysFunc.Round(num4 * num3 / new Decimal(100), StartUpTrans.M_ROUND);

            Decimal num5 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"], new Decimal(0)); 
            if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = (object) SysFunc.Round((num5 / num1),StartUpTrans.M_ROUND_NT);
            }
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0) + this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_km", 0)  + Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"]);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0) + this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_km_nt", 0) + Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"]);


        }

        private void UpdateTotal(string columnname, string columnname_nt, bool isPrice)
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname_nt], new Decimal(0));
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname] = num2;
                }
                else
                {
                    Decimal num3 = SysFunc.Round(num2 * num1, !isPrice ? StartUpTrans.M_ROUND : StartUpTrans.M_ROUND_GIA);
                    if (num3 != new Decimal(0))
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index][columnname] = num3;
                }
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
        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtTk_thue_co_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.IsInEditMode.Value && (this.txtTk_thue_co.RowResult != null && !string.IsNullOrEmpty(this.txtTk_thue_co.Text.Trim())))
            {
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_cn"] = this.txtTk_thue_co.RowResult["tk_cn"];
                this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtTk_thue_co.Text = this.txtTk_thue_co.Text.Trim()), DispatcherPriority.Background);
            }
        }

        private void ChkSua_tien_Click(object sender, RoutedEventArgs e)
        {
            bool? isChecked = this.ChkSua_tien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 && sender.GetType().Name.Equals("CheckBox"))
            {
                this.UpdateTotalChkSua_tien();
                this.Ty_gia_ValueChanged(false);
            }
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
        }

        private void ChkCo_ck_Click(object sender, RoutedEventArgs e)
        {
            this.IsCheckedCo_km.Value = this.ChkCo_km.IsChecked.Value;
        }

        private void UpdateTotalChkSua_tien()
        {
            int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
            if (count <= 0)
                return;
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            for (int index = 0; index < count; ++index)
            {
                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["so_luong"], new Decimal(0));
                Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt2"], new Decimal(0));
                Decimal num4 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["gia_nt"], new Decimal(0));
                if (num2 != new Decimal(0))
                {
                    if (num3 != new Decimal(0))
                    {
                        Decimal num5 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND_NT);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt2"] = num5;
                        Decimal num6 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tl_ck"], new Decimal(0));
                        if (num6 != new Decimal(0))
                        {
                            Decimal num7 = SysFunc.Round(num5 * num6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck_nt"] = num7;
                            Decimal num8 = SysFunc.Round(num7 * num1, StartUpTrans.M_ROUND);
                            if (num8 != new Decimal(0))
                                StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck"] = num8;
                        }
                    }
                    if (num4 != new Decimal(0))
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt"] = SysFunc.Round(num2 * num4, StartUpTrans.M_ROUND_NT);
                }
            }
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
            this.UpdateKM();
        }

        private void GrdCt_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (!this.IsInEditMode.Value || (this.GrdCt.ActiveCell == null || StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) != null))
                    return;
                Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
                this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
                this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_ck"], 0);
                string nh_kh3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["nh_kh3"].ToString();
                switch (e.Cell.Field.Name)
                {
                    case "ma_vt":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        DataRowView dataItem1 = e.Cell.Record.DataItem as DataRowView;
                        CellCollection cells = e.Cell.Record.Cells;
                        DataRow rowResult = autoCompleteControl1.RowResult;
                        if (e.Cell.IsDataChanged && rowResult != null)
                        {
                            cells["ten_vt"].Value = rowResult["ten_vt"];
                            cells["ten_vt2"].Value = rowResult["ten_vt2"];
                            cells["dvt"].Value = rowResult["dvt"];
                            cells["tk_dt_dmvt"].Value = rowResult["tk_dt"].ToString();
                            cells["tk_ck_dmvt"].Value = rowResult["tk_ck"].ToString();
                            cells["tk_gv_dmvt"].Value = rowResult["tk_gv"].ToString();
                            cells["tk_km_dmvt"].Value = rowResult["tk_km"].ToString();
                            (e.Cell.Record.DataItem as DataRowView)["vt_ton_kho"] = autoCompleteControl1.RowResult["vt_ton_kho"];
                            if (rowResult["tk_dt"].ToString().Trim() != "")
                                cells["tk_dt"].Value = rowResult["tk_dt"].ToString();
                            if (rowResult["tk_ck"].ToString().Trim() != "")
                                cells["tk_ck"].Value = rowResult["tk_ck"].ToString();
                            if (rowResult["tk_km"].ToString().Trim() != "")
                                cells["tk_km_i"].Value = rowResult["tk_km"];
                            if (cells["km_ck"].Value.ToString() == "0")
                                cells["tk_km_i"].Value = "";
                            if (rowResult["tk_gv"].ToString().Trim() != "")
                            {
                                e.Cell.Record.Cells["tk_gv"].Value = autoCompleteControl1.RowResult["tk_gv"].ToString();
                                if (cells["tk_gv_dmvt"].Value.ToString().Trim() == "")
                                    cells["tk_gv"].Value = cells["tk_km_i"].Value;
                            }
                            cells["gia_ton"].Value = rowResult["gia_ton"];
                            cells["vt_ton_kho"].Value = rowResult["vt_ton_kho"];
                            if (this.ParseInt(rowResult["vt_ton_kho"], 0) == 0)
                            {
                                cells["so_luong"].Value = 0;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_so_luong"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                                cells["gia_nt2"].Value = 0;
                                cells["gia2"].Value = 0;
                                cells["gia_nt"].Value = 0;
                                cells["gia"].Value = 0;
                                cells["tien"].Value = 0;
                                cells["tien_nt"].Value = 0;
                            }
                            dataItem1["sua_tk_vt"] = rowResult["sua_tk_vt"];
                            dataItem1["sl_min"] = rowResult["sl_min"];
                            dataItem1["tk_vt_dmvt"] = rowResult["tk_vt"].ToString();
                            if (rowResult["tk_vt"].ToString() != "")
                                cells["tk_vt"].Value = rowResult["tk_vt"].ToString();
                            if (this.txtNgay_ct.dValue != new DateTime())
                            {
                                DataRow dataRow = StartUpTrans.Getdmgia2(e.Editor.Value.ToString(), string.Format("{0:yyyyMMdd}", this.txtNgay_ct.dValue), nh_kh3);
                                if (dataRow != null)
                                {
                                    cells["gia_nt2"].Value = !this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? dataRow["gia_nt2"] : dataRow["gia2"];
                                    cells["gia2"].Value = dataRow["gia2"];
                                }
                            }
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(cells["ma_kho_i"]).Editor as ControlHostEditor);
                            if (autoCompleteControl2 != null)
                            {
                                autoCompleteControl2.SearchInit();
                                if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                                {
                                    cells["tk_vt"].Value = autoCompleteControl2.RowResult["tk_dl"].ToString();
                                    dataItem1["tk_vt_dmvt"] = autoCompleteControl2.RowResult["tk_dl"].ToString();
                                }
                            }
                            ControlFunction.RefreshSingleBinding((DependencyObject)CellValuePresenter.FromCell(cells["tk_vt"]), AutoCompleteTextBox.IsReadOnlyProperty);
                            if (this.ParseInt(rowResult["vt_ton_kho"], 0) == 1)
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
                                cells["ton13"].Value = DBNull.Value;

                            if (string.IsNullOrEmpty(cells["dvt1"].Value.ToString()))
                            {
                                cells["dvt1"].Value = cells["dvt"].Value;
                                AutoCompleteTextBox autoCompleteControlmavtdvt = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["dvt1"]).Editor as ControlHostEditor);
                                if (autoCompleteControlmavtdvt != null)
                                {
                                    autoCompleteControlmavtdvt.SearchInit();
                                    if (autoCompleteControlmavtdvt.RowResult != null && (autoCompleteControlmavtdvt.RowResult["hs_qd"] != DBNull.Value))
                                    {
                                        dataItem1["he_so1"] = (decimal)autoCompleteControlmavtdvt.RowResult["hs_qd"];
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
                        AutoCompleteTextBox autoCompleteControlDvt = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        DataRowView dataItemDvt = e.Cell.Record.DataItem as DataRowView;
                        if (autoCompleteControlDvt.IsDataChanged)
                        {
                            if (autoCompleteControlDvt != null)
                            {
                                autoCompleteControlDvt.SearchInit();
                                if (autoCompleteControlDvt.RowResult != null && (autoCompleteControlDvt.RowResult["hs_qd"] != DBNull.Value))
                                {
                                    dataItemDvt["he_so1"] = (decimal)autoCompleteControlDvt.RowResult["hs_qd"];
                                }
                                else
                                {
                                    dataItemDvt["he_so1"] = (decimal)0;
                                }
                            }
                            break;
                        }
                        break;
                    case "km_ck":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["km_ck"].Value = 0;
                            if (e.Cell.Record.Cells["km_ck"].Value.ToString() == "1")
                            {
                                e.Cell.Record.Cells["tk_km_i"].Value = e.Cell.Record.Cells["tk_km_dmvt"].Value;
                                e.Cell.Record.Cells["tk_dt"].Value = "";
                                if (e.Cell.Record.Cells["tk_gv"].Value.ToString().Trim() == "")
                                    e.Cell.Record.Cells["tk_gv"].Value = e.Cell.Record.Cells["tk_km_i"].Value;
                            }
                            else
                            {
                                e.Cell.Record.Cells["tk_dt"].Value = e.Cell.Record.Cells["tk_dt_dmvt"].Value;
                                e.Cell.Record.Cells["tk_km_i"].Value = "";
                            }
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_so_luong"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            this.UpdateKM();
                        }
                        object count = this.GrdCt.Records.Count;
                        break;
                    case "tk_km_i":
                        if (e.Cell.IsDataChanged && e.Cell.Record.Cells["tk_gv"].Value.ToString().Trim() == "")
                        {
                            e.Cell.Record.Cells["tk_gv"].Value = e.Cell.Record.Cells["tk_km_i"].Value;
                            break;
                        }
                        break;
                    case "ma_kho_i":
                        if (e.Editor.Value == null)
                            break;
                        AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                        if (autoCompleteControl3.IsDataChanged)
                        {
                            DataRowView dataItem2 = e.Cell.Record.DataItem as DataRowView;
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_vt"]).Editor as ControlHostEditor);
                            if (autoCompleteControl2.RowResult == null)
                                autoCompleteControl2.SearchInit();
                            if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["sua_tk_vt"] != DBNull.Value && Convert.ToDecimal(autoCompleteControl2.RowResult["sua_tk_vt"]) == new Decimal(0)))
                            {
                                e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl2.RowResult["tk_vt"].ToString();
                                dataItem2["tk_vt_dmvt"] = autoCompleteControl2.RowResult["tk_vt"].ToString();
                            }
                            if (autoCompleteControl3 != null)
                            {
                                autoCompleteControl3.SearchInit();
                                if (autoCompleteControl3.RowResult != null && (autoCompleteControl3.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl3.RowResult["tk_dl"].ToString().Trim())))
                                {
                                    if (autoCompleteControl3.RowResult["tk_dl"].ToString() != "")
                                        e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl3.RowResult["tk_dl"].ToString();
                                    dataItem2["tk_vt_dmvt"] = autoCompleteControl3.RowResult["tk_dl"].ToString();
                                }
                            }
                            if (autoCompleteControl2.RowResult != null)
                            {
                                if (this.ParseInt(autoCompleteControl2.RowResult["vt_ton_kho"], 0) == 1)
                                {
                                    if (!string.IsNullOrEmpty(e.Cell.Record.Cells["ma_vt"].Value.ToString()) && !string.IsNullOrEmpty(e.Cell.Record.Cells["ma_kho_i"].Value.ToString()))
                                        e.Cell.Record.Cells["ton13"].Value = InFuncLib.GetTon13(StartupBase.SasObj, e.Cell.Record.Cells["ma_kho_i"].Value.ToString(), e.Cell.Record.Cells["ma_vt"].Value.ToString(), (e.Cell.Record.DataItem as DataRowView)["ma_vv_i"].ToString());
                                }
                                else
                                    e.Cell.Record.Cells["ton13"].Value = DBNull.Value;
                            }
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
                        }
                        break;
                    case "so_luong":
                        if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                            e.Cell.Record.Cells["so_luong"].Value = 0;
                        Decimal num2 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                        int result = 0;
                        if (e.Cell.Record.Cells["gia_ton"].Value != null)
                            int.TryParse(e.Cell.Record.Cells["gia_ton"].Value.ToString(), out result);
                        if (result == 3 && num2 == new Decimal(0))
                        {
                            int num3 = (int)ExMessageBox.Show(800, StartupBase.SasObj, "Vật tư tính tồn kho theo phương pháp NTXT không được nhập số lượng = 0!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdCt.ActiveCell = e.Cell.Record.Cells["so_luong"]));
                            break;
                        }
                        if (e.Cell.IsDataChanged)
                        {
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_so_luong"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_sl_km"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 1);
                            if (num2 == new Decimal(0))
                            {
                                e.Cell.Record.Cells["gia_nt"].Value = 0;
                                e.Cell.Record.Cells["gia"].Value = 0;
                                e.Cell.Record.Cells["gia_nt2"].Value = 0;
                                e.Cell.Record.Cells["gia2"].Value = 0;
                            }
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                            Decimal num5 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND_GIA);
                            Decimal num6 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND_GIA);
                            if (num2 != new Decimal(0))
                            {
                                if (num5 != new Decimal(0))
                                    e.Cell.Record.Cells["gia2"].Value = num5;
                                if (num6 != new Decimal(0))
                                    e.Cell.Record.Cells["gia"].Value = num6;
                            }
                            Decimal num7 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND_NT);
                            Decimal num8 = SysFunc.Round(num2 * num4, StartUpTrans.M_ROUND_NT);
                            if (num7 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien_nt2"].Value = num7;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
                            }
                            if (num8 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien_nt"].Value = num8;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            }
                            Decimal num9 = SysFunc.Round(num7 * num1, StartUpTrans.M_ROUND);
                            Decimal num10 = SysFunc.Round(num8 * num1, StartUpTrans.M_ROUND);
                            if (num9 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien2"].Value = num9;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            }
                            if (num10 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien"].Value = num10;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            }
                            Decimal num11 = this.ParseDecimal(e.Cell.Record.Cells["tl_ck"].Value, new Decimal(0));
                            if (num11 != new Decimal(0))
                            {
                                Decimal num12 = SysFunc.Round(num7 * num11 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["ck_nt"].Value = num12;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
                                Decimal num13 = SysFunc.Round(num12 * num1, StartUpTrans.M_ROUND);
                                if (num13 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["ck"].Value = (object)num13;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
                                }
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["gia"].Value = e.Cell.Record.Cells["gia_nt"].Value;
                                e.Cell.Record.Cells["gia2"].Value = e.Cell.Record.Cells["gia_nt2"].Value;
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                e.Cell.Record.Cells["tien2"].Value = e.Cell.Record.Cells["tien_nt2"].Value;
                                e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"];
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"];
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"];
                            }                           
                            AddKhuyenmai();
                            break;
                        }
                        break;
                    case "gia_nt2":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["gia_nt2"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                            Decimal num5 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND_GIA);
                            if (num3 != new Decimal(0) && num5 != new Decimal(0))
                                e.Cell.Record.Cells["gia2"].Value = num5;
                            Decimal num6 = new Decimal(0);
                            Decimal num7 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                            bool? isChecked = this.ChkSua_tien.IsChecked;
                            if ((!isChecked.GetValueOrDefault() ? 1 : (!isChecked.HasValue ? 1 : 0)) != 0)
                            {
                                num7 = SysFunc.Round(num3 * num4, StartUpTrans.M_ROUND_NT);
                                if (num7 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["tien_nt2"].Value = num7;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
                                }
                            }
                            Decimal num8 = SysFunc.Round(num7 * num1, StartUpTrans.M_ROUND);
                            if (num8 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien2"].Value = num8;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["gia2"].Value = e.Cell.Record.Cells["gia_nt2"].Value;
                                e.Cell.Record.Cells["tien2"].Value = e.Cell.Record.Cells["tien_nt2"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"];
                            }
                            //if (this.ParseInt(e.Cell.Record.Cells["km_ck"].Value, 0) == 1)
                            //    this.UpdateKM();
                            AddKhuyenmai();
                            break;
                        }
                        break;
                    case "tien_nt2":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien_nt2"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num5 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                            if (num5 == new Decimal(0))
                            {
                                if (num4 != new Decimal(0))
                                    num5 = SysFunc.Round(num3 / num4, StartUpTrans.M_ROUND_GIA);
                                e.Cell.Record.Cells["gia_nt2"].Value = num5;
                            }
                            Decimal num6 = this.ParseDecimal(e.Cell.Record.Cells["tl_ck"].Value, new Decimal(0));
                            if (num6 != new Decimal(0))
                            {
                                Decimal num7 = SysFunc.Round(num3 * num6 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["ck_nt"].Value = num7;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
                                Decimal num8 = SysFunc.Round(num7 * num1, StartUpTrans.M_ROUND);
                                if (num8 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["ck"].Value = num8;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
                                }
                            }
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
                            Decimal num9 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND);
                            if (num9 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien2"].Value = num9;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien2"].Value = e.Cell.Record.Cells["tien_nt2"].Value;
                                e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"];
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"];
                            }
                            if (this.ParseInt(e.Cell.Record.Cells["km_ck"].Value, 0) == 1)
                                this.UpdateKM();
                            break;
                        }
                        break;
                    case "gia_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["gia_nt"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                            Decimal num4 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt"].Value, new Decimal(0));
                            Decimal num5 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND_GIA);
                            if (num3 != new Decimal(0) && num5 != new Decimal(0))
                                e.Cell.Record.Cells["gia"].Value = num5;
                            Decimal num6 = SysFunc.Round(num3 * num4, StartUpTrans.M_ROUND_NT);
                            if (num6 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien_nt"].Value = num6;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            }
                            Decimal num7 = SysFunc.Round(num6 * num1, StartUpTrans.M_ROUND);
                            if (num7 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien"].Value = num7;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["gia"].Value = e.Cell.Record.Cells["gia_nt"].Value;
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"];
                            }
                            AddKhuyenmai();
                            break;
                        }
                        break;
                    case "tien_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien_nt"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt"].Value, new Decimal(0));
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            Decimal num4 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND);
                            if (num4 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["tien"].Value = num4;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["tien"].Value = e.Cell.Record.Cells["tien_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"];
                            }
                            break;
                        }
                        break;
                    case "tl_ck":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tl_ck"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["tl_ck"].Value, new Decimal(0));
                            if (num3 != new Decimal(0))
                            {
                                Decimal num4 = SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0)) * num3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                                e.Cell.Record.Cells["ck_nt"].Value = num4;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
                                Decimal num5 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND);
                                if (num5 != new Decimal(0))
                                {
                                    e.Cell.Record.Cells["ck"].Value = num5;
                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
                                }
                            }
                            else
                            {
                                e.Cell.Record.Cells["ck_nt"].Value = 0;
                                e.Cell.Record.Cells["ck"].Value = 0;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = 0;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = 0;
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"];
                            }
                            break;
                        }
                        break;
                    case "tien2":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien2"].Value = 0;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            // if (this.ParseInt(e.Cell.Record.Cells["km_ck"].Value, 0) == 1)
                            //     this.UpdateKM();
                            break;
                        }
                        break;
                    case "tien":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["tien"].Value = (object)0;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = (object)this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            break;
                        }
                        break;
                    case "ck_nt":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["ck_nt"].Value = 0;
                            Decimal num3 = this.ParseDecimal(e.Cell.Record.Cells["ck_nt"].Value, new Decimal(0));
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
                            Decimal num4 = SysFunc.Round(num3 * num1, StartUpTrans.M_ROUND);
                            if (num4 != new Decimal(0))
                            {
                                e.Cell.Record.Cells["ck"].Value = num4;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
                            }
                            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                            {
                                e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"];
                            }
                            break;
                        }
                        break;
                    case "ck":
                        if (e.Cell.IsDataChanged)
                        {
                            if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                e.Cell.Record.Cells["ck"].Value = 0;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
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
                dataRow["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), result);
                dataRow["ma_ct"] = StartUpTrans.Ma_ct;
                dataRow["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                dataRow["ngay_ct"] = (this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                dataRow["so_luong"] = 0;
                dataRow["gia_nt"] = 0;
                dataRow["tien_nt"] = 0;
                dataRow["gia"] = 0;
                dataRow["tien"] = 0;
                dataRow["tl_ck"] = 0;
                dataRow["ck"] = 0;
                dataRow["ck_nt"] = 0;
                dataRow["gia_nt2"] = 0;
                dataRow["tien_nt2"] = 0;
                dataRow["gia2"] = 0;
                dataRow["tien2"] = 0;
                dataRow["ton13"] = DBNull.Value;
                dataRow["km_ck"] = 0;
                dataRow["ma_ctkm"] = "";
                dataRow["pt_km"] = (object)0;
                dataRow["tien_km"] = (object)0;
                dataRow["tien_km_nt"] = (object)0;
                dataRow["khuyen_mai"] = (object)false;
                int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                if (count > 0)
                {
                    dataRow["ma_kho_i"] = StartUpTrans.DsTrans.Tables[1].DefaultView[count - 1].Row["ma_kho_i"];
                    SqlCommand sqlCommand = new SqlCommand("select ma_dm from dmctct where ma_ct=@ma_ct and ma_dm=@ma_dm");
                    sqlCommand.Parameters.Add("@ma_ct", SqlDbType.Char).Value = StartUpTrans.Ma_ct;
                    sqlCommand.Parameters.Add("@ma_dm", SqlDbType.Char).Value = "dmvv";
                }
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
            if(this.GroupTotal.Visibility == Visibility.Visible)
            {
                this.txtMa_thue.SelectAllOnFocus = true;
                this.txtMa_thue.IsFocus = true;
            }    
            else
            {
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus()));
            }    
        }

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            if (Keyboard.IsKeyDown(Key.N) && Keyboard.Modifiers == ModifierKeys.Control)
            {
                this.NewRowCt();
                this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
            }
            if (!Keyboard.IsKeyDown(Key.Tab) || Keyboard.Modifiers != ModifierKeys.Control)
                return;
            this.txtMa_thue.IsFocus = true;
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            StartUpTrans.CheckPhanBo(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
            if (!this.IsInEditMode.Value)
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
                case Key.F5:
                    if (this.stt_mau_temlate == 226 ||this.stt_mau_temlate == 227)
                        break;
                    if (this.GrdCt.ActiveRecord != null && this.GrdCt.ActiveRecord.RecordType == RecordType.DataRecord)
                    {
                        CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell((this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"]);
                        if (cellValuePresenter != null && cellValuePresenter.Editor is ControlHostEditor editor)
                        {
                            AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(editor);
                            if (string.IsNullOrEmpty(autoCompleteControl.Text.Trim()))
                            {
                                int num1 = (int)ExMessageBox.Show(805, StartupBase.SasObj, "Chưa nhập mã vật tư!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            else if (autoCompleteControl != null)
                            {
                                if (autoCompleteControl.CheckLostFocus())
                                {
                                    string ma_vt = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_vt"].Value.ToString();
                                    string ten_vt = !StartUpTrans.M_LAN.Equals("V") ? (this.GrdCt.ActiveRecord as DataRecord).Cells["ten_vt2"].Value.ToString() : (this.GrdCt.ActiveRecord as DataRecord).Cells["ten_vt"].Value.ToString();
                                    string ma_kho = (this.GrdCt.ActiveRecord as DataRecord).Cells["ma_kho_i"].Value.ToString();
                                    object ngay_ct = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ngay_ct"];
                                    DataTable pn = StartUp.GetPN(ma_vt, ma_kho, ngay_ct);
                                    if (pn.Rows.Count > 0)
                                    {
                                        FrmSocthda_Pn frmSocthdaPn = new FrmSocthda_Pn(pn, ten_vt);
                                        if (this.stt_mau_temlate == 226 || this.stt_mau_temlate == 227)
                                        {
                                            if (frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia")))
                                            {
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia"].Visibility = Visibility.Hidden;
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia"].Settings.CellMaxWidth = 0.0;
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia"].Settings.CellWidth = 0.0;
                                            }
                                            if (frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "gia_nt")))
                                            {
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
                                                frmSocthdaPn.GrdSOCTHDA_PN.FieldLayouts[0].Fields["gia_nt"].Settings.CellWidth = 0.0;
                                            }
                                        }
                                        frmSocthdaPn.ShowDialog();
                                        int index = this.GrdCt.ActiveRecord.Index;
                                        if (index >= 0 && index < this.GrdCt.Records.Count)
                                        {
                                            DataRowView drvFrmSoctpnfPn = frmSocthdaPn.drvFrmSOCTPNF_PN;
                                            if (drvFrmSoctpnfPn != null)
                                            {
                                                if (StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0))
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmSoctpnfPn["gia"];
                                                else
                                                    (this.GrdCt.DataSource as DataView)[index]["gia_nt"] = drvFrmSoctpnfPn["gia_nt"];
                                                (this.GrdCt.DataSource as DataView)[index]["gia"] = drvFrmSoctpnfPn["gia"];
                                                if (this.ParseInt((this.GrdCt.DataSource as DataView)[index]["gia_ton"], 0) == 1 || this.ParseInt((this.GrdCt.DataSource as DataView)[index]["gia_ton"], 0) == 4)
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["px_gia_dd"] = 1;
                                                this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                                                this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"], 0);
                                                Decimal num2 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["so_luong"], new Decimal(0));
                                                Decimal num3 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["gia_nt"], new Decimal(0));
                                                Decimal num4 = this.ParseDecimal((this.GrdCt.DataSource as DataView)[index]["gia"], new Decimal(0));
                                                Decimal num5 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND_NT);
                                                if (num5 != new Decimal(0))
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["tien_nt"] = num5;
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                                                }
                                                Decimal num6 = SysFunc.Round(num2 * num4, StartUpTrans.M_ROUND);
                                                if (num6 != new Decimal(0))
                                                {
                                                    (this.GrdCt.DataSource as DataView)[index]["tien"] = num6;
                                                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        int num7 = (int)ExMessageBox.Show(810, StartupBase.SasObj, "Không có phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    }
                                }
                                else
                                {
                                    int num8 = (int)ExMessageBox.Show(815, StartupBase.SasObj, "Không có phiếu nhập cho vật tư này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                }
                            }
                        }
                        break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(820, StartupBase.SasObj, "Có xoá dòng ghi hiện thời?", "Xoá dòng", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                        break;
                    DataRow row = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow];
                    if (this.GrdCt.ActiveRecord is DataRecord activeRecord1)
                    {
                        int num = this.GrdCt.ActiveCell == null ? 0 : this.GrdCt.ActiveCell.Field.Index;
                        int index = activeRecord1.Index;
                        this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndDiscardChanges);
                        if (num >= 0)
                        {
                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[index].Row);
                            StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                            row["t_so_luong"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 0);
                            row["t_tien_nt2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 0);
                            row["t_tien2"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 0);
                            row["t_tien_nt"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt", 0);
                            row["t_tien"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien", 0);
                            this.UpdateKM();
                            if (index == 0 && this.GrdCt.Records.Count == 0)
                                this.GrdCt_AddNewRecord((object)null, (EditModeEndedEventArgs)null);
                            if (this.GrdCt.Records.Count > 0)
                                this.GrdCt.ActiveRecord = this.GrdCt.Records[index > this.GrdCt.Records.Count - 1 ? this.GrdCt.Records.Count - 1 : index];
                        }
                        break;
                    }
                    break;
            }
        }

        private Decimal SumFunction(DataTable datatable, string columnname, int km)
        {
            Decimal num = new Decimal(0);
            Decimal? nullable = datatable.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(r => r.RowState != DataRowState.Deleted)).Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() && b.Field<string>("km_ck") == km.ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>(columnname)));
            if (nullable.HasValue)
                num = this.ParseDecimal((object)nullable, new Decimal(0));
            return num;
        }
        private Decimal SumFunction1(DataTable datatable, string columnname, bool km)
        {
            Decimal num = new Decimal(0);
            Decimal? nullable = datatable.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(r => r.RowState != DataRowState.Deleted)).Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() && b.Field<bool>("khuyen_mai") == km)).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>(columnname)));
            if (nullable.HasValue)
                num = this.ParseDecimal((object)nullable, new Decimal(0));
            return num;
        }
        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void ChkPx_gia_dd_Click(object sender, RoutedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGridByPx_gia_dd();
        }

        private void ChkSua_tkthue_Click(object sender, RoutedEventArgs e)
        {
            bool? isChecked = this.ChkSua_tkthue.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                return;
            this.txtTk_thue_no.IsFocus = true;
        }

        private void ChkSua_thue_Click(object sender, RoutedEventArgs e)
        {
            bool? isChecked = this.ChkSua_thue.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                this.txtT_thue_nt.Focus();
            this.UpdateTongThue();
        }

        private void ChkTinh_ck_Click(object sender, RoutedEventArgs e)
        {
            this.UpdateTongThue();
        }

        private void txtt_tien_nt2_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateMoney_NT();
        }

        private void UpdateMoney_NT()
        {
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];

            Decimal kmhd_nt = new Decimal(0);

            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"].ToString()))
            {
                kmhd_nt = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"]);
            }
            Decimal num1 = this.ParseDecimal(dataRowView["ty_gia"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(dataRowView["thue_suat"], new Decimal(0));
            this.ParseInt(dataRowView["sua_tien"], 0);
            int num3 = this.ParseInt(dataRowView["tinh_ck"], 0);
            int num4 = this.ParseInt(dataRowView["sua_thue"], 0);
            Decimal num5 = this.ParseDecimal(dataRowView["t_tien_nt2"], new Decimal(0));
            Decimal num6 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck_nt", 0);
            Decimal num9 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_km_nt", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck_nt"] = num6 + kmhd_nt +num9;
            Decimal num7 = num5 - (num6 + kmhd_nt + num9);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_sau_ck_nt"] = num7;
            Decimal num8 = this.ParseDecimal(dataRowView["t_thue_nt"], new Decimal(0));
            if (num4 == 0)
            {
                num8 = num3 != 1 ? SysFunc.Round(num7 * num2 / new Decimal(100), StartUpTrans.M_ROUND_NT) : SysFunc.Round(num5 * num2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_nt"] = num8;
                if (dataRowView["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"] = num8;
            }
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"] = (num8 + num7);
            SysFunc.Round(num5 * num1, StartUpTrans.M_ROUND);
        }

        private void UpdateMoney()
        {
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            Decimal kmhd = new Decimal(0);
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"].ToString()))
            {
                kmhd = Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"]);
            }
            Decimal num1 = this.ParseDecimal(dataRowView["thue_suat"], new Decimal(0));
            this.ParseInt(dataRowView["sua_tien"], 0);
            int num2 = this.ParseInt(dataRowView["tinh_ck"], 0);
            int num3 = this.ParseInt(dataRowView["sua_thue"], 0);
            Decimal num4 = this.ParseDecimal(dataRowView["t_tien2"], new Decimal(0));
            Decimal num5 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "ck", 0);
            Decimal num9 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_km", 0);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_ck"] = num5 + kmhd + num9;
            Decimal num6 = num4 - (num5 + kmhd + num9);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_sau_ck"] = num6;
            Decimal num7 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"], new Decimal(0));
            if (num3 == 0)
            {
                num7 = num2 != 1 ? SysFunc.Round(num6 * num1 / new Decimal(100), StartUpTrans.M_ROUND) : SysFunc.Round(num4 * num1 / new Decimal(100), StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"] = num7;
            }
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"] = (num7 + num6);
        }

        private void txtt_tien2_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateMoney();
        }

        private void txtt_ck_nt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateMoney_NT();
        }

        private void txtt_ck_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateMoney();
        }

        private void txtT_thue_nt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateTongTT_NT();
        }

        private void txtT_thue_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.UpdateTongTT();
        }

        private void txtThue_suat_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (this.IsInEditMode == null || !this.IsInEditMode.Value)
                return;
            this.UpdateTongThue();
            this.UpdateThueKM();
        }

        private void txtMa_thue_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_thue.RowResult != null && this.txtMa_thue.IsDataChanged)
            {
                this.txtThue_suat.Value = this.ParseDecimal(this.txtMa_thue.RowResult["thue_suat"], new Decimal(0));
                this.txtTk_thue_co.Text = this.txtMa_thue.RowResult["tk_thue_co"].ToString().Trim();
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_dmthue"] = this.txtMa_thue.RowResult["tk_thue_co"].ToString().Trim();
                this.txtTk_thue_co.SearchInit();
                if (this.txtTk_thue_co.RowResult == null || string.IsNullOrEmpty(this.txtTk_thue_co.Text.Trim()))
                    return;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_cn"] = this.txtTk_thue_co.RowResult["tk_cn"];
            }
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtMa_thue.Text = this.txtMa_thue.Text.Trim()), DispatcherPriority.Background);
        }

        private void UpdateTongTT_NT()
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck_nt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"] = (num1 + num2);
            bool? isChecked = this.ChkSua_thue.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                return;
            Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ty_gia"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"] = SysFunc.Round(num1 * num3, StartUpTrans.M_ROUND);
        }

        private void UpdateTongTT()
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"] = (num1 + num2);
        }

        private void UpdateTongThue()
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            int num2 = this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_ck"], 0);
            int num3 = this.ParseInt(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_thue"], 0);
            Decimal num4 = new Decimal(0);
            Decimal num5 = new Decimal(0);
            Decimal num6;
            Decimal num7;
            if (num2 == 1)
            {
                Decimal num8 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"], new Decimal(0));
                Decimal num9 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"], new Decimal(0));
                num6 = SysFunc.Round(num8 * num1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                num7 = SysFunc.Round(num9 * num1 / new Decimal(100), StartUpTrans.M_ROUND);
            }
            else
            {
                Decimal num8 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck_nt"], new Decimal(0));
                Decimal num9 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck"], new Decimal(0));
                num6 = SysFunc.Round(num8 * num1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                num7 = SysFunc.Round(num9 * num1 / new Decimal(100), StartUpTrans.M_ROUND);
            }
            if (num3 != 0)
                return;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_nt"] = num6;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"] = num7;
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
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ton13"] = DBNull.Value;
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
                        listTon13.Rows[index]["ma_kho"] = str4;
                        string str5 = listTon13.Rows[index]["ma_vt"].ToString().Trim();
                        listTon13.Rows[index]["ma_vt"] = str5;
                        string str6 = listTon13.Rows[index]["ma_vv"].ToString().Trim();
                        listTon13.Rows[index]["ma_vv"] = str6;
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

        private Decimal GetTonKho(string ma_kho, string ma_vt, object ngay_ct)
        {
            Decimal num = new Decimal(0);
            try
            {
                string str = "ma_kho = '" + ma_kho + "'" + " AND ma_vt = '" + ma_vt + "'";
                SqlCommand sqlcmd = new SqlCommand("exec CheckTonXuatAm 1, @ngay_ct, @stt_rec, @advance");
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.SmallDateTime).Value = ngay_ct;
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["stt_rec"];
                sqlcmd.Parameters.Add("@advance", SqlDbType.VarChar).Value = str;
                DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                if (table.Rows.Count > 0)
                    num = this.ParseDecimal(table.Rows[0]["ton00"], new Decimal(0));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return num;
        }

        private void btnThongTin_Click(object sender, RoutedEventArgs e)
        {
            FrmTotalInfo frmTotalInfo = new FrmTotalInfo();
            frmTotalInfo.Title = StartUp.M_Tilte;
            frmTotalInfo.ShowDialog();
        }

        private void btnThongTinvc_Click(object sender, RoutedEventArgs e)
        {
            FrmThongTinvc frmThongTinvc = new FrmThongTinvc();
            frmThongTinvc.DataContext = StartUpTrans.DsTrans.Tables[0].DefaultView;
            frmThongTinvc.IsInEditMode.Value = this.IsInEditMode.Value;
            frmThongTinvc.ShowDialog();
        }

        private void UpdateTienKM_NT()
        {
            Decimal num1 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 1);
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km_nt"] = num1;
            Decimal num3 = new Decimal(0);
            if (StartUp.M_THUE_KM_CK == 1)
                num3 = SysFunc.Round(num1 * num2 / new Decimal(100), StartUpTrans.M_ROUND_NT);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km_nt"] = num3;
            Decimal num4 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc_nt"] = (num1 + num3);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_km_nt"] = (num4 + num1 + num3);
        }

        private void UpdateTienKM()
        {
            Decimal num1 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 1);
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km"] = num1;
            Decimal num3 = new Decimal(0);
            if (StartUp.M_THUE_KM_CK == 1)
                num3 = SysFunc.Round(num1 * num2 / new Decimal(100), StartUpTrans.M_ROUND);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km"] = num3;
            Decimal num4 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc"] = (num1 + num3);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_km"] = (num4 + num1 + num3);
        }

        private void UpdateThueKM()
        {
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_km_nt"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_km"], new Decimal(0));
            Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            Decimal num4 = new Decimal(0);
            Decimal num5 = new Decimal(0);
            Decimal num6 = SysFunc.Round(num1 * num3 / new Decimal(100), StartUpTrans.M_ROUND_NT);
            Decimal num7 = SysFunc.Round(num2 * num3 / new Decimal(100), StartUpTrans.M_ROUND);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km_nt"] = num6;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km"] = num7;
            this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"], new Decimal(0));
            this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc_nt"] = (num1 + num6);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc"] = (num2 + num7);
        }

        private void UpdateKM()
        {
            Decimal num1 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien_nt2", 1);
            Decimal num2 = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "tien2", 1);
            Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["thue_suat"], new Decimal(0));
            Decimal num4 = SysFunc.Round(num1 * num3 / new Decimal(100), StartUpTrans.M_ROUND);
            Decimal num5 = SysFunc.Round(num2 * num3 / new Decimal(100), StartUpTrans.M_ROUND);
            Decimal num6 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"], new Decimal(0));
            Decimal num7 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_sl_km"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 1);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km_nt"] = num1;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km"] = num2;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km_nt"] = num4;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km"] = num5;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc_nt"] = (num1 + num4);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc"] = (num2 + num5);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_km_nt"] = (num6 + num1 + num4);
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_km"] = (num7 + num2 + num5);

        }
        private void AddKhuyenmai()
        {
            if (allowKhuyenmai)
            {
                Tinhkhuyenmai();
                ReSum_all();
            }
        }
        private void txtT_thue_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtT_thue.IsFocusWithin || !this.IsInEditMode.Value || this.txtT_thue.Value != null && !(this.txtT_thue.Value.ToString() == ""))
                return;
            this.txtT_thue.Value = (object)0;
        }

        private void txtT_thue_nt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtT_thue_nt.IsFocusWithin || !this.IsInEditMode.Value)
                return;
            if (this.txtT_thue_nt.Value == null || this.txtT_thue_nt.Value.ToString() == "")
                this.txtT_thue_nt.Value = 0;
            if (StartUpTrans.M_ma_nt0 == this.txtMa_nt.Text.Trim())
                StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_nt"];
        }

        private void txtt_so_luong_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_sl_km"] = this.SumFunction(StartUpTrans.DsTrans.Tables[1], "so_luong", 1);
        }

        private void txtma_bp_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.IsInEditMode.Value || StartUp.M_BP_BH != 0)
                return;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtTk_thue_no_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.IsEditMode)
                return;
            bool? isChecked = this.ChkSua_tkthue.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                return;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtma_bp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            this.txtTen_bp.Text = this.txtma_bp.RowResult != null ? (StartUpTrans.M_LAN == "V" ? this.txtma_bp.RowResult["ten_bp"].ToString() : this.txtma_bp.RowResult["ten_bp2"].ToString()) : "";
        }

        private void txtTk_thue_co_PreviewGotFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!this.IsEditMode)
                return;
            if (this.txtTk_thue_co.IsReadOnly)
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtTk_thue_co.Text = this.txtTk_thue_co.Text.Trim()), DispatcherPriority.Background);
        }

        private void FormMain_Closed(object sender, EventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.None || this.IsInEditMode == null || !this.IsInEditMode.Value)
                return;
            this.V_Huy();
        }

        private void PhanBoThueInCT_CBTien()
        {
            try
            {
                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count == 0)
                    return;
                Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"], new Decimal(0));
                Decimal num3 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"], new Decimal(0));
                Decimal num4 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"], new Decimal(0));
                Decimal num5 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"], new Decimal(0));
                Decimal num6 = num5;
                Decimal num7 = num4;
                bool? isChecked = this.ChkTinh_ck.IsChecked;
                if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                {
                    num6 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck"], new Decimal(0));
                    num7 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_sau_ck_nt"], new Decimal(0));
                }
                Decimal num8 = new Decimal(0);
                Decimal num9 = new Decimal(0);
                Decimal num10 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_km_nt"], new Decimal(0));
                Decimal num11 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_km"], new Decimal(0));
                Decimal num12 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_km_nt"], new Decimal(0));
                Decimal num13 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_km"], new Decimal(0));
                Decimal num14 = SysFunc.Round(num4 * num1, StartUpTrans.M_ROUND);
                Decimal num15 = SysFunc.Round(num10 * num1, StartUpTrans.M_ROUND);
                Decimal num16 = new Decimal(0);
                Decimal num17 = new Decimal(0);
                Decimal num18 = new Decimal(0);
                Decimal num19 = new Decimal(0);
                bool flag1 = false;
                bool flag2 = false;
                int num20;
                if (!(num1 == new Decimal(0)))
                {
                    isChecked = this.ChkSua_tien.IsChecked;
                    num20 = (!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 ? 1 : 0;
                }
                else
                    num20 = 0;
                if (num20 == 0)
                {
                    flag1 = true;
                    flag2 = true;
                }
                int index1 = -1;
                int index2 = -1;
                Decimal num21 = new Decimal(0);
                Decimal num22 = new Decimal(0);
                Decimal num23 = new Decimal(0);
                Decimal num24 = new Decimal(0);
                for (int index3 = 0; index3 < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index3)
                {
                    Decimal num25 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["tien_nt2"], new Decimal(0));
                    Decimal num26 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["tien2"], new Decimal(0));
                    if (num25 != new Decimal(0) && num26 != new Decimal(0))
                    {
                        if (this.ParseInt(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["km_ck"], 0) == 0 && !flag1)
                        {
                            num26 += num14 - num5;
                            num5 = num14;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["tien2"] = num26;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = num5;
                            flag1 = true;
                        }
                        else if (!flag2)
                        {
                            num26 += num15 - num11;
                            num11 = num15;
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["tien2"] = num26;
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km"] = num11;
                            flag2 = true;
                        }
                    }
                    if (this.ParseInt(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["km_ck"], 0) == 0)
                    {
                        Decimal num27 = new Decimal(0);
                        Decimal num28 = new Decimal(0);
                        isChecked = this.ChkTinh_ck.IsChecked;
                        Decimal num29;
                        Decimal num30;
                        if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                        {
                            num29 = num26;
                            num30 = num25;
                        }
                        else
                        {
                            num29 = num26 - this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["ck"], new Decimal(0));
                            num30 = num25 - this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["ck_nt"], new Decimal(0));
                        }
                        Decimal num31;
                        Decimal num32;
                        if (this.txtMa_nt.Text != StartUpTrans.M_ma_nt0)
                        {
                            num31 = !(num30 == new Decimal(0)) ? (num7 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num30 / num7 * num2, StartUpTrans.M_ROUND_NT)) : new Decimal(0);
                            num32 = !(num29 == new Decimal(0)) ? (num6 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num29 / num6 * num3, StartUpTrans.M_ROUND)) : new Decimal(0);
                        }
                        else
                        {
                            num31 = num7 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num30 / num7 * num2, StartUpTrans.M_ROUND_NT);
                            num32 = num6 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num29 / num6 * num3, StartUpTrans.M_ROUND);
                        }
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["thue_nt"] = num31;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["thue"] = num32;
                        num21 += num32;
                        num22 += num31;
                        if (index1 == -1)
                            index1 = index3;
                    }
                    else
                    {
                        Decimal num27;
                        Decimal num28;
                        if (this.txtMa_nt.Text != StartUpTrans.M_ma_nt0)
                        {
                            num27 = !(num25 == new Decimal(0)) ? (num10 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num25 / num10 * num12, StartUpTrans.M_ROUND_NT)) : new Decimal(0);
                            num28 = !(num26 == new Decimal(0)) ? (num11 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num26 / num11 * num13, StartUpTrans.M_ROUND)) : new Decimal(0);
                        }
                        else
                        {
                            num27 = num10 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num25 / num10 * num12, StartUpTrans.M_ROUND_NT);
                            num28 = num11 == new Decimal(0) ? new Decimal(0) : SysFunc.Round(num26 / num11 * num13, StartUpTrans.M_ROUND);
                        }
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["thue_nt"] = num27;
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index3]["thue"] = num28;
                        num23 += num28;
                        num24 += num27;
                        if (index2 == -1)
                            index2 = index3;
                    }
                }
                if (index1 != -1)
                {
                    if (num3 != num21)
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index1]["thue"] = (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index1]["thue"], new Decimal(0)) + (num3 - num21));
                    if (num2 != num22)
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index1]["thue_nt"] = (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index1]["thue_nt"], new Decimal(0)) + (num2 - num22));
                }
                if (index2 != -1)
                {
                    if (num13 != num23)
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index2]["thue"] = (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index2]["thue"], new Decimal(0)) + (num13 - num23));
                    if (num12 != num24)
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index2]["thue_nt"] = (this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index2]["thue_nt"], new Decimal(0)) + (num12 - num24));
                }
                StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtT_tt_nt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt_nt"], new Decimal(0));
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km_nt"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km_nt"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc_nt"] = (num1 + num2);
        }

        private void txtT_tt_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tt"], new Decimal(0));
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_km"], new Decimal(0));
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_thue_km"], new Decimal(0));
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_tc"] = (num1 + num2);
        }

        private void txtHan_tt_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtHan_tt.SelectAll();
        }

        private void txtHt_tt_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (!this.IsInEditMode.Value)
                return;
            try
            {
                FrmSOCTHDAHdm frmSocthdaHdm = new FrmSOCTHDAHdm();
                frmSocthdaHdm.txtMa_kh.Text = this.txtMa_kh.Text;
                frmSocthdaHdm.tblTen_kh.Text = this.txtTen_kh.Text;
                frmSocthdaHdm.ShowDialog();
                if (frmSocthdaHdm.isOk)
                {
                    int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                    for (int index = 0; index < count; ++index)
                        StartUpTrans.DsTrans.Tables[1].DefaultView.Delete(0);
                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    string upper1 = this.txtMa_nt.Text.ToUpper();
                    string upper2 = frmSocthdaHdm.dsHdm.Tables[0].DefaultView[0]["ma_nt"].ToString().ToUpper();
                    for (int index = 0; index < frmSocthdaHdm.dsHdm.Tables[1].DefaultView.Count; ++index)
                    {
                        DataRow row1 = frmSocthdaHdm.dsHdm.Tables[1].DefaultView[index].Row;
                        DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                        DataTable table = frmSocthdaHdm.dsHdm.Tables[1].Clone();
                        table.Rows.Add(row1.ItemArray);
                        DataTable dataTable = StartUpTrans.DsTrans.Tables[1].Clone();
                        dataTable.Merge(table, true, MissingSchemaAction.Ignore);
                        if (dataTable.Rows.Count > 0)
                            row2.ItemArray = dataTable.Rows[0].ItemArray;
                        Decimal result1 = new Decimal(0);
                        Decimal.TryParse(row1["tl_ck"].ToString(), out result1);
                        Decimal result2 = new Decimal(0);
                        Decimal.TryParse(row1["so_luong"].ToString(), out result2);
                        if (upper1.Equals(upper2))
                        {
                            if (upper1.Equals(StartUpTrans.M_ma_nt0))
                            {
                                row2["gia_nt2"] = row1["gia2"];
                                row2["gia2"] = row1["gia2"];
                                row2["tien_nt2"] = row1["tien2"];
                                row2["tien2"] = row1["tien2"];
                                row2["ck_nt"] = row1["ck_nt"];
                                row2["ck"] = row1["ck"];
                            }
                            else
                            {
                                row2["gia_nt2"] = row1["gia_nt2"];
                                row2["gia2"] = SysFunc.Round(Convert.ToDecimal(row2["gia_nt2"]) * this.txtTy_gia.nValue, StartUpTrans.M_ROUND_GIA);
                                row2["tien_nt2"] = row1["tien_nt2"];
                                row2["tien2"] = SysFunc.Round(Convert.ToDecimal(row2["gia2"]) * result2, StartUpTrans.M_ROUND);
                                row2["ck_nt"] = row1["ck_nt"];
                                row2["ck"] = SysFunc.Round(Convert.ToDecimal(row2["tien2"]) * result1 / new Decimal(100), StartUpTrans.M_ROUND);
                            }
                        }
                        else if (upper1.Equals(StartUpTrans.M_ma_nt0))
                        {
                            row2["gia_nt2"] = row1["gia2"];
                            row2["gia2"] = row1["gia2"];
                            row2["tien_nt2"] = row1["tien2"];
                            row2["tien2"] = row1["tien2"];
                            row2["ck_nt"] = row1["ck_nt"];
                            row2["ck"] = row1["ck"];
                        }
                        else
                        {
                            row2["gia_nt2"] = SysFunc.Round(Convert.ToDecimal(row1["gia2"]) / this.txtTy_gia.nValue, StartUpTrans.M_ROUND_GIA_NT);
                            row2["gia2"] = row1["gia2"];
                            row2["tien_nt2"] = SysFunc.Round(Convert.ToDecimal(row2["gia_nt2"]) * result2, StartUpTrans.M_ROUND_NT);
                            row2["tien2"] = row1["tien2"];
                            row2["ck_nt"] = SysFunc.Round(Convert.ToDecimal(row2["tien_nt2"]) * result1 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                            row2["ck"] = row1["ck"];
                        }
                        row2["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                        int result3 = 0;
                        if (this.GrdCt.Records.Count > 0)
                        {
                            string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                            if (str != null)
                                int.TryParse(str.ToString(), out result3);
                        }
                        int num = result3 + 1;
                        row2["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), num);
                        row2["km_ck"] = 0;
                        row2["ma_hd_i"] = row1["ma_hd"];
                        bool allowAdd = true;
                        if (!allowKhuyenmai)
                        {
                            row2["ma_ctkm"] = "";
                            row2["tien_km"] = 0;
                            row2["ck"] = 0;
                            row2["ck_nt"] = 0;
                            row2["tien_km_nt"] = 0;
                            row2["ma_nh_km"] = "";
                            row2["pt_km"] = 0;
                            row2["khuyen_mai"] = (object)false;
                            if (Convert.ToBoolean(row1["khuyen_mai"]))
                            {
                                allowAdd = false;
                            }
                        }
                        if (allowAdd)
                            StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                    }
                    if (allowKhuyenmai)
                    {
                        if (frmSocthdaHdm.dsHdm.Tables[0].Rows.Count > 0)
                        {
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_ctkm"] = frmSocthdaHdm.dsHdm.Tables[0].Rows[0]["ma_ctkm"];
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_km"] = frmSocthdaHdm.dsHdm.Tables[0].Rows[0]["tien_km"];
                            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["tien_km_nt"] = frmSocthdaHdm.dsHdm.Tables[0].Rows[0]["tien_km_nt"];
                        }
                    }

                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_so_luong"] = this.SumFunction1(StartUpTrans.DsTrans.Tables[1], "so_luong", false);
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien_nt2"] = this.SumFunction1(StartUpTrans.DsTrans.Tables[1], "tien_nt2", false);
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_tien2"] = this.SumFunction1(StartUpTrans.DsTrans.Tables[1], "tien2", false);
                    this.UpdateTonKho();
                    this.UpdateMoney_NT();
                    this.UpdateMoney();
                    this.GrdCt.Focus();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtSo_seri_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtSo_seri.Text = this.txtSo_seri.Text.Trim();
        }

        private void txtma_thck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtma_thck.RowResult == null || this.txtHan_tt.Value != DBNull.Value && !(this.txtHan_tt.nValue == new Decimal(0)))
                return;
            this.txtHan_tt.Value = this.txtma_thck.RowResult["han_tt"];
        }

        private void txtHan_tt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!this.txtHan_tt.IsFocusWithin)
            {

            }
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void txtHt_tt_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtHt_tt.SelectionStart = 0;
            this.txtHan_tt.SelectionLength = 0;
        }

        public void DeleteVoucherPT(string _stt_rec, string _ma_ct)
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
        private void btnViewPT1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()))
                    return;
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim().Equals("PT1"))
                {
                    SqlCommand sqlcmd1 = new SqlCommand();
                    sqlcmd1.CommandText = "Select count(1) from ph41 WHERE stt_Rec = @stt_rec_pt";
                    sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString();
                    if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                    {
                        if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu thu không tồn tại, có xóa thông tin phiếu thu trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) != MessageBoxResult.Yes)
                            return;
                        SqlCommand sqlcmd2 = new SqlCommand();
                        sqlcmd2.CommandText = "UPDATE ph81 Set stt_rec_pt = '', so_ct_pt = '', ma_ct_pt = '' WHERE stt_rec = @stt_rec_pt; ";
                        sqlcmd2.CommandText += "UPDATE cttt20 Set tat_toan = 0, stt_rec_tt = '' WHERE stt_rec = @stt_rec";
                        sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                        sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                        this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                    }
                    else
                    {
                        SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "PT1", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString(), Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), this.BindingSasObj.M_ProcessName);
                    }

                }
                else
                {
                    SqlCommand sqlcmd1 = new SqlCommand();
                    sqlcmd1.CommandText = "Select count(1) from ph51 WHERE stt_Rec = @stt_rec_pt";
                    sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString();
                    if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                    {
                        if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu thu không tồn tại, có xóa thông tin phiếu thu trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                        {
                            SqlCommand sqlcmd2 = new SqlCommand();
                            sqlcmd2.CommandText = "UPDATE ph81 Set stt_rec_pt = '', so_ct_pt = '', ma_ct_pt = '' WHERE stt_rec = @stt_rec_pt; ";
                            sqlcmd2.CommandText += "UPDATE cttt20 Set tat_toan = 0, stt_rec_tt = '' WHERE stt_rec = @stt_rec";
                            sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                            sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                            this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                        }
                    }
                    else
                        SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "BC1", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString(), Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), this.BindingSasObj.M_ProcessName);
                }

            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void TabInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.TabInfo.SelectedIndex != 1 || !this.IsInEditMode.Value)
                return;
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtso_hd.Focus()), DispatcherPriority.Background);
        }

        private void GrdCt_CellDeactivating(object sender, CellDeactivatingEventArgs e)
        {
            try
            {
                if (!this.IsInEditMode.Value || e.Cell.Field.Name != "ma_vt")
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
                        e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl1.RowResult["tk_vt"].ToString();
                }
                AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["ma_kho_i"]).Editor as ControlHostEditor);
                if (autoCompleteControl2 != null)
                {
                    autoCompleteControl2.SearchInit();
                    if (autoCompleteControl2.RowResult != null && (autoCompleteControl2.RowResult["tk_dl"] != DBNull.Value && !string.IsNullOrEmpty(autoCompleteControl2.RowResult["tk_dl"].ToString().Trim())))
                    {
                        e.Cell.Record.Cells["tk_vt"].Value = autoCompleteControl2.RowResult["tk_dl"].ToString();
                        dataItem["tk_vt_dmvt"] = autoCompleteControl2.RowResult["tk_dl"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void txtMa_thue_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtMa_thue.Text = this.txtMa_thue.Text.Trim()), DispatcherPriority.Background);
        }

        private void FormMain_PreviewKeyDown(object sender, KeyEventArgs e)
        {
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void txtsd_hddt_yn_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtsd_hddt_yn.Text))
                return;
            this.txtsd_hddt_yn.Text = "0";
        }

        private void Sum_all()
        {
            Decimal num1 = new Decimal(0);
            Decimal num2 = new Decimal(0);
            Decimal num3 = new Decimal(0);
            string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            foreach (DataRow row in (InternalDataCollectionBase)StartUpTrans.DsTrans.Tables[1].Rows)
            {
                if (!(row["stt_rec"].ToString() != str) && !(row["km_ck"].ToString() == "1"))
                {
                    num3 += this.ParseDecimal(row["so_luong"], new Decimal(0));
                    num2 += this.ParseDecimal(row["tien_nt2"], new Decimal(0));
                    num1 += this.ParseDecimal(row["tien2"], new Decimal(0));
                }
            }
            if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                num1 = num2;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["t_so_luong"] = num3;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"] = num2;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = num1;
            this.UpdateTongThue();
            this.UpdateTienKM();
            this.UpdateTienKM_NT();
            this.UpdateThueKM();

        }
        private void ReSum_all()
        {
            UpdateMoney();
            UpdateMoney_NT();          
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
                    kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"]).ToString("yyyy-MM-dd");
                    DataTable tblhd = StartupBase.SasObj.ExcuteReader(kmcmd).Tables[0];
                    if (tblhd.Rows.Count > 0)
                    {
                        foreach (DataRow nr in tblhd.Rows)
                        {
                            DataRow r = StartUp.CTKMTable.NewRow();
                            r["stt_rec"] = nr["stt_rec"];
                            r["ischoose"] = (object)true;
                            r["pt_ck"] = nr["pt_ck"];
                            r["tien_ck"] = nr["tien_ck"];
                            r["loai"] = 1;
                            r["loai_km"] = nr["loai_km"];
                            r["ten_km"] = (object)"KMHD: " + nr["ten_km"].ToString().Trim();
                            r["ngay_bd"] = nr["ngay_bd"];
                            r["ngay_kt"] = nr["ngay_kt"];
                            StartUp.CTKMTable.Rows.Add(r);
                        }
                    }
                }
                if (!string.IsNullOrEmpty(ma_kh) && StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable().Rows)
                    {
                        if (!string.IsNullOrEmpty(r["khuyen_mai"].ToString()) && !Convert.ToBoolean(r["khuyen_mai"]))
                        {
                            String ma_vt = r["ma_vt"].ToString();
                            string sql = "Exec checkkmtheohanghoa @ma_vt, @ma_kh,@ngay";
                            SqlCommand kmcmd = new SqlCommand(sql);
                            kmcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)ma_vt.Trim();
                            kmcmd.Parameters.Add("@ma_kh", SqlDbType.Char).Value = (object)ma_kh.Trim();
                            kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"]).ToString("yyyy-MM-dd");
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
                        if (!string.IsNullOrEmpty(r["khuyen_mai"].ToString()) && !Convert.ToBoolean(r["khuyen_mai"]))
                        {
                            String ma_vt = r["ma_vt"].ToString();
                            Double so_Luong = Convert.ToDouble(r["so_luong"]);
                            string sql = "Exec checkkmhangtanghang @ma_vt, @ma_kh, @so_luong,@ngay";
                            SqlCommand kmcmd = new SqlCommand(sql);
                            kmcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)ma_vt.Trim();
                            kmcmd.Parameters.Add("@ma_kh", SqlDbType.Char).Value = (object)ma_kh.Trim();
                            kmcmd.Parameters.Add("@so_luong", SqlDbType.Decimal).Value = (object)so_Luong;
                            kmcmd.Parameters.Add("@ngay", SqlDbType.Char).Value = (object)Convert.ToDateTime(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"]).ToString("yyyy-MM-dd");
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
                this.GrdCtKM.DataSource = StartUp.CTKMTable.DefaultView;
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
            try
            {
                if (Convert.ToBoolean(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"]))
                    ThemKhuyenmai();
            }
            catch
            {

            }
        }
        int newstt_rec0 = 0;         
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
                                            string sql_tkdt = "Select gia_ton,tk_vt,tk_dt,tk_gv from dmvt where ma_vt = '" + rvt[y]["ma_vt"].ToString().Trim() + "'";
                                            DataTable tbltk = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql_tkdt)).Tables[0];
                                            if (tbltk.Rows.Count > 0)
                                            {
                                                rnew["gia_ton"] = tbltk.Rows[0]["gia_ton"];
                                                rnew["tk_vt"] = tbltk.Rows[0]["tk_vt"];
                                                rnew["tk_dt"] = tbltk.Rows[0]["tk_dt"];
                                                rnew["tk_gv"] = tbltk.Rows[0]["tk_gv"];
                                            }
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
        }      
        void ResetKhuyenmai()
        {
            foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].Rows)
            {
                r["ma_nh_km"] = "";
                r["ma_ctkm"] = "";
                r["pt_km"] = 0;
                r["tien_km"] = 0;
                r["tien_km_nt"] = 0;              
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

        //HDDT
        private void btnAddInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsInEditMode.Value)
                return;

            FrmWaiting frmWaiting = new FrmWaiting(600);
            try
            {
                frmWaiting.Show();
                frmWaiting.Set(100);
                string voucherIds = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                string msg = string.Empty;
                string sGUID = string.Empty;
                string sInvoiceForm = string.Empty;
                string sInvoiceSerial = string.Empty;
                string sInvoiceNo = string.Empty;
                DataSet ds = null;
                string format = "exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter,@sl_ct;";
                using (SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 0 ? string.Format(format, "[LoadVoucher]") : string.Format(format, StartUpTrans.Process_Store[0])))
                {
                    sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = StartUpTrans.Ma_ct;
                    sqlcmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = "stt_rec = '" + voucherIds + "'";
                    sqlcmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = "1=1";
                    sqlcmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = "1=1";
                    sqlcmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = 1;
                    ds = DataProvider.FillCommand(StartupBase.SasObj, sqlcmd);
                }
                frmWaiting.Set(200);
                if (ds != null)
                {
                    SasCusEinvoice.UpdateInvoice updateInvoice = new SasCusEinvoice.UpdateInvoice(StartupBase.SasObj, ds.Copy(), "BKAV", "HDA");
                    ds.Clear();
                    ds.Dispose();
                    msg = updateInvoice.BKAV_Create(out sGUID, out sInvoiceForm, out sInvoiceSerial, out sInvoiceNo);
                    frmWaiting.Set(300);
                    if (msg.Length > 0)
                    {
                        MessageBox.Show(msg, "Thong bao");
                    }
                    else
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["inv_guid"] = sGUID.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_hddt"] = sInvoiceNo.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri_hddt"] = sInvoiceSerial.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["mau_hddt"] = sInvoiceForm.ToString().Trim();

                        SqlCommand sqlcmd1 = new SqlCommand();
                        sqlcmd1.CommandText = "Update PH81 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CTTT20 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CT00 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CT70 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CTGT20 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.Parameters.Add("@sGUID", SqlDbType.NVarChar).Value = sGUID.ToString().Trim();
                        sqlcmd1.Parameters.Add("@sInvoiceForm", SqlDbType.NVarChar).Value = sInvoiceForm.ToString().Trim();
                        sqlcmd1.Parameters.Add("@sInvoiceSerial", SqlDbType.NVarChar).Value = sInvoiceSerial.ToString().Trim();
                        sqlcmd1.Parameters.Add("@sInvoiceNo", SqlDbType.NVarChar).Value = sInvoiceNo.ToString().Trim();

                        if (StartupBase.SasObj.ExcuteNonQuery(sqlcmd1) > 0)
                            MessageBox.Show("Thêm mới thành công: " + sGUID.ToString().Trim(), "Thong bao");
                        else
                            MessageBox.Show("Thêm mới thành công: " + sGUID.ToString().Trim() + " - Lỗi cập nhật mẫu số, ký hiệu số hóa đơn điện tử vào phần mềm?!", "Thong bao");

                    }
                }
                else
                {
                    MessageBox.Show("Lỗi, không tìm thấy dữ liệu để cập nhật sang HĐĐT: Stt_rec = " + voucherIds.ToString().Trim(), "Thong bao");
                }
                frmWaiting.Set(500);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            frmWaiting.Set(600);
            frmWaiting.Close();
            frmWaiting = null;
        }

        private void btnAddAllInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsInEditMode.Value)
                return;

            FrmWaiting frmWaiting = new FrmWaiting(600);
            try
            {
                frmWaiting.Show();
                frmWaiting.Set(100);
                DataTable dterror = new DataTable();
                dterror.Columns.Add("tag", typeof(string));
                dterror.Columns.Add("so_ct", typeof(string));
                dterror.Columns.Add("error", typeof(string));
                foreach (DataRow row in StartUp.DsTrans.Tables[0].Rows)
                {
                    string voucherIds = row["stt_rec"].ToString();
                    if (string.IsNullOrEmpty(voucherIds))
                        continue;

                    DataRow rowError = dterror.NewRow();
                    string msg = string.Empty;
                    string sGUID = string.Empty;
                    string sInvoiceForm = string.Empty;
                    string sInvoiceSerial = string.Empty;
                    string sInvoiceNo = string.Empty;
                    DataSet ds = null;
                    string format = "exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter,@sl_ct;";
                    using (SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 0 ? string.Format(format, "[LoadVoucher]") : string.Format(format, StartUpTrans.Process_Store[0])))
                    {
                        sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = StartUpTrans.Ma_ct;
                        sqlcmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = "stt_rec = '" + voucherIds + "'";
                        sqlcmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = "1=1";
                        sqlcmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = "1=1";
                        sqlcmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = 1;
                        ds = DataProvider.FillCommand(StartupBase.SasObj, sqlcmd);
                    }

                    if (ds != null)
                    {
                        SasCusEinvoice.UpdateInvoice updateInvoice = new SasCusEinvoice.UpdateInvoice(StartupBase.SasObj, ds.Copy(), "BKAV", "HDA");
                        ds.Clear();
                        ds.Dispose();
                        msg = updateInvoice.BKAV_Create(out sGUID, out sInvoiceForm, out sInvoiceSerial, out sInvoiceNo);
                        frmWaiting.Set(300);
                        if (msg.Length > 0)
                        {
                            rowError["tag"] = "1";
                            rowError["so_ct"] = row["so_ct"].ToString();
                            rowError["error"] = msg + " : Stt_rec = " + voucherIds.ToString().Trim();
                        }
                        else
                        {
                            row["inv_guid"] = sGUID.ToString().Trim();
                            row["so_ct_hddt"] = sInvoiceNo.ToString().Trim();
                            row["so_seri_hddt"] = sInvoiceSerial.ToString().Trim();
                            row["mau_hddt"] = sInvoiceForm.ToString().Trim();

                            SqlCommand sqlcmd1 = new SqlCommand();
                            sqlcmd1.CommandText = "Update PH81 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CTTT20 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CT00 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CT70 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CTGT20 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.Parameters.Add("@sGUID", SqlDbType.NVarChar).Value = sGUID.ToString().Trim();
                            sqlcmd1.Parameters.Add("@sInvoiceForm", SqlDbType.NVarChar).Value = sInvoiceForm.ToString().Trim();
                            sqlcmd1.Parameters.Add("@sInvoiceSerial", SqlDbType.NVarChar).Value = sInvoiceSerial.ToString().Trim();
                            sqlcmd1.Parameters.Add("@sInvoiceNo", SqlDbType.NVarChar).Value = sInvoiceNo.ToString().Trim();

                            if (StartupBase.SasObj.ExcuteNonQuery(sqlcmd1) > 0)
                            {
                                rowError["tag"] = "0";
                                rowError["so_ct"] = row["so_ct"].ToString();
                                rowError["error"] = "Thêm mới thành công: " + sGUID.ToString().Trim();
                            }
                            else
                            {
                                rowError["tag"] = "1";
                                rowError["so_ct"] = row["so_ct"].ToString();
                                rowError["error"] = "Thêm mới thành công: " + sGUID.ToString().Trim() + " - Lỗi cập nhật mẫu số, ký hiệu số hóa đơn điện tử vào phần mềm?!";
                            }
                        }
                    }
                    else
                    {
                        rowError["tag"] = "1";
                        rowError["so_ct"] = row["so_ct"].ToString();
                        rowError["error"] = "không tìm thấy dữ liệu để cập nhật sang HĐĐT: Stt_rec = " + voucherIds.ToString().Trim();
                    }

                    dterror.Rows.Add(rowError);
                }

                frmWaiting.Set(500);
                if (StartupBase.M_LAN == "V")
                {
                    BrowseError(dterror, "so_ct:100:H=Số chứng từ; error:400:h=Diễn giải", "Danh sách số chứng từ đẩy sang hóa đơn điện tử bị lỗi");
                }
                else
                {
                    BrowseError(dterror, "so_ct:100:H=Invoice No.; error:400:h=Description", "List of failed e-invoices push to e-invoices");
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            frmWaiting.Set(600);
            frmWaiting.Close();
            frmWaiting = null;
        }

        public static void BrowseError(DataTable data, string fields, string title)
        {
            FormBrowse oBrowse = new FormBrowse(StartupBase.SasObj, data.DefaultView, fields);
            oBrowse.SetRowColorByTag("tag", "1", System.Windows.Media.Colors.Red, false);//Chưa tạo phiếu
            oBrowse.frmBrw.Title = "Thong bao day du lieu len hoa don dien tu";
            oBrowse.frmBrw.LanguageID = "Socthda_viewError";
            oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)delegate
            {
                oBrowse.frmBrw.ChangeLanguage(StartupBase.M_LAN);
            });
            oBrowse.ShowDialog();
        }

        private void btnSendEmail_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsInEditMode.Value)
                return;
            try
            {
                string msg = "";
                string voucherIds = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                if (string.IsNullOrEmpty(voucherIds))
                    return;

                string folder = StartupBase.SasObj.M_StartUp_Path + "Email\\";
                SasFormReport.ReportLib.CreateDirectory(new DirectoryInfo(folder));
                string filename = System.IO.Path.Combine(folder, voucherIds.ToString().Trim() + ".pdf");
                if (!File.Exists(filename))
                {
                    SasCusEinvoice.UpdateInvoice updateInvoice = new SasCusEinvoice.UpdateInvoice(StartupBase.SasObj, (DataSet)null, "BKAV", "HDA");
                    msg = updateInvoice.BKAV_GetFilePDF(voucherIds, filename);
                    if (msg.Length > 0)
                    {
                        MessageBox.Show(msg, "Thông báo");
                        return;
                    }
                }
                if (!File.Exists(filename))
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
                    string EmailTo = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["e_mail"].ToString().Trim();
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        EmailFrom = dt.Rows[0]["e_mail"].ToString().Trim();
                        EmailPassFrom = SasUtilities.ClsEmail.DeCrypt(dt.Rows[0]["pass"].ToString().Trim(), "0123456789").ToString().Trim();
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
                        msg = SasUtilities.ClsEmail.sendmail(EmailFrom, EmailPassFrom, frm.txtTEmail.Text.ToString().Trim(), "", "", frm.txtTieude.Text.ToString().Trim(), frm.txtNoidung.Text.ToString().Trim().Trim(), filename);
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
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        //Xuất kho NVL
        private void BtnXuatNVL_Click_cu(object sender, RoutedEventArgs e)
        {
            if (this.IsInEditMode.Value)
                return;

            bool _createPXD = true;
            string newstt_recPXD = string.Empty;

            DataTable dataTable = (DataTable)null;
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString()))
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai,ma_kho FROM PH84 WHERE stt_rec LIKE '{0}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString());
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

            DataRow row = dt.NewRow();
            dt.Rows.Add(row);

            //if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString().Trim() == "")
            //{
            FrmTaoPXNVL frmTaoPXNVL = new FrmTaoPXNVL();
            frmTaoPXNVL.tbInfoPT = dataTable;
            frmTaoPXNVL.DataContext = (object)dt.DefaultView;
            frmTaoPXNVL.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"].ToString();
            frmTaoPXNVL.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"].ToString().Trim().PadLeft(frmTaoPXNVL.txtso_ct_pt.MaxLength);
            frmTaoPXNVL.txtnguoi_nop.Text = this.txtOng_ba.Text;
            if (Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"]) != 1 && Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"]) != 2)
            {
                frmTaoPXNVL.kind = 1;
            }
            else
            {
                frmTaoPXNVL.kind = Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"]);
            }

            frmTaoPXNVL.Ma_nt_ht = this.txtMa_nt.Text;
            frmTaoPXNVL.so_hd = this.txtSo_ct.Text.Trim();
            frmTaoPXNVL.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
            frmTaoPXNVL.ShowDialog();
            if (!frmTaoPXNVL.isOk)
            {
                _createPXD = false;
            }
            else
            {
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
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_px"] = frmTaoPXNVL.txtMa_qs_pt.Text;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_px"] = frmTaoPXNVL.txtso_ct_pt.Text.Trim().PadLeft(frmTaoPXNVL.txtso_ct_pt.MaxLength);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_px"] = "PXD";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"] = newstt_recPXD;
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"] = frmTaoPXNVL.txtKind.Text.Equals("1") ? 1 : 2;
            }
            //}


            string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
            this.so_ct_px_length = StartupBase.SasObj.GetDatabaseFieldLength("so_ct");
            new Thread((ThreadStart)(() =>
            {
                if (_createPXD)
                {
                    this.CreatePXD(dt);
                    //Kiểm tra xem đã tao hóa đơn thành công không
                    SqlCommand sqlcmd3 = new SqlCommand();
                    sqlcmd3.CommandText = string.Format("SELECT TOP 1 stt_rec FROM PH84 WHERE stt_rec LIKE '{0}' UNION ALL SELECT TOP 1 stt_rec FROM CT84 WHERE stt_rec LIKE '{0}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_px"].ToString());
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
                                dataRowArray[0]["loai_xnvl"] = dt.Rows[0]["loai_xnvl"];
                            }
                        }));

                        MessageBox.Show((StartUp.M_LAN == "E" ? "Successfully created NVL export slip!" : "Tạo phiếu xuất NVL thành công!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());

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
                        MessageBox.Show((StartUp.M_LAN == "E" ? "Create an NVL output slip with error. Please check the quota declaration!" : "Tạo phiếu xuất NVL bị lỗi.Hãy kiểm tra lại khai báo định mức!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());

                    }
                }
            })).Start();
        }

        private void BtnXuatNVL_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsInEditMode.Value)
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
            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString()) &&Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString().Trim()) != 1 && Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString().Trim()) != 2)
            {
                frmTaoPXNVL.kind = 1;
            }
            else
            {
                frmTaoPXNVL.kind = string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"].ToString()) ? 2 : Convert.ToInt32(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_xnvl"]);
            }

            frmTaoPXNVL.Ma_nt_ht = this.txtMa_nt.Text;
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

                if (isSuccess)
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Successfully created NVL export slip!" : "Tạo phiếu xuất NVL thành công!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                else
                    MessageBox.Show((StartUp.M_LAN == "E" ? "Create an NVL output slip with error. Please check the quota declaration!" : "Tạo phiếu xuất NVL bị lỗi.Hãy kiểm tra lại khai báo định mức!"), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());

            }
        }

        private void CreatePXD(DataTable dt)
        {
            string mparams = "'" + dt.Rows[0]["stt_rec"].ToString().Trim() + "','" + dt.Rows[0]["stt_recpx"].ToString().Trim() + "','" + dt.Rows[0]["ma_qs"].ToString().Trim() + "','" + dt.Rows[0]["so_ct"].ToString().Trim() + "','" + dt.Rows[0]["ma_nt"].ToString().Trim() + "','" + dt.Rows[0]["ty_gia"].ToString().Trim() + "','" + dt.Rows[0]["ty_giaf"].ToString().Trim() + "','" + dt.Rows[0]["nguoinop"].ToString().Trim() + "','" + dt.Rows[0]["lydonop"].ToString().Trim() + "','" + dt.Rows[0]["ma_gd"].ToString().Trim() + "','" + dt.Rows[0]["ma_ct"].ToString().Trim() + "','" + dt.Rows[0]["loai_xnvl"].ToString().Trim() + "','" + dt.Rows[0]["ma_kho"].ToString().Trim() + "','" + dt.Rows[0]["ma_nx"].ToString().Trim() + "'";
            Console.WriteLine(mparams);
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[SOCTHDA-CREATEPXD] @Stt_rec, @Stt_recpx, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct, @loai_xnvl, @ma_kho, @ma_nx");
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = dt.Rows[0]["stt_rec"];
                sqlcmd.Parameters.Add("@Stt_recpx", SqlDbType.VarChar).Value = dt.Rows[0]["stt_recpx"];
                sqlcmd.Parameters.Add("@ma_qs", SqlDbType.VarChar).Value = dt.Rows[0]["ma_qs"];
                sqlcmd.Parameters.Add("@so_ct", SqlDbType.VarChar).Value = dt.Rows[0]["so_ct"];
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nt"];
                sqlcmd.Parameters.Add("@ty_gia", SqlDbType.Decimal).Value = dt.Rows[0]["ty_gia"];
                sqlcmd.Parameters.Add("@ty_giaf", SqlDbType.Decimal).Value = dt.Rows[0]["ty_giaf"];
                sqlcmd.Parameters.Add("@nguoinop", SqlDbType.NVarChar).Value = dt.Rows[0]["nguoinop"];
                sqlcmd.Parameters.Add("@lydonop", SqlDbType.NVarChar).Value = dt.Rows[0]["lydonop"];
                sqlcmd.Parameters.Add("@ma_gd", SqlDbType.VarChar).Value = dt.Rows[0]["ma_gd"];
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = dt.Rows[0]["ma_ct"];
                sqlcmd.Parameters.Add("@loai_xnvl", SqlDbType.TinyInt).Value = dt.Rows[0]["loai_xnvl"];
                sqlcmd.Parameters.Add("@ma_kho", SqlDbType.VarChar).Value = dt.Rows[0]["ma_kho"];
                sqlcmd.Parameters.Add("@ma_nx", SqlDbType.VarChar).Value = dt.Rows[0]["ma_nx"].ToString().Trim();
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
                sqlcmd2.CommandText = "UPDATE ph81 Set stt_rec_px = '', so_ct_px = '', ma_ct_px = '', ma_ct_qs = '', loai_xnvl = 0 WHERE stt_rec = @stt_rec_px; ";
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

        private void BtnVnpay_Click(object sender, RoutedEventArgs e)
        {
            if (Vnpayrow != null)
            {
                if (!string.IsNullOrEmpty(Vnpayrow["order_id"].ToString()) && !string.IsNullOrEmpty(Vnpayrow["t_tt"].ToString()) && !string.IsNullOrEmpty(Vnpayrow["ma_kh"].ToString()))
                {
                    string tk_nx = Vnpayrow["ma_nx"].ToString().Trim();
                    string order_id = Vnpayrow["order_id"].ToString();
                    long so_tien = Convert.ToInt64(Vnpayrow["t_tt"]);
                    string ho_ten = Vnpayrow["ma_kh"].ToString();
                    string dien_thoai = "0900999999";
                    string e_mail = "sisvn@gmail.com";
                    string dia_chi = Vnpayrow["dia_chi"].ToString();
                    string cong_ty = "sisvn";
                    string parammeter = "?" + order_id + "&" + so_tien.ToString() + "&" + ho_ten.Trim() + "&" + dien_thoai + "&" + e_mail + "&" + dia_chi + "&" + cong_ty + "&" + tk_nx;

                    string sql_urlreturn = "select vnpay_api_url from v_dmtknh where tk='" + tk_nx + "'";
                    DataTable tblurl = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql_urlreturn)).Tables[0];
                    if (tblurl.Rows.Count < 1)
                    {
                        MessageBox.Show("Tài khoản: " + tk_nx + " chưa cấu hình Vnpay", "Thông báo");
                        return;
                    }
                    string url = tblurl.Rows[0]["vnpay_api_url"].ToString().Trim() + parammeter;
                    try
                    {
                        System.Diagnostics.Process.Start(url);
                    }
                    catch
                    {
                        MessageBox.Show((StartUp.M_LAN == "E" ? "Can not connect to sercer!" : "Lỗi kết nối tới Vnpay Server."), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                    }
                }
            }
            else
            {
                MessageBox.Show((StartUp.M_LAN == "E" ? "Order is not save!" : "Chưa lưu đơn hàng."), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
            }

        }

        private void BtnCheckvnpay_Click(object sender, RoutedEventArgs e)
        {
            if (Vnpayrow != null)
            {
                if (!string.IsNullOrEmpty(Vnpayrow["order_id"].ToString()))
                {
                    string sql = "Select stt_rec,responsecode from v_ph81 where order_id = '" + Vnpayrow["order_id"].ToString().Trim() + "'";
                    DataTable tbl = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql)).Tables[0];
                    string message = "";
                    if (tbl.Rows.Count > 0)
                    {
                        if (string.IsNullOrEmpty(tbl.Rows[0]["responsecode"].ToString()))
                        {
                            message = StartUp.M_LAN == "E" ? tbl.Rows[0]["stt_rec"].ToString() + "Order Error" : tbl.Rows[0]["stt_rec"].ToString() + "Giao dịch thất bại";
                        }
                        else
                        {
                            if (tbl.Rows[0]["responsecode"].ToString().Trim().Equals("00"))
                            {
                                message = StartUp.M_LAN == "E" ? tbl.Rows[0]["stt_rec"].ToString() + "Order Success" : tbl.Rows[0]["stt_rec"].ToString() + "Giao dịch thành công";
                            }
                        }
                        MessageBox.Show(message, StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
                    }
                }
            }
            else
            {
                MessageBox.Show((StartUp.M_LAN == "E" ? "Chưa chọn đơn hàng!" : "Chưa chọn đơn hàng."), StartUp.SasObj.GetSysvar("M_SAS_VER").ToString().Trim());
            }
        }

        private void BtnHuyKM_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.CTKMTable.Rows.Count <= 0)
            {
                MessageBox.Show("Không có CTKM nào!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            int count = 0;
            foreach (DataRow r in StartUpTrans.DsTrans.Tables[1].Rows)
            {
                count++;
                r["stt_rec0"] = string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)count);
                r["ma_nh_km"] = "";
                r["ma_ctkm"] = "";
                r["pt_km"] = 0;
                r["tien_km"] = 0;
                r["tien_km_nt"] = 0;
                r["ck"] = 0;
                r["ck_nt"] = 0;
            }
            DataRow[] rdel = StartUpTrans.DsTrans.Tables[1].Select("khuyen_mai=1");
            foreach (DataRow r in rdel)
            {
                StartUpTrans.DsTrans.Tables[1].Rows.Remove(r);
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ctkm"] = "";
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km"] = 0;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_km_nt"] = 0;
            ReSum_all();
            foreach (DataRow r in StartUp.CTKMTable.Rows)
            {
                r["ischoose"] = (object)false;
            }
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)false;
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
            ReSum_all();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["co_km"] = (object)true;
            this.tabItemHT.IsSelected = true;
        }

        private void BtnCheckKM_Click(object sender, RoutedEventArgs e)
        {
            AddKhuyenmai();
        }

        private void FormMain_MouseDown(object sender, MouseButtonEventArgs e)
        {

        }
    }
}
