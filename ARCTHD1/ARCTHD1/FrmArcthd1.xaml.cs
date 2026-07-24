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
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace ARCTHD1
{
    public partial class FrmArcthd1 : FormTrans
    {
        public static int iRow = 0;
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
        private bool txtDiaChiFocusable = true;
        public bool isCopy = true;
        public static CodeValueBindingObject IsInEditMode;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private CodeValueBindingObject Voucher_Lan0;
        private CodeValueBindingObject IsCheckedSua_tien;
        private CodeValueBindingObject IsCheckedSua_HT_Thue;
        private CodeValueBindingObject IsCheckedSua_Thue;
        private CodeValueBindingObject Ty_Gia_ValueChange;
        private CodeValueBindingObject M_Ngay_lct;
        private CodeValueBindingObject M_BP_BH;
        private CodeValueBindingObject IsUseCK;
        public DataSet DsVitual;
        private DataSet dsCheckData;
        public static string hinhthuc_tt;

        public FrmArcthd1()
        {
            this.InitializeComponent();
            this.LanguageProvider.Language = StartUpTrans.M_LAN;
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FrmSocthda_Loaded);
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;
            this.C_So_ct = this.txtSo_ct;
            this.C_Ma_nt = this.txtMa_nt;
        }

        private void FrmSocthda_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (this.stt_mau_temlate == 134)
                {
                    this.GrdLayoutThue.Children.Remove((UIElement)this.txtHTTT);
                    this.GrdLayoutHan_ck.Children.Add((UIElement)this.txtHTTT);
                    this.txtHTTT.SetValue(Grid.RowProperty, (object)0);
                    this.txtHTTT.SetValue(Grid.ColumnProperty, (object)1);
                    this.GrdLayoutThue.Children.Remove((UIElement)this.lbhttt);
                    this.GrdLayoutHan_ck.Children.Add((UIElement)this.lbhttt);
                    this.lbhttt.SetValue(Grid.RowProperty, (object)0);
                    this.lbhttt.SetValue(Grid.ColumnProperty, (object)0);
                }
                StartUp.M_AR_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_CK"));
                StartUp.M_AR_TT = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_TT"));
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmArcthd1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmArcthd1.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.IsCheckedSua_HT_Thue = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_HT_Thue");
                this.IsCheckedSua_Thue = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_Thue");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");
                this.M_Ngay_lct = (CodeValueBindingObject)this.FormMain.FindResource((object)"M_Ngay_lct");
                this.M_BP_BH = (CodeValueBindingObject)this.FormMain.FindResource((object)"M_BP_BH");
                this.M_Ngay_lct.Value = StartUp.M_Ngay_lct.Equals("1");
                this.M_BP_BH.Value = StartUp.M_BP_BH.Equals("1");
                this.IsUseCK = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsUseCK");
                this.IsUseCK.Value = StartUp.M_AR_CK == 1;
                this.SetBinding(FormTrans.IsEditModeProperty, (BindingBase)new Binding("Value")
                {
                    Source = (object)FrmArcthd1.IsInEditMode,
                    Mode = BindingMode.TwoWay
                });
                if (FormTrans.SasO.GetOption("M_CDKH13").ToString().Trim() != "1")
                    this.txtSoDuKH.Visibility = this.tblSoDuKH.Visibility = Visibility.Hidden;
                if (StartUp.M_SD_HDDT.Equals("1"))
                    this.tabHDDT.Visibility = Visibility.Visible;
                else
                    this.tabHDDT.Visibility = Visibility.Hidden;
                this.GrdCt.Lan = StartUpTrans.M_LAN;
                this.M_LAN = StartUpTrans.M_LAN;
                FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, (BasicGridView)this.GrdCt, StartUpTrans.Ma_ct, 1);
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
                    this.LoadData();
                    this.LoadDataDu13();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.ListChanged += new ListChangedEventHandler(this.DefaultView_ListChanged);
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    this.IsCheckedSua_tien.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString() == "1";
                    this.Ty_Gia_ValueChange.Value = false;
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                }
                this.TabInfo.SelectedIndex = 0;
                this.ExSetFocusToolBar();
                this.GrdCt.FieldLayouts[0].Settings.RecordSelectorExtent = 350.0;
                this.txtHan_ck.SearchInit();
                this.txtHan_ck_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
                if (StartUp.M_AR_CK == 0)
                {
                    this.tblten_han_ck.Visibility = Visibility.Collapsed;
                    this.GrdLayout00.RowDefinitions[4].Height = new GridLength(114.0);
                    using (IEnumerator<Field> enumerator = this.GrdCt.FieldLayouts[0].Fields.GetEnumerator())
                    {
                        while (enumerator.MoveNext())
                        {
                            Field f = enumerator.Current;
                            if (((IEnumerable<string>)FrmArcthd1.FieldCk).Any<string>((Func<string, bool>)(x => x == f.Name)))
                                f.Visibility = Visibility.Collapsed;
                        }
                    }
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
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;
            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            this.txtStatus.IsEnabled = false;
        }

        public void ExSetFocusToolBar()
        {
            try
            {
                SasVoucherLib.ToolBarButton btnMoi = this.Toolbar.FindName("btnNew") as SasVoucherLib.ToolBarButton;
                Action action = (Action)(() => btnMoi.Focus());
                btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)action);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void DefaultView_ListChanged(object sender, ListChangedEventArgs e)
        {
            this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {
            this.IsVisibilityFieldsXamDataGridByMa_NT(ma_nt);
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
        }

        private void IsVisibilityFieldsXamDataGridByMa_NT(string ma_nt)
        {
            if (ma_nt == StartUpTrans.M_ma_nt0)
            {
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["ck"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["thue"].Visibility = Visibility.Hidden;
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["ck"].Settings.CellMaxWidth = 0.0;
                this.GrdCt.FieldLayouts[0].Fields["thue"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdCt.FieldLayouts[0].Fields["tien2"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["gia2"].Visibility = Visibility.Visible;
                if (StartUp.M_AR_CK == 1)
                    this.GrdCt.FieldLayouts[0].Fields["ck"].Visibility = Visibility.Visible;
                this.GrdCt.FieldLayouts[0].Fields["thue"].Visibility = Visibility.Visible;
                FieldSettings settings1 = this.GrdCt.FieldLayouts[0].Fields["tien2"].Settings;
                FieldLength? width = this.GrdCt.FieldLayouts[0].Fields["tien2"].Width;
                double num1 = width.Value.Value;
                settings1.CellMaxWidth = num1;
                FieldSettings settings2 = this.GrdCt.FieldLayouts[0].Fields["gia2"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["gia2"].Width;
                double num2 = width.Value.Value;
                settings2.CellMaxWidth = num2;
                FieldLength fieldLength;
                if (StartUp.M_AR_CK == 1)
                {
                    FieldSettings settings3 = this.GrdCt.FieldLayouts[0].Fields["ck"].Settings;
                    width = this.GrdCt.FieldLayouts[0].Fields["ck"].Width;
                    fieldLength = width.Value;
                    double num3 = fieldLength.Value;
                    settings3.CellMaxWidth = num3;
                }
                FieldSettings settings4 = this.GrdCt.FieldLayouts[0].Fields["thue"].Settings;
                width = this.GrdCt.FieldLayouts[0].Fields["thue"].Width;
                fieldLength = width.Value;
                double num4 = fieldLength.Value;
                settings4.CellMaxWidth = num4;
            }
            if (StartUp.M_AR_CK == 0)
            {
                this.GrdLayoutTong_NT.RowDefinitions[1].Height = new GridLength(0.0);
                this.GrdLayoutTong_NT.RowDefinitions[2].Height = new GridLength(0.0);
                this.GrdLayoutTong.RowDefinitions[1].Height = new GridLength(0.0);
                this.GrdLayoutTong.RowDefinitions[2].Height = new GridLength(0.0);
                this.GrdLayout00.RowDefinitions[4].Height = new GridLength(115.0);
            }
            if (StartUp.M_AR_TT == 0)
            {
                this.GrdLayoutThue.RowDefinitions[1].Height = new GridLength(0.0);
                this.GrdLayoutThue.RowDefinitions[2].Height = new GridLength(0.0);
            }
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0][nameof(ma_nt)].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0][nameof(ma_nt)].ToString().Equals(StartUpTrans.M_ma_nt0);
            this.ChangeLanguage();
        }

        private void IsVisibilityFieldsXamDataGridBySua_Tien()
        {
            this.IsCheckedSua_tien.Value = this.Chksua_tien.IsChecked.Value;
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void LoadDataDu13()
        {
            this.txtSoDuKH.Value = (object)ArFuncLib.GetSdkh13(StartupBase.SasObj, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString());
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmArcthd1.IsInEditMode.Value)
                return;
            DataRowView dataRowView = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
            DataRow rowResult = this.txtMa_kh.RowResult;
            if (rowResult == null || string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()))
                return;
            dataRowView.BeginEdit();
            if (this.txtMa_kh.IsDataChanged)
            {
                dataRowView["ten_kh_thue"] = (object)rowResult["ten_kh"].ToString().Trim();
                dataRowView["ten_kh"] = (object)rowResult["ten_kh"].ToString().Trim();
                dataRowView["ten_kh2"] = (object)rowResult["ten_kh2"].ToString().Trim();
                if (e != null)
                    dataRowView["ma_so_thue"] = (object)rowResult["ma_so_thue"].ToString().Trim();
                if (string.IsNullOrEmpty(dataRowView["ong_ba"].ToString().Trim()))
                    dataRowView["ong_ba"] = (object)rowResult["doi_tac"].ToString().Trim();
                if (rowResult["tk"].ToString().Trim() != "")
                    dataRowView["ma_nx"] = (object)rowResult["tk"].ToString().Trim();
                this.txtHan_ck.Text = rowResult["ma_thck"].ToString().Trim();
                this.txtHan_ck.SearchInit();
                this.txtHan_ck_PreviewLostFocus((object)null, (KeyboardFocusChangedEventArgs)null);
                if (rowResult["han_tt"] != null && !string.IsNullOrEmpty(rowResult["han_tt"].ToString()))
                    this.txtHan_tt.Text = rowResult["han_tt"].ToString();
                this.LoadDataDu13();
            }
            if (string.IsNullOrEmpty(rowResult["dia_chi"].ToString().Trim()))
            {
                this.txtDiaChiFocusable = true;
            }
            else
            {
                dataRowView["dia_chi"] = (object)rowResult["dia_chi"].ToString().Trim();
                this.txtDiaChiFocusable = false;
            }
            dataRowView["tk_nh"] = (object)rowResult["tk_nh"].ToString().Trim();
            if (StartUp.M_SD_HDDT.Equals("1") && this.txtMa_kh.IsDataChanged)
                dataRowView["sd_hddt_yn"] = (object)SysFunc.suDungHDDT(dataRowView["ma_kh"].ToString());
            dataRowView.EndEdit();
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
                if (this.txtMa_nx.RowResult != null)
                {
                    if (this.M_LAN.ToUpper().Equals("V"))
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx"] = (object)this.txtMa_nx.RowResult["ten_nx"].ToString();
                    else
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_nx2"] = (object)this.txtMa_nx.RowResult["ten_nx2"].ToString();
                }
                if (!this.Chksua_tk_thue.IsChecked.Value)
                    this.txtTk_du_voi_Tk_thue.Text = this.txtMa_nx.Text;
            }
            this.LoadDataDu13();
        }

        private void txtNgay_ct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value == DBNull.Value)
                this.txtNgay_ct.Value = (object)DateTime.Now;
            if (this.txtNgay_ct.IsFocusWithin || FormTrans.currActionTask != ActionTask.Add && FormTrans.currActionTask != ActionTask.Edit && FormTrans.currActionTask != ActionTask.Copy || (!StartUp.M_Ngay_lct.Equals("0") && !(this.txtNgay_lhd.dValue == new DateTime()) || !(this.txtNgay_ct.dValue != new DateTime())))
                return;
            this.txtNgay_lhd.Value = (object)this.txtNgay_ct.dValue.Date;
        }

        private void txtNgay_lhd_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_lhd.Value == DBNull.Value)
                this.txtNgay_lhd.Value = this.txtNgay_ct.Value;

            if (!txtNgay_lhd.IsFocusWithin && FrmArcthd1.IsInEditMode.Value && txtNgay_lhd.IsValueValid && txtNgay_lhd.Value != null && txtNgay_ct.Value.ToString() != txtNgay_lhd.Value.ToString())
            {
                int num = (int)ExMessageBox.Show(225, StartupBase.SasObj, "Ngày lập chứng từ khác với ngày hạch toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!FrmArcthd1.IsInEditMode.Value)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri"] = (object)this.txtMa_qs.RowResult["so_seri"].ToString().Trim();
            if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["so_ct1"].ToString()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct1"] = (object)this.txtMa_qs.RowResult["so_ct1"].ToString();
            if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["so_ct2"].ToString()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct2"] = (object)this.txtMa_qs.RowResult["so_ct2"].ToString();
            if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["transform"].ToString()))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["transform"] = (object)this.txtMa_qs.RowResult["transform"].ToString();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() =>
           {
               if (!FrmArcthd1.IsInEditMode.Value || e.NewFocus.GetType().Equals(typeof(SasVoucherLib.ToolBarButton)) || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                   return;
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
               if (this.CheckValidSoct(StartupBase.SasObj, this.txtMa_qs.Text, this.txtSo_ct.Text, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
               {
                   this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)this.txtSo_ct.Text;
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qstmp"] = (object)this.txtMa_qs.Text;
               }
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["so_lien_hd"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"] = this.txtMa_qs.RowResult["so_lien_hd"];
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien1"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien1"] = (object)this.txtMa_qs.RowResult["ten_lien1"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien2"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien2"] = (object)this.txtMa_qs.RowResult["ten_lien2"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien3"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien3"] = (object)this.txtMa_qs.RowResult["ten_lien3"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien4"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien4"] = (object)this.txtMa_qs.RowResult["ten_lien4"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien5"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien5"] = (object)this.txtMa_qs.RowResult["ten_lien5"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien6"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien6"] = (object)this.txtMa_qs.RowResult["ten_lien6"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien7"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien7"] = (object)this.txtMa_qs.RowResult["ten_lien7"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien8"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien8"] = (object)this.txtMa_qs.RowResult["ten_lien8"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_lien9"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_lien9"] = (object)this.txtMa_qs.RowResult["ten_lien9"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ten_dn_in"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_dn_in"] = (object)this.txtMa_qs.RowResult["ten_dn_in"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["mst_dn_in"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["mst_dn_in"] = (object)this.txtMa_qs.RowResult["mst_dn_in"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["mau_hd"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["mau_hd"] = (object)this.txtMa_qs.RowResult["mau_hd"].ToString();
               if (!string.IsNullOrEmpty(this.txtMa_qs.RowResult["ma_file"].ToString()))
                   StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_file"] = (object)this.txtMa_qs.RowResult["ma_file"].ToString();
           }));
        }

        private void txtMa_nt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.Voucher_Ma_nt0 == null || !this.txtMa_nt.IsDataChanged)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = this.txtMa_nt.RowResult["loai_tg"];
            this.IsVisibilityFieldsXamDataGridByMa_NT(this.txtMa_nt.Text.Trim());
            if (this.txtMa_nt.RowResult != null)
            {
                if (this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    this.txtTy_gia.Value = (object)1;
                else
                    this.txtTy_gia.Value = (object)StartUp.GetRates(this.txtMa_nt.Text.Trim(), Convert.ToDateTime(this.txtNgay_ct.Value).Date);
            }
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
            Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            int result1 = 0;
            int.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sua_tien"].ToString(), out result1);
            if (!(num1 != new Decimal(0)))
                return;
            this.UpdateTotal(StartUpTrans.DsTrans.Tables[1].DefaultView, "tien2", "tien_nt2");
            this.UpdateTotal(StartUpTrans.DsTrans.Tables[1].DefaultView, "gia2", "gia_nt2");
            this.UpdateTotal(StartUpTrans.DsTrans.Tables[1].DefaultView, "thue", "thue_nt");
            this.UpdateTotal(StartUpTrans.DsTrans.Tables[1].DefaultView, "ck", "ck_nt");
            Decimal num2 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
            {
                Decimal.TryParse(dataRowView.Row["thue_suati"].ToString(), out result2);
                Decimal.TryParse(dataRowView.Row["ck_nt"].ToString(), out result3);
                Decimal num3 = SysFunc.Round(result3 * num1, StartUpTrans.M_ROUND);
                dataRowView.Row["ck"] = (object)num3;
            }
            this.UpdateTotalHT();
        }

        private void UpdateTotal(DataView dtview, string columnname, string columnname_nt)
        {
            Decimal num1 = new Decimal(0);
            Decimal result = new Decimal(0);
            Decimal num2 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
            if (dtview.Count <= 0)
                return;
            foreach (DataRowView dataRowView in dtview)
            {
                Decimal.TryParse(dataRowView.Row[columnname_nt].ToString(), out result);
                Decimal num3 = SysFunc.Round(result * num2, StartUpTrans.M_ROUND);
                dataRowView.Row[columnname] = (object)num3;
            }
        }

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
            KeyboardNavigation.SetTabNavigation((DependencyObject)this.GrNT, KeyboardNavigationMode.Continue);
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Return);
        }

        private void V_Truoc()
        {
            if (FrmArcthd1.iRow <= 1)
                return;
            --FrmArcthd1.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Sau()
        {
            if (FrmArcthd1.iRow >= StartUpTrans.DsTrans.Tables[0].Rows.Count - 1)
                return;
            ++FrmArcthd1.iRow;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
            if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("0"))
                this.txtsd_hddt_yn.IsReadOnly = true;
            else
                this.txtsd_hddt_yn.IsReadOnly = false;
        }

        private void V_Dau()
        {
            FrmArcthd1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count < 2 ? 0 : 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void V_Cuoi()
        {
            FrmArcthd1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString() + "'";
            if (StartUp.M_SD_HDDT.Trim().Equals("1"))
                StartUp.refresh(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
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
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.txtsd_hddt_yn.IsReadOnly = false;
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
                    row1["ma_nt"] = StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["ma_nt"];
                    row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["ma_qs"].ToString().Trim());
                }
                row1["ty_giaf"] = !row1["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)StartUp.GetRates(row1["ma_nt"].ToString().Trim(), Convert.ToDateTime(row1["ngay_ct"]).Date) : (object)1;
                row1["sd_hddt_yn"] = (object)0;
                row1["tinh_trang_hddt"] = (object)0;
                row1["mau_hddt"] = (object)DBNull.Value;
                row1["so_seri_hddt"] = (object)DBNull.Value;
                row1["so_ct_hddt"] = (object)DBNull.Value;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["sl_in"] = (object)0;
                row1["t_thue_nt"] = (object)0;
                row1["t_tt_nt"] = (object)0;
                row1["ma_dvcs"] = (object)StartupBase.SasObj.M_ma_dvcs;
                row1["so_seri"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri"];
                row1["so_dh"] = (object)string.Empty;
                row1["so_lo"] = (object)string.Empty;
                row1["ten_kh"] = (object)"";
                row1["dia_chi"] = (object)string.Empty;
                row1["ma_vv"] = (object)string.Empty;
                row1["tk_ck"] = (object)string.Empty;
                row1["t_ck_nt"] = (object)0;
                row1["han_tt"] = (object)0;
                row1["dien_giai"] = (object)string.Empty;
                row1["tien_hg"] = (object)0;
                row1["tien_hg_nt"] = (object)0;
                row1["t_tien2"] = (object)0;
                row1["t_tien_nt2"] = (object)0;
                row1["sua_thue"] = (object)0;
                row1["sua_tkthue"] = (object)0;
                row1["sua_tien"] = (object)0;
                DataRow row2 = StartUpTrans.DsTrans.Tables[1].NewRow();
                row2["stt_rec"] = (object)str;
                row2["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)1);
                row2["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row2["ngay_ct"]= ((this.txtNgay_ct.Value == null) ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
                row2["tien_nt2"] = (object)0;
                row2["tien2"] = (object)0;
                row2["gia_nt2"] = (object)0;
                row2["gia2"] = (object)0;
                row2["so_luong"] = (object)0;
                row2["tl_ck"] = (object)0;
                row2["ck_nt"] = (object)0;
                row2["thue_suati"] = (object)0;
                row2["thue_nt"] = (object)0;
                row2["ck"] = (object)0;
                row2["thue"] = (object)0;
                StartUpTrans.DsTrans.Tables[0].Rows.Add(row1);
                StartUpTrans.DsTrans.Tables[1].Rows.Add(row2);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                this.iRow_old = FrmArcthd1.iRow;
                FrmArcthd1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmArcthd1.IsInEditMode.Value = true;
                this.txtSoDuKH.Text = "";
                this.TabInfo.SelectedIndex = 0;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
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
                int num = (int)ExMessageBox.Show(230, StartupBase.SasObj, "Không có dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim() != "0" && ExMessageBox.Show(1815, StartupBase.SasObj, "Đã phát hành HĐĐT. Có muốn sửa dữ liệu không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Asterisk) == MessageBoxResult.No)
                    return;
                FormTrans.currActionTask = ActionTask.Edit;
                if (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("0"))
                    this.txtsd_hddt_yn.IsReadOnly = true;
                else
                    this.txtsd_hddt_yn.IsReadOnly = false;
                this.DsVitual = new DataSet();
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                this.DsVitual.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                FrmArcthd1.IsInEditMode.Value = true;
                this.IsVisibilityFieldsXamDataGridBySua_Tien();
                this.TabInfo.SelectedIndex = 0;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
            }
        }

        private void V_Copy()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            FormTrans.currActionTask = ActionTask.Copy;
            FrmARCTHD1Copy frmArcthD1Copy = new FrmARCTHD1Copy();
            frmArcthD1Copy.Closed += new EventHandler(this._formcopy_Closed);
            frmArcthD1Copy.ShowDialog();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
        }

        private void _formcopy_Closed(object sender, EventArgs e)
        {
            if (!(sender as FrmARCTHD1Copy).isCopy)
                return;
            string str = DataProvider.NewTrans(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.Ws_Id);
            if (!string.IsNullOrEmpty(str))
            {
                this.DsVitual = StartUpTrans.DsTrans.Copy();
                this.txtsd_hddt_yn.IsReadOnly = false;
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
                DataRow row1 = StartUpTrans.DsTrans.Tables[0].NewRow();
                row1.ItemArray = StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow].ItemArray;
                if (StartUp.M_SD_HDDT.Equals("0"))
                    row1["sd_hddt_yn"] = (object)0;
                row1["tinh_trang_hddt"] = (object)0;
                row1["mau_hddt"] = (object)DBNull.Value;
                row1["so_seri_hddt"] = (object)DBNull.Value;
                row1["so_ct_hddt"] = (object)DBNull.Value;
                row1["stt_rec"] = (object)str;
                row1["stt_rec_pt"] = (object)"";
                row1["so_ct_pt"] = (object)"";
                row1["ma_ct_pt"] = (object)"";
                row1["ngay_ct"] = (object)FrmARCTHD1Copy.ngay_ct;
                row1["status"] = StartUpTrans.DmctInfo["ma_post"];
                row1["ten_post"] = StartUpTrans.tbStatus.Select("ma_post =" + StartUpTrans.DmctInfo["ma_post"].ToString())[0]["ten_post"];
                if (StartUp.M_Ngay_lct.Trim().Equals("0"))
                    row1["ngay_lct"] = (object)FrmARCTHD1Copy.ngay_ct;
                row1["ma_qs"] = (object)this.GetDMQS(this.BindingSasObj, StartUpTrans.Ma_ct, Convert.ToDateTime(row1["ngay_ct"]), StartUpTrans.M_User_Id, row1["ma_qs"].ToString().Trim());
                row1["so_ct"] = !(row1["ma_qs"].ToString().Trim() != "") ? (object)"" : (object)this.GetNewSoct(StartupBase.SasObj, row1["ma_qs"].ToString());
                row1["so_cttmp"] = row1["so_ct"];
                row1["sl_in"] = (object)0;
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
                this.iRow_old = FrmArcthd1.iRow;
                FrmArcthd1.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str + "'";
                FrmArcthd1.IsInEditMode.Value = true;
                this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.txtMa_kh.IsFocus = true));
            }
        }

        private void V_Xoa()
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim()))
                return;
            if (FormTrans.currActionTask == ActionTask.None || FormTrans.currActionTask == ActionTask.View)
                FormTrans.currActionTask = ActionTask.Delete;
            try
            {
                string _stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()) && ExMessageBox.Show(391, StartupBase.SasObj, "Hóa đơn đã được thanh toán, có muốn xóa phiếu thanh toán hay không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                    StartUp.DeletePT(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim());
                StartUpTrans.UpdateTkSd13(1, 0);
                StartUp.DeleteVoucher(_stt_rec, this.txtMa_qs.Text, FormTrans.currActionTask, this.IsNd51);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(FrmArcthd1.iRow);
                if (StartUpTrans.DsTrans.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + _stt_rec + "'"))
                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                }
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                {
                    FrmArcthd1.iRow = FrmArcthd1.iRow > StartUpTrans.DsTrans.Tables[0].Rows.Count - 1 ? FrmArcthd1.iRow - 1 : FrmArcthd1.iRow;
                    StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            FormTrans.currActionTask = ActionTask.None;
        }

        private void V_Huy()
        {
            FrmArcthd1.IsInEditMode.Value = false;
            this.txtsd_hddt_yn.IsReadOnly = true;
            switch (FormTrans.currActionTask)
            {
                case ActionTask.Add:
                case ActionTask.Copy:
                    if (!this.isCopy)
                        this.V_Xoa();
                    if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
                    {
                        FrmArcthd1.iRow = this.iRow_old;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
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
                    StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow].ItemArray = this.DsVitual.Tables[0].Rows[0].ItemArray;
                    StartUpTrans.DsTrans.Tables[1].Merge(this.DsVitual.Tables[1]);
                    this.IsVisibilityFieldsXamDataGrid(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                    break;
            }
            this.TabInfo.SelectedIndex = 0;
            try
            {
                this.txtMa_kh.IsFocus = true;
            }
            catch
            {
            }
            FormTrans.currActionTask = ActionTask.None;
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
                bool flag1 = false;
                if (!this.IsSequenceSave)
                {
                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (this.GrdCt.Records.Count == 0)
                    {
                        int num = (int)ExMessageBox.Show(250, StartupBase.SasObj, "Chưa nhập tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                        return;
                    }
                    if (!flag1)
                        this.txtMa_qs.SearchInit();
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show((int)byte.MaxValue, StartupBase.SasObj, "Chưa có mã khách hàng!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_kh.IsFocus = true;
                        flag1 = true;
                    }
                    else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"].ToString().Trim()))
                    {
                        int num = (int)ExMessageBox.Show(260, StartupBase.SasObj, "Chưa nhập tài khoản nợ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_nx.IsFocus = true;
                        flag1 = true;
                    }
                    else if (this.txtNgay_ct.dValue == new DateTime())
                    {
                        int num = (int)ExMessageBox.Show(265, StartupBase.SasObj, "Chưa vào ngày hạch toán!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                                    goto label_20;
                                }
                            }
                            num1 = 0;
                        }
                        else
                            num1 = 1;
                        label_20:
                        if (num1 == 0)
                        {
                            int num2 = (int)ExMessageBox.Show(1024, StartupBase.SasObj, "Ngày hạch toán không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = true;
                            this.txtNgay_ct.Focus();
                        }
                        else if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["tk_dt"].ToString().Trim()))
                        {
                            int num2 = (int)ExMessageBox.Show(270, StartupBase.SasObj, "Chưa vào tài khoản doanh thu!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            this.TabInfo.SelectedIndex = 0;
                            this.GrdCt.ExecuteCommand(DataPresenterCommands.CellFirstOverall);
                            this.GrdCt.Focus();
                            flag1 = true;
                        }
                    }
                    if (StartUp.M_SD_HDDT.ToString().Equals("1") && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sd_hddt_yn"].ToString() == "1" && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() != string.Empty)
                    {
                        SqlCommand sqlcmd = new SqlCommand();
                        sqlcmd.CommandText = "SELECT * FROM khhddt WHERE ma_kh='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString() + "' AND status in (1,2)";
                        if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count == 0)
                        {
                            int num = (int)ExMessageBox.Show(999, StartupBase.SasObj, "Khách hàng [" + (StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kh.RowResult["ten_kh"].ToString().Trim() : this.txtMa_kh.RowResult["ten_kh2"].ToString().Trim()) + "] chưa cập nhật khách hàng hóa đơn điện tử!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            flag1 = true;
                            this.txtMa_kh.IsFocus = true;
                        }
                    }
                    if (!flag1)
                    {
                        if (!this.IsNd51)
                        {
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(1240, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = true;
                                this.txtMa_qs.IsFocus = true;
                            }
                            if (string.IsNullOrEmpty(this.txtSo_ct.Text.Trim()))
                            {
                                int num = (int)ExMessageBox.Show(275, StartupBase.SasObj, "Chưa vào số hóa đơn!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtSo_ct.Text = this.txtSo_ct.Text.Trim();
                                this.txtSo_ct.Focus();
                                flag1 = true;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(1240, StartupBase.SasObj, "Chưa nhập quyển chứng từ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = true;
                                this.txtMa_qs.IsFocus = true;
                            }
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString()) && !flag1)
                            {
                                int num = (int)ExMessageBox.Show(295, StartupBase.SasObj, "Chưa vào số hóa đơn!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag1 = true;
                                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                            }
                            if (!flag1 && !this.CheckSo_ct(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["transform"].ToString(), this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["so_ct1"], new Decimal(0)), this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["so_ct2"], new Decimal(0)), this.txtSo_ct.Text.Trim()))
                            {
                                int num = (int)ExMessageBox.Show(300, StartupBase.SasObj, "Số hóa đơn không thuộc ký hiệu hiện hành!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                                flag1 = true;
                            }
                        }
                    }
                }
                if (!flag1)
                {
                    this.UpdateTotalHT();
                    if (!this.IsSequenceSave)
                    {
                        if (this.txtMa_kh.RowResult == null)
                            this.txtMa_kh.SearchInit();
                        if (string.IsNullOrEmpty(this.txtMa_kh.RowResult["dia_chi"].ToString().Trim()) || string.IsNullOrEmpty(this.txtMa_kh.RowResult["ma_so_thue"].ToString().Trim()))
                        {
                            KhInfoFrm khInfoFrm = new KhInfoFrm();
                            khInfoFrm.ShowDialog();
                            if (!khInfoFrm.IsAllowSave)
                                return;
                        }
                        if (StartUpTrans.DsTrans.Tables[0].DefaultView[0][StartUpTrans.M_LAN.Equals("V") ? "ten_kh" : "ten_kh2"].ToString().Trim() == "")
                        {
                            this.txtMa_kh.SearchInit();
                            this.txtMa_kh.IsDataChanged = true;
                            this.txtMa_kh_PreviewLostFocus((object)this.txtMa_kh, (KeyboardFocusChangedEventArgs)null);
                        }
                        if (!SysFunc.CheckSumMaSoThue(this.txtma_so_thue.Text.Trim()) && !string.IsNullOrEmpty(this.txtma_so_thue.Text.Trim()))
                        {
                            switch (StartUpTrans.M_MST_CHECK.Trim())
                            {
                                case "1":
                                    int num1 = (int)ExMessageBox.Show(335, StartupBase.SasObj, "Mã số thuế không hợp lệ!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    break;
                                case "2":
                                    int num2 = (int)ExMessageBox.Show(340, StartupBase.SasObj, "Mã số thuế không hợp lệ, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    flag1 = true;
                                    break;
                            }
                        }
                        if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                        {
                            int index = 0;
                            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                            {
                                if (string.IsNullOrEmpty(dataRowView.Row["tk_dt"].ToString().Trim()))
                                {
                                    StartUpTrans.DsTrans.Tables[1].Rows.Remove(dataRowView.Row);
                                    StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(dataRowView.Row["tk_ck"].ToString().Trim()) && Convert.ToDecimal(dataRowView.Row["tl_ck"].ToString()) > new Decimal(0))
                                    {
                                        int num3 = (int)ExMessageBox.Show(345, StartupBase.SasObj, "Chưa vào tk ck, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        flag1 = true;
                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[index] as DataRecord).Cells["tk_ck"];
                                        this.GrdCt.Focus();
                                        this.GrdCt.ExecuteCommand(DataPresenterCommands.StartEditMode);
                                    }
                                    ++index;
                                }
                            }
                        }
                        object obj1 = StartUpTrans.DsTrans.Tables[1].Compute("sum(thue_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj2 = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt2)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj3 = StartUpTrans.DsTrans.Tables[1].Compute("sum(thue)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        object obj4 = StartUpTrans.DsTrans.Tables[1].Compute("sum(tien2)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
                        Decimal num4 = Convert.ToDecimal(obj1.Equals((object)DBNull.Value) ? (object)0 : obj1);
                        Decimal num5 = Convert.ToDecimal(obj2.Equals((object)DBNull.Value) ? (object)0 : obj2);
                        Decimal num6 = Convert.ToDecimal(obj3.Equals((object)DBNull.Value) ? (object)0 : obj3);
                        Decimal num7 = Convert.ToDecimal(obj4.Equals((object)DBNull.Value) ? (object)0 : obj4);
                        if (FormTrans.currActionTask == ActionTask.Copy && (num5 != num4 || num6 != num7) || !(num5 != num4) && !(num6 != num7))
                            ;
                        if (!flag1)
                        {
                            Decimal nValue = this.txtTy_gia.nValue;
                            if (!this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) && this.GrdCt.Records.Count > 0 && nValue != new Decimal(0) && !this.Chksua_tien.IsChecked.Value)
                            {
                                Decimal result1 = new Decimal(0);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b =>
                               {
                                   int num;
                                   if (b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())
                                   {
                                       Decimal? nullable = b.Field<Decimal?>("tien_nt2");
                                       if ((!(nullable.GetValueOrDefault() == new Decimal(0)) ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                                       {
                                           nullable = b.Field<Decimal?>("tien2");
                                           num = nullable.GetValueOrDefault() != new Decimal(0) ? 1 : (!nullable.HasValue ? 1 : 0);
                                           goto label_4;
                                       }
                                   }
                                   num = 0;
                               label_4:
                                   return num != 0;
                               })).Count<DataRow>().ToString(), out result1);
                                if (result1 == new Decimal(0))
                                {
                                    Decimal num3 = SysFunc.Round(nValue * num5, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    this.txtt_tien.Value = (object)num3;
                                    this.txttien_sau_ck.Value = (object)(num3 - this.txtt_ck.nValue);
                                    Decimal result2 = new Decimal(0);
                                    Decimal? nullable = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien2")));
                                    if (nullable.HasValue)
                                        Decimal.TryParse(nullable.ToString(), out result2);
                                    (this.GrdCt.Records[0] as DataRecord).Cells["tien2"].Value = (object)(Convert.ToDecimal((this.GrdCt.Records[0] as DataRecord).Cells["tien2"].Value) + (num3 - result2));
                                    this.UpdateTotalHT();
                                }
                                else if (result1 < (Decimal)this.GrdCt.Records.Count)
                                {
                                    Decimal num3 = SysFunc.Round(nValue * num5, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    this.txtt_tien.Value = (object)num3;
                                    this.txttien_sau_ck.Value = (object)(num3 + this.txtt_ck.nValue);
                                    Decimal result2 = new Decimal(0);
                                    Decimal? nullable = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien2")));
                                    if (nullable.HasValue)
                                        Decimal.TryParse(nullable.ToString(), out result2);
                                    for (int index = 0; index < this.GrdCt.Records.Count; ++index)
                                    {
                                        DataRecord record = this.GrdCt.Records[index] as DataRecord;
                                        Decimal result3 = new Decimal(0);
                                        Decimal result4 = new Decimal(0);
                                        Decimal.TryParse(record.Cells["tien_nt2"].Value.ToString(), out result3);
                                        Decimal.TryParse(record.Cells["tien2"].Value.ToString(), out result4);
                                        if (!(result3 == new Decimal(0)) || !(result4 != new Decimal(0)))
                                        {
                                            record.Cells["tien2"].Value = (object)(result4 + (num3 - result2));
                                            break;
                                        }
                                    }
                                    this.UpdateTotalHT();
                                }
                            }
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString()))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"] = StartUpTrans.DmctInfo["ma_gd"];
                            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = (object)StartupBase.SasObj.GetOption("M_MA_DVCS").ToString();
                        }
                    }
                    if (!flag1)
                    {
                        DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
                        LocalTable1.Rows.Add(StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row.ItemArray);
                        if (!this.IsSequenceSave)
                            LocalTable1.Rows[0]["status"] = (object)0;
                        if (LocalTable1.Rows[0]["ten_kh_thue"].ToString().Trim() != "")
                            LocalTable1.Rows[0]["ten_kh"] = LocalTable1.Rows[0]["ten_kh_thue"];
                        DataProvider.UpdateDataTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", LocalTable1, "stt_rec;row_id");
                        DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        {
                            if (!this.IsSequenceSave)
                            {
                                dataRowView["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                                dataRowView["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                                dataRowView["ma_ct"] = (object)StartUpTrans.Ma_ct;
                                dataRowView["ten_vt"] = dataRowView["dien_giaii"];
                                if (this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    dataRowView["tien2"] = dataRowView["tien_nt2"];
                                    dataRowView["ck"] = dataRowView["ck_nt"];
                                    dataRowView["thue"] = dataRowView["thue_nt"];
                                }
                            }
                            LocalTable2.Rows.Add(dataRowView.Row.ItemArray);
                        }
                        if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            int num = (int)ExMessageBox.Show(385, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            return;
                        }
                        if (!this.IsSequenceSave)
                        {
                            if (!flag1)
                            {
                                this.dsCheckData = StartUp.CheckData(FormTrans.currActionTask == ActionTask.Edit ? 0 : 1);
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
                                                        if (ExMessageBox.Show(390, StartupBase.SasObj, "Số chứng từ đã tồn tại, số cuối cùng là: [" + this.GetLastSoct(StartupBase.SasObj, this.txtMa_qs.Text).Trim() + "]. Có lưu chứng từ này không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                        {
                                                            this.txtSo_ct.SelectAll();
                                                            this.txtSo_ct.Focus();
                                                            flag1 = true;
                                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)"";
                                                            break;
                                                        }
                                                        break;
                                                    }
                                                    if (StartUpTrans.M_trung_so.Equals("2"))
                                                    {
                                                        int num = (int)ExMessageBox.Show(395, StartupBase.SasObj, "Số chứng từ đã tồn tại!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        this.txtSo_ct.SelectAll();
                                                        this.txtSo_ct.Focus();
                                                        flag1 = true;
                                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_cttmp"] = (object)"";
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
                                                                int num1 = (int)ExMessageBox.Show(400, FormTrans.SasO, "Ngày bắt đầu sử dụng ký hiệu hóa đơn không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtNgay_ct.Focus();
                                                                flag1 = true;
                                                                break;
                                                            case 2:
                                                                int num2 = (int)ExMessageBox.Show(405, FormTrans.SasO, "Quyền sử dụng ký hiệu hóa đơn không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtMa_qs.IsFocus = true;
                                                                flag1 = true;
                                                                break;
                                                            case 3:
                                                                string lower = this.dsCheckData.Tables[4].Rows[0]["ten_tthd"].ToString().Trim().ToLower();
                                                                int num3 = (int)ExMessageBox.Show(410, FormTrans.SasO, "Số hóa đơn của ký hiệu [" + this.txtMa_qs.Text.Trim() + "] đã [" + lower + "]!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                                this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                                                                flag1 = true;
                                                                break;
                                                        }
                                                        break;
                                                    }
                                                    break;
                                                case "PH03":
                                                    int num4 = (int)ExMessageBox.Show(415, StartupBase.SasObj, "Số hóa đơn không liên tục!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtSo_ct.Text = this.GetNewSoct(StartupBase.SasObj, this.txtMa_qs.Text);
                                                    break;
                                                case "PH04":
                                                    int num5 = (int)ExMessageBox.Show(420, StartupBase.SasObj, "Tk nợ là tài khoản tổng hợp không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtMa_nx.IsFocus = true;
                                                    break;
                                                case "PH05":
                                                    int num6 = (int)ExMessageBox.Show(425, StartupBase.SasObj, "Tk đ.ứng với tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.txtTk_du_voi_Tk_thue.IsFocus = true;
                                                    break;
                                                case "CT01":
                                                    int int16_1 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num7 = (int)ExMessageBox.Show(430, StartupBase.SasObj, "Tk doanh thu là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_1] as DataRecord).Cells["tk_dt"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GrdCt.Focus()));
                                                    break;
                                                case "CT02":
                                                    if (StartUp.M_AR_CK == 1)
                                                    {
                                                        int int16_2 = (int)Convert.ToInt16(dataRowView[1]);
                                                        int num1 = (int)ExMessageBox.Show(435, StartupBase.SasObj, "Tk c.khấu là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        flag1 = true;
                                                        this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_2] as DataRecord).Cells["tk_ck"];
                                                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GrdCt.Focus()));
                                                        break;
                                                    }
                                                    break;
                                                case "CT03":
                                                    int int16_3 = (int)Convert.ToInt16(dataRowView[1]);
                                                    int num8 = (int)ExMessageBox.Show(440, StartupBase.SasObj, "Tk thuế là tk tổng hợp, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                    flag1 = true;
                                                    this.GrdCt.ActiveCell = (this.GrdCt.Records[int16_3] as DataRecord).Cells["tk_thue_i"];
                                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GrdCt.Focus()));
                                                    break;
                                            }
                                            this.dsCheckData.Tables[0].Rows.Remove(dataRowView.Row);
                                        }
                                        else
                                            break;
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
                            string format1 = "EXEC  {0} '" + this.txtMa_qs.Text.Trim() + "', '" + this.txtSo_ct.Text.Trim() + "'";
                            this.BindingSasObj.ExcuteNonQuery(new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 2 ? string.Format(format1, (object)"SetSoct") : string.Format(format1, (object)StartUpTrans.Process_Store[2])));
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
                                string str1 = this.BindingSasObj.GetOption("M_TK_TK_VT").ToString();
                                char[] chArray = new char[1] { ',' };
                                foreach (string str2 in str1.Split(chArray))
                                {
                                    if (!string.IsNullOrEmpty(str2.ToString().Trim()) && this.txtMa_nx.Text.Trim().StartsWith(str2.ToString().Trim()))
                                        flag2 = true;
                                }
                                if (flag2)
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
                                        if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()))
                                        {
                                            DataTable dataTable = (DataTable)null;
                                            if (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString()))
                                            {
                                                SqlCommand sqlcmd = new SqlCommand();
                                                sqlcmd.CommandText = string.Format("SELECT stt_rec,ma_ct,ma_gd,ma_qs,so_ct,ma_nt,ong_ba,dien_giai FROM {0} WHERE stt_rec LIKE '{1}'", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Equals("PT1") ? (object)"ph41" : (object)"ph51", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString());
                                                dataTable = this.BindingSasObj.ExcuteReader(sqlcmd).Tables[0];
                                                if (dataTable.Rows.Count == 1)
                                                    StartUp.DeletePT(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim());
                                            }
                                            int num;
                                            if (this.txtT_tt.nValue > Convert.ToDecimal(this.BindingSasObj.GetSysvar("M_MUC_TIEN_PT1")))
                                            {
                                                num = 1;
                                            }
                                            else
                                            {
                                                SqlCommand sqlcmd = new SqlCommand();
                                                sqlcmd.CommandText = "SELECT COUNT(1) FROM dmtknh WHERE tk LIKE @tk";
                                                sqlcmd.Parameters.Add(new SqlParameter("@tk", SqlDbType.VarChar)).Value = (object)this.txtMa_nx.Text.Trim();
                                                num = (int)this.BindingSasObj.ExcuteScalar(sqlcmd);
                                            }
                                            DataRow row = dt.NewRow();
                                            dt.Rows.Add(row);
                                            FrmTaoPT frmTaoPt = new FrmTaoPT();
                                            frmTaoPt.tbInfoPT = dataTable;
                                            frmTaoPt.DataContext = (object)dt.DefaultView;
                                            frmTaoPt.txtMa_qs_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"].ToString();
                                            frmTaoPt.txtso_ct_pt.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"].ToString();
                                            frmTaoPt.txtnguoi_nop.Text = this.txtOng_ba.Text;
                                            frmTaoPt.kind = num > 0 ? 2 : 1;
                                            frmTaoPt.Ma_nt_ht = this.txtMa_nt.Text;
                                            frmTaoPt.so_hd = this.txtSo_ct.Text.Trim();
                                            frmTaoPt.ngay_hd = this.txtNgay_ct.dValue.ToShortDateString();
                                            frmTaoPt.filterma_qs = this.txtMa_qs.Filter;
                                            frmTaoPt.ShowDialog();
                                            if (!frmTaoPt.isOk)
                                            {
                                                _createPT1 = false;
                                            }
                                            else
                                            {
                                                dt = dt.Copy();
                                                newstt_recPt1 = !frmTaoPt.txtKind.Text.Equals("1") ? DataProvider.NewTrans(StartupBase.SasObj, "BC1", StartUpTrans.Ws_Id) : DataProvider.NewTrans(StartupBase.SasObj, "PT1", StartUpTrans.Ws_Id);
                                                dt.Rows[0]["ma_ct"] = frmTaoPt.txtKind.Text.Equals("1") ? (object)"PT1" : (object)"BC1";
                                                dt.Rows[0]["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                                                dt.Rows[0]["stt_recPT"] = (object)newstt_recPt1;
                                                dt.Rows[0]["ma_qs"] = (object)frmTaoPt.txtMa_qs_pt.Text;
                                                dt.Rows[0]["so_ct"] = (object)frmTaoPt.txtso_ct_pt.Text.PadLeft(frmTaoPt.txtso_ct_pt.MaxLength, ' ');
                                                dt.Rows[0]["ma_nt"] = (object)frmTaoPt.txtMa_nt.Text;
                                                dt.Rows[0]["ty_gia"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? (object)1 : this.txtTy_gia.Rate;
                                                dt.Rows[0]["ty_giaf"] = frmTaoPt.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? (object)1 : this.txtTy_gia.RateF;
                                                dt.Rows[0]["nguoinop"] = (object)frmTaoPt.txtnguoi_nop.Text;
                                                dt.Rows[0]["lydonop"] = (object)frmTaoPt.txtlydo_nop.Text;
                                                dt.Rows[0]["ma_gd"] = (object)frmTaoPt.txtMa_gd.Text;
                                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs_pt"] = (object)frmTaoPt.txtMa_qs_pt.Text;
                                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = (object)frmTaoPt.txtso_ct_pt.Text;
                                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = frmTaoPt.txtKind.Text.Equals("1") ? (object)"PT1" : (object)"BC1";
                                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = (object)newstt_recPt1;
                                                string format2 = "EXEC  {0} '" + frmTaoPt.txtMa_qs_pt.Text.Trim() + "', '" + frmTaoPt.txtso_ct_pt.Text.Trim() + "'";
                                                this.BindingSasObj.ExcuteNonQuery(new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 2 ? string.Format(format2, (object)"SetSoct") : string.Format(format2, (object)StartUpTrans.Process_Store[2])));
                                            }
                                        }
                                        else
                                            _createPT1 = false;
                                    }
                                }
                            }
                            string _stt_rec1 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString();
                            new Thread((ThreadStart)(() =>
                           {
                               this.Post(_createPT1 ? 1 : 0);
                               if (_createPT1)
                                   this.CreatePT1(dt);
                               this.Dispatcher.BeginInvoke(DispatcherPriority.Normal, (Delegate) new Action(() =>
                 {
                     if (this.IsSequenceSave)
                         return;
                     if (StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Equals(_stt_rec1))
                         this.LoadDataDu13();
                     if (_createPT1 && !string.IsNullOrEmpty(newstt_recPt1))
                     {
                         DataRow[] dataRowArray = StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + _stt_rec1 + "'");
                         if (dataRowArray.Length == 1)
                         {
                             dataRowArray[0]["stt_rec_pt"] = (object)newstt_recPt1;
                             dataRowArray[0]["so_ct_pt"] = (object)dt.Rows[0]["so_ct"].ToString().Trim();
                             dataRowArray[0]["ma_ct_pt"] = dt.Rows[0]["ma_ct"];
                         }
                     }
                 }));
                           })).Start();
                            if (!this.IsSequenceSave)
                            {
                                int pos = this.GetiRow(StartUpTrans.DsTrans.Tables[0], StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString());
                                if (FrmArcthd1.iRow != pos)
                                {
                                    DataRow row1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row;
                                    DataRow row2 = StartUpTrans.DsTrans.Tables[0].NewRow();
                                    row2.ItemArray = row1.ItemArray;
                                    if (FrmArcthd1.iRow > pos)
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos);
                                    else
                                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row2, pos + 1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    StartUpTrans.DsTrans.Tables[0].Rows.Remove(row1);
                                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                                    FrmArcthd1.iRow = pos;
                                }
                                FrmArcthd1.IsInEditMode.Value = false;
                                FormTrans.currActionTask = ActionTask.View;
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
            this.isCopy = false;
        }

        private void CreatePT1(DataTable dt)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[ARCTHD1-CREATEPT1] @Stt_rec, @Stt_recPT, @ma_qs, @so_ct, @ma_nt, @ty_gia, @ty_giaf, @nguoinop, @lydonop, @ma_gd, @ma_ct");
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

        private void V_Xem()
        {
            FormTrans.currActionTask = ActionTask.View;
            string str1 = "";
            string str2 = "";
            if (StartUpTrans.M_LAN.Equals("V"))
            {
                if (StartUpTrans.CommandInfo["Vbrowse2"] != null)
                {
                    string str3 = StartUpTrans.CommandInfo["Vbrowse2"].ToString();
                    str1 = str3.Split('|')[0];
                    str2 = str3.Split('|')[1];
                }
            }
            else if (StartUpTrans.CommandInfo["Ebrowse2"] != null)
            {
                string str3 = StartUpTrans.CommandInfo["Ebrowse2"].ToString();
                str1 = str3.Split('|')[0];
                str2 = str3.Split('|')[1];
            }
            if (StartUp.M_AR_CK == 0)
            {
                str1 = StartUp.EditFields(str1);
                str2 = StartUp.EditFields(str2);
            }
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, str1, str2, "stt_rec");
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.TongCongLabel = "Tổng cộng:";
            formView.frmBrw.LanguageID = "ARCTHD1_6";
            formView.ShowDialog();
            if (formView.DataGrid.ActiveRecord == null)
                return;
            int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
            if (index >= 0)
            {
                string str3 = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                FrmArcthd1.iRow = index + 1;
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
            }
        }

        private void V_Tim()
        {
            try
            {
                FormTrans.currActionTask = ActionTask.View;
                FrmTim3 frmTim3 = new FrmTim3(StartupBase.SasObj, StartUpTrans.filterId, StartUpTrans.filterView);
                frmTim3.txtMa_qs.Filter = "ma_cts like '%HD1%' and status =1 AND (" + this.StrFilterQS + ")";
                SysFunc.LoadIcon((Window)frmTim3);
                frmTim3.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void V_In()
        {
            FrmArcthd1.hinhthuc_tt = this.txtHTTT.Text.Trim();
            new FrmIn(this.IsNd51).ShowDialog();
        }

        private void NewRowCt()
        {
            DataRow dataRow = StartUpTrans.DsTrans.Tables[1].NewRow();
            dataRow["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            int result = 0;
            if (this.GrdCt.Records.Count > 0)
            {
                string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                if (str != null)
                    int.TryParse(str.ToString(), out result);
            }
            int num = result + 1;
            dataRow["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)num);
            dataRow["ma_ct"] = (object)StartUpTrans.Ma_ct;
            dataRow["ngay_ct"] = (object)(this.txtNgay_ct.Value == null ? DateTime.Now.Date : this.txtNgay_ct.dValue.Date);
            dataRow["tien_nt2"] = (object)0;
            dataRow["tien2"] = (object)0;
            dataRow["gia_nt2"] = (object)0;
            dataRow["gia2"] = (object)0;
            dataRow["so_luong"] = (object)0;
            dataRow["tl_ck"] = (object)0;
            dataRow["ck_nt"] = (object)0;
            dataRow["thue_suati"] = (object)0;
            dataRow["thue_nt"] = (object)0;
            dataRow["ck"] = (object)0;
            dataRow["thue"] = (object)0;
            int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
            dataRow["dien_giaii"] = count <= 0 ? StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row["dien_giai"] : StartUpTrans.DsTrans.Tables[1].DefaultView[count - 1].Row["dien_giaii"];
            FreeCodeFieldLib.CarryFreeCodeFields(StartupBase.SasObj, StartUpTrans.Ma_ct, StartUpTrans.DsTrans.Tables[1].DefaultView, dataRow, 1);
            StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow);
        }

        private void GrdCt_RecordDelete(object sender, RecordsDeletedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate) new Action(() =>
           {
               if (this.stt_mau_temlate == 137)
               {
                   if (!this.IsCheckedSua_HT_Thue.Value || this.txtTk_du_voi_Tk_thue.IsReadOnly)
                       this.txtHTTT.Focus();
                   else
                       this.txtTk_du_voi_Tk_thue.IsFocus = true;
               }
               else
                   this.txtHan_ck.IsFocus = true;
           }), DispatcherPriority.Background);
        }

        private void GrdCt_PreviewEditModeEnded(object sender, EditModeEndedEventArgs e)
        {
            try
            {
                if (this.IsEditMode && this.GrdCt.ActiveCell != null && StartUpTrans.DsTrans.Tables[1].DefaultView.Count > this.GrdCt.ActiveRecord.Index && StartUpTrans.DsTrans.Tables[1].GetChanges(DataRowState.Deleted) == null)
                {
                    Decimal num1 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
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
                    switch (e.Cell.Field.Name)
                    {
                        case "tk_dt":
                            AutoCompleteTextBox autoCompleteControl1 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl1.RowResult != null && !autoCompleteControl1.Text.Trim().Equals(""))
                            {
                                e.Cell.Record.Cells["ten_tk_dt"].Value = autoCompleteControl1.RowResult["ten_tk"];
                                e.Cell.Record.Cells["ten_tk_dt2"].Value = autoCompleteControl1.RowResult["ten_tk2"];
                            }
                            if (autoCompleteControl1.Text.Trim().Equals(""))
                            {
                                e.Cell.Record.Cells["ten_tk_dt"].Value = (object)"";
                                break;
                            }
                            break;
                        case "so_luong":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                num2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                num5 = new Decimal(0);
                                Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                                Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal nValue = this.txtTy_gia.nValue;
                                if (num14 == new Decimal(0))
                                {
                                    e.Cell.Record.Cells["gia_nt2"].Value = (object)0;
                                    e.Cell.Record.Cells["gia2"].Value = (object)0;
                                    break;
                                }
                                if (num13 * num14 != new Decimal(0))
                                {
                                    Decimal num15 = !this.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? SysFunc.Round(num13 * num14, StartUpTrans.M_ROUND_NT) : SysFunc.Round(num13 * num14, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien_nt2"].Value = (object)num15;
                                    if (!this.Chksua_tien.IsChecked.Value)
                                        e.Cell.Record.Cells["tien2"].Value = (object)(num15 * nValue == new Decimal(0) ? num15 : SysFunc.Round(num15 * nValue, StartUpTrans.M_ROUND));
                                    this.GrdCt_PreviewEditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien_nt2"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien_nt2"]).Editor, true));
                                }
                                break;
                            }
                            break;
                        case "gia_nt2":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                num2 = new Decimal(0);
                                num3 = new Decimal(0);
                                num4 = new Decimal(0);
                                num5 = new Decimal(0);
                                Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                                Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                Decimal nValue = this.txtTy_gia.nValue;
                                if (num13 * num14 != new Decimal(0))
                                {
                                    Decimal num15 = !this.txtMa_nt.Text.Equals(StartupBase.M_MA_NT0) ? SysFunc.Round(num13 * num14, StartUpTrans.M_ROUND_NT) : SysFunc.Round(num13 * num14, StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["tien_nt2"].Value = (object)num15;
                                    if (!this.Chksua_tien.IsChecked.Value)
                                    {
                                        e.Cell.Record.Cells["gia2"].Value = (object)(num15 * nValue == new Decimal(0) ? num15 : SysFunc.Round(num13 * nValue, StartUpTrans.M_ROUND_GIA));
                                        e.Cell.Record.Cells["tien2"].Value = (object)(num15 * nValue == new Decimal(0) ? num15 : SysFunc.Round(num15 * nValue, StartUpTrans.M_ROUND));
                                    }
                                }
                                if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    e.Cell.Record.Cells["gia2"].Value = e.Cell.Record.Cells["gia_nt2"].Value;
                                this.GrdCt_PreviewEditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien_nt2"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien_nt2"]).Editor, true));
                                break;
                            }
                            break;
                        case "gia2":
                            if (e.Cell.IsDataChanged && (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim())))
                            {
                                Decimal num13 = new Decimal(0);
                                num4 = new Decimal(0);
                                Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["gia2"].Value, new Decimal(0));
                                Decimal num15 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                if (num15 * num14 != new Decimal(0))
                                    e.Cell.Record.Cells["tien2"].Value = (object)SysFunc.Round(num15 * num14, StartUpTrans.M_ROUND);
                                this.GrdCt_PreviewEditModeEnded((object)this.GrdCt, new EditModeEndedEventArgs(e.Cell.Record.Cells["tien2"], CellValuePresenter.FromCell(e.Cell.Record.Cells["tien2"]).Editor, true));
                                break;
                            }
                            break;
                        case "tien_nt2":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    num6 = new Decimal(0);
                                    num7 = new Decimal(0);
                                    num8 = new Decimal(0);
                                    num9 = new Decimal(0);
                                    num10 = new Decimal(0);
                                    num11 = new Decimal(0);
                                    num12 = new Decimal(0);
                                    Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["tl_ck"].Value, new Decimal(0));
                                    Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                    Decimal nValue2 = this.txtTy_gia.nValue;
                                    Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["thue_suati"].Value, new Decimal(0));
                                    Decimal num15 = SysFunc.Round(nValue2 * nValue1, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    if (num15 != new Decimal(0))
                                        e.Cell.Record.Cells["tien2"].Value = (object)num15;

                                    Decimal num73 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                                    Decimal num74 = this.ParseDecimal(e.Cell.Record.Cells["so_luong"].Value, new Decimal(0));
                                    Decimal num75 = this.ParseDecimal(e.Cell.Record.Cells["gia_nt2"].Value, new Decimal(0));
                                    if (num75 == new Decimal(0))
                                    {
                                        if (num74 != new Decimal(0))
                                            num75 = SysFunc.Round(num73 / num74, StartUpTrans.M_ROUND_GIA);
                                        e.Cell.Record.Cells["gia_nt2"].Value = num75;
                                        Decimal num76 = SysFunc.Round(nValue2 * num75, (int)Convert.ToInt16(StartUpTrans.M_ROUND_GIA_NT));
                                        if (num76 != new Decimal(0))
                                            e.Cell.Record.Cells["gia2"].Value = (object)num76;
                                        else
                                            e.Cell.Record.Cells["gia2"].Value = num75;
                                    }

                                    Decimal num16 = SysFunc.Round(num13 * num15 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    if (num16 != new Decimal(0))
                                        e.Cell.Record.Cells["ck"].Value = (object)num16;
                                    Decimal tien_hang1 = num15;
                                    Decimal ck1 = num16;
                                    Decimal thue_suat1 = num14;
                                    bool? isChecked = this.Chkthue_ck0.IsChecked;
                                    int num17 = isChecked.Value ? 1 : 0;
                                    Decimal num18 = SysFunc.Round(this.TinhThue(tien_hang1, ck1, thue_suat1, num17 != 0), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["thue"].Value = (object)num18;
                                    Decimal num19 = SysFunc.Round(nValue1 * num13 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    if (num19 != new Decimal(0))
                                        e.Cell.Record.Cells["ck_nt"].Value = (object)num19;
                                    Decimal tien_hang2 = nValue1;
                                    Decimal ck2 = num19;
                                    Decimal thue_suat2 = num14;
                                    isChecked = this.Chkthue_ck0.IsChecked;
                                    int num20 = isChecked.Value ? 1 : 0;
                                    Decimal num21 = this.TinhThue(tien_hang2, ck2, thue_suat2, num20 != 0);
                                    Decimal num22 = !(this.txtMa_nt.Text != StartUpTrans.M_ma_nt0) ? SysFunc.Round(num21, StartUpTrans.M_ROUND) : SysFunc.Round(num21, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)num22;
                                }
                                if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["tien2"].Value = e.Cell.Record.Cells["tien_nt2"].Value;
                                    e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                }
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tien2":
                            if (e.Cell.IsDataChanged)
                            {
                                if (!this.IsCheckedSua_tien.Value && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    Decimal tien_hang = this.ParseDecimal(e.Cell.Record.Cells["tien2"].Value, new Decimal(0));
                                    Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["tl_ck"].Value, new Decimal(0)) / new Decimal(100);
                                    Decimal thue_suat = this.ParseDecimal(e.Cell.Record.Cells["thue_suati"].Value, new Decimal(0));
                                    Decimal ck = SysFunc.Round(tien_hang * num13, StartUpTrans.M_ROUND);
                                    Decimal num14 = SysFunc.Round(this.TinhThue(tien_hang, ck, thue_suat, this.Chkthue_ck0.IsChecked.Value), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["ck"].Value = (object)ck;
                                    e.Cell.Record.Cells["thue"].Value = (object)num14;
                                }
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tl_ck":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(e.Editor.Text.Trim()))
                                {
                                    num6 = new Decimal(0);
                                    num7 = new Decimal(0);
                                    num8 = new Decimal(0);
                                    num9 = new Decimal(0);
                                    num10 = new Decimal(0);
                                    num11 = new Decimal(0);
                                    num12 = new Decimal(0);
                                    Decimal nValue1 = (e.Editor as NumericTextBox).nValue;
                                    Decimal tien_hang1 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                                    Decimal thue_suat = this.ParseDecimal(e.Cell.Record.Cells["thue_suati"].Value, new Decimal(0));
                                    Decimal nValue2 = this.txtTy_gia.nValue;
                                    Decimal tien_hang2 = SysFunc.Round(nValue2 * tien_hang1, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    if (tien_hang2 != new Decimal(0))
                                        e.Cell.Record.Cells["tien2"].Value = (object)tien_hang2;
                                    Decimal ck1 = SysFunc.Round(tien_hang1 * nValue1 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND_NT));
                                    e.Cell.Record.Cells["ck_nt"].Value = (object)ck1;
                                    Decimal num13 = this.TinhThue(tien_hang1, ck1, thue_suat, this.Chkthue_ck0.IsChecked.Value);
                                    Decimal num14 = !(this.txtMa_nt.Text != StartUpTrans.M_ma_nt0) ? SysFunc.Round(num13, StartUpTrans.M_ROUND) : SysFunc.Round(num13, StartUpTrans.M_ROUND_NT);
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)num14;
                                    Decimal ck2 = SysFunc.Round(ck1 * nValue2, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                    e.Cell.Record.Cells["ck"].Value = (object)ck2;
                                    Decimal num15 = SysFunc.Round(this.TinhThue(tien_hang2, ck2, thue_suat, this.Chkthue_ck0.IsChecked.Value), StartUpTrans.M_ROUND);
                                    e.Cell.Record.Cells["thue"].Value = (object)num15;
                                }
                                if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                                {
                                    e.Cell.Record.Cells["tien2"].Value = e.Cell.Record.Cells["tien_nt2"].Value;
                                    e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
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
                        case "ma_thue_i":
                            AutoCompleteTextBox autoCompleteControl2 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl2.RowResult != null)
                            {
                                e.Cell.Record.Cells["tk_thue_i"].Value = autoCompleteControl2.RowResult["tk_thue_co"];
                                AutoCompleteTextBox autoCompleteControl3 = ControlFunction.GetAutoCompleteControl(CellValuePresenter.FromCell(e.Cell.Record.Cells["tk_thue_i"]).Editor as ControlHostEditor);
                                if (autoCompleteControl3.RowResult == null)
                                    autoCompleteControl3.SearchInit();
                                e.Cell.Record.Cells["loai_tk_thue"].Value = (object)0;
                                if (autoCompleteControl3.RowResult != null)
                                {
                                    e.Cell.Record.Cells["tk_cn"].Value = autoCompleteControl3.RowResult["tk_cn"];
                                    e.Cell.Record.Cells["loai_tk_thue"].Value = autoCompleteControl3.RowResult["loai_tk"];
                                }
                                e.Cell.Record.Cells["thue_suati"].Value = autoCompleteControl2.RowResult["thue_suat"];
                            }
                            if (!string.IsNullOrEmpty(e.Cell.Record.Cells["thue_suati"].Value.ToString()))
                            {
                                bool? isChecked = this.Chkthue_ck0.IsChecked;
                                if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                                {
                                    Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                                    Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["tien2"].Value, new Decimal(0));
                                    Decimal num15 = this.ParseDecimal(e.Cell.Record.Cells["thue_suati"].Value, new Decimal(0));
                                    Decimal num16 = num13 * num15 / new Decimal(100);
                                    Decimal num17 = this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num16, (int)Convert.ToInt16(StartUpTrans.M_ROUND)) : SysFunc.Round(num16, (int)Convert.ToInt16(StartUpTrans.M_ROUND_NT));
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)num17;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(num14 * num15 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                }
                                else
                                {
                                    Decimal num13 = this.ParseDecimal(e.Cell.Record.Cells["tien_nt2"].Value, new Decimal(0));
                                    Decimal num14 = this.ParseDecimal(e.Cell.Record.Cells["tien2"].Value, new Decimal(0));
                                    Decimal num15 = this.ParseDecimal(e.Cell.Record.Cells["thue_suati"].Value, new Decimal(0));
                                    Decimal num16 = this.ParseDecimal(e.Cell.Record.Cells["ck_nt"].Value, new Decimal(0));
                                    Decimal num17 = this.ParseDecimal(e.Cell.Record.Cells["ck"].Value, new Decimal(0));
                                    Decimal num18 = (num13 - num16) * num15 / new Decimal(100);
                                    Decimal num19 = this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? SysFunc.Round(num18, (int)Convert.ToInt16(StartUpTrans.M_ROUND)) : SysFunc.Round(num18, (int)Convert.ToInt16(StartUpTrans.M_ROUND_NT));
                                    e.Cell.Record.Cells["thue_nt"].Value = (object)num19;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round((num14 - num17) * num15 / new Decimal(100), (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                                }
                                if (this.txtMa_nt.Text == StartUpTrans.M_ma_nt0)
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "tk_thue_i":
                            if (e.Editor.Value == null)
                                break;
                            AutoCompleteTextBox autoCompleteControl4 = ControlFunction.GetAutoCompleteControl(e.Editor as ControlHostEditor);
                            if (autoCompleteControl4.RowResult != null)
                            {
                                e.Cell.Record.Cells["tk_cn"].Value = autoCompleteControl4.RowResult["tk_cn"];
                                break;
                            }
                            break;
                        case "thue_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    e.Cell.Record.Cells["thue"].Value = e.Cell.Record.Cells["thue_nt"].Value;
                                }
                                else
                                {
                                    Decimal nValue = this.txtTy_gia.nValue;
                                    e.Cell.Record.Cells["thue"].Value = (object)SysFunc.Round(Convert.ToDecimal(e.Cell.Record.Cells["thue_nt"].Value) * nValue, StartUpTrans.M_ROUND);
                                }
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "ck_nt":
                            if (e.Cell.IsDataChanged)
                            {
                                if (e.Editor.Value == null || e.Editor.Value != null && e.Editor.Value.ToString().Trim() == "")
                                    e.Cell.Record.Cells["ck_nt"].Value = (object)0;
                                if (this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                                {
                                    e.Cell.Record.Cells["ck"].Value = e.Cell.Record.Cells["ck_nt"].Value;
                                }
                                else
                                {
                                    Decimal num13 = SysFunc.Round(this.ParseDecimal(e.Cell.Record.Cells["ck_nt"].Value, new Decimal(0)) * num1, StartUpTrans.M_ROUND);
                                    if (num13 != new Decimal(0))
                                        e.Cell.Record.Cells["ck"].Value = (object)num13;
                                }
                                this.UpdateTotalHT();
                                break;
                            }
                            break;
                        case "thue":
                        case "ck":
                            this.UpdateTotalHT();
                            break;
                    }
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

        private void GrdCt_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmArcthd1.IsInEditMode.Value || (!Keyboard.IsKeyDown(Key.N) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl)))
                return;
            this.NewRowCt();
            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            if (!FrmArcthd1.IsInEditMode.Value)
                return;
            switch (e.Key)
            {
                case Key.F4:
                    switch (Keyboard.Modifiers)
                    {
                        case ModifierKeys.None:
                            this.NewRowCt();
                            this.GrdCt.ActiveRecord = this.GrdCt.Records[this.GrdCt.Records.Count - 1];
                            this.GrdCt.ActiveCell = (this.GrdCt.ActiveRecord as DataRecord).Cells[0];
                            break;
                        case ModifierKeys.Control:
                            this.InsertRecord((Action)(() => this.NewRowCt()), this.GrdCt, "tk_dt");
                            break;
                    }
                    break;
                case Key.F8:
                    if (ExMessageBox.Show(445, StartupBase.SasObj, "Có xóa dòng ghi hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
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

        private Decimal TinhThue(
          Decimal tien_hang,
          Decimal ck,
          Decimal thue_suat,
          bool TinhThueTruocCK)
        {
            Decimal num = new Decimal(0);
            return !TinhThueTruocCK ? (tien_hang - ck) * thue_suat / new Decimal(100) : tien_hang * thue_suat / new Decimal(100);
        }

        private void UpdateTotalHT()
        {
            try
            {
                if (FormTrans.currActionTask == ActionTask.View)
                    return;
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                Decimal result1 = new Decimal(0);
                Decimal result2 = new Decimal(0);
                Decimal? nullable1 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien_nt2")));
                if (nullable1.HasValue)
                    Decimal.TryParse(nullable1.ToString(), out result1);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_hg"] = (object)SysFunc.Round(result1, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_hg_nt"] = (object)SysFunc.Round(result1, StartUpTrans.M_ROUND_NT);
                Decimal result3 = new Decimal(0);
                Decimal? nullable2 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("ck_nt")));
                if (nullable2.HasValue)
                    Decimal.TryParse(nullable2.ToString(), out result3);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck_nt"] = (object)SysFunc.Round(result3, StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"] = (object)SysFunc.Round(result3, StartUpTrans.M_ROUND);
                Decimal? nullable3 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("thue_nt")));
                if (nullable3.HasValue)
                    Decimal.TryParse(nullable3.ToString(), out result2);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue_nt"] = (object)SysFunc.Round(result2, StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)SysFunc.Round(result2, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)SysFunc.Round(result1 - result3 + result2, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)SysFunc.Round(result1 - result3 + result2, StartUpTrans.M_ROUND_NT);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = (object)SysFunc.Round(result1 - result3, StartUpTrans.M_ROUND);
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt2"] = (object)SysFunc.Round(result1 - result3, StartUpTrans.M_ROUND_NT);
                if (!this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                {
                    Decimal result4 = new Decimal(0);
                    Decimal? nullable4 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("ck")));
                    if (nullable4.HasValue)
                        Decimal.TryParse(nullable4.ToString(), out result4);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_ck"] = (object)SysFunc.Round(result4, StartUpTrans.M_ROUND);
                    Decimal result5 = new Decimal(0);
                    Decimal? nullable5 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("thue")));
                    if (nullable5.HasValue)
                        Decimal.TryParse(nullable5.ToString(), out result5);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] = (object)SysFunc.Round(result5, StartUpTrans.M_ROUND);
                    Decimal result6 = new Decimal(0);
                    Decimal? nullable6 = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Sum<DataRow>((Func<DataRow, Decimal?>)(x => x.Field<Decimal?>("tien2")));
                    if (nullable6.HasValue)
                        Decimal.TryParse(nullable6.ToString(), out result6);
                    Decimal num = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"] == DBNull.Value ? new Decimal(0) : Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_thue"]);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_hg"] = (object)SysFunc.Round(result6, StartUpTrans.M_ROUND);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt"] = (object)SysFunc.Round(result6 - result4 + num, StartUpTrans.M_ROUND);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien2"] = (object)SysFunc.Round(result6 - result4, StartUpTrans.M_ROUND);
                }
                else
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tien_hg_nt"] = (object)SysFunc.Round(result1, StartUpTrans.M_ROUND_NT);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tt_nt"] = (object)SysFunc.Round(result1 - result3 + result2, StartUpTrans.M_ROUND_NT);
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void UpdateTotalThue()
        {
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
            StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["t_thue_nt"] = (object)num6;
            StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["t_thue"] = (object)num7;
        }

        public int ParseInt(object obj, int defaultvalue)
        {
            int result = defaultvalue;
            int.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void Chksua_tk_thue_Checked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_HT_Thue.Value = true;
            this.txtTk_du_voi_Tk_thue.IsFocus = true;
        }

        private void Chksua_tk_thue_Unchecked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_HT_Thue.Value = false;
            this.txtHan_tt.Focus();
        }

        private void Chksua_thue_Checked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_Thue.Value = true;
        }

        private void Chksua_thue_Unchecked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_Thue.Value = false;
            for (int index = 0; index < StartUpTrans.DsTrans.Tables[1].DefaultView.Count; ++index)
            {
                DataRowView dataRowView = StartUpTrans.DsTrans.Tables[1].DefaultView[index];
                if (this.txtTy_gia.Value != null && !string.IsNullOrEmpty(dataRowView["tien_nt2"].ToString()))
                {
                    Decimal num1 = new Decimal(0);
                    Decimal num2 = new Decimal(0);
                    Decimal num3 = new Decimal(0);
                    Decimal num4 = new Decimal(0);
                    Decimal num5 = new Decimal(0);
                    Decimal num6 = new Decimal(0);
                    Decimal num7 = new Decimal(0);
                    Decimal num8 = this.ParseDecimal((object)dataRowView["tl_ck"].ToString(), new Decimal(0));
                    Decimal num9 = this.ParseDecimal((object)dataRowView["tien_nt2"].ToString(), new Decimal(0));
                    Decimal nValue = this.txtTy_gia.nValue;
                    Decimal num10 = this.ParseDecimal((object)dataRowView["thue_suati"].ToString(), new Decimal(0));
                    Decimal num11 = SysFunc.Round(nValue * num9, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                    Decimal num12 = SysFunc.Round(num9 * num8 / new Decimal(100) * nValue, (int)Convert.ToInt16(StartUpTrans.M_ROUND));
                    bool? isChecked = this.Chksua_tien.IsChecked;
                    if (!isChecked.Value || this.txtMa_nt.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
                    {
                        if (num11 > new Decimal(0))
                            dataRowView["tien2"] = (object)SysFunc.Round(num11, StartUpTrans.M_ROUND);
                        if (num12 > new Decimal(0))
                            dataRowView["ck"] = (object)SysFunc.Round(num12, StartUpTrans.M_ROUND);
                    }
                    Decimal tien_hang1 = num11;
                    Decimal ck1 = num12;
                    Decimal thue_suat1 = num10;
                    isChecked = this.Chkthue_ck0.IsChecked;
                    int num13 = isChecked.Value ? 1 : 0;
                    Decimal num14 = this.TinhThue(tien_hang1, ck1, thue_suat1, num13 != 0);
                    dataRowView["thue"] = (object)SysFunc.Round(num14, StartUpTrans.M_ROUND);
                    Decimal num15 = SysFunc.Round(num9 * num8 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                    if (num15 > new Decimal(0))
                        dataRowView["ck_nt"] = (object)num15;
                    Decimal tien_hang2 = num9;
                    Decimal ck2 = num15;
                    Decimal thue_suat2 = num10;
                    isChecked = this.Chkthue_ck0.IsChecked;
                    int num16 = isChecked.Value ? 1 : 0;
                    Decimal num17 = this.TinhThue(tien_hang2, ck2, thue_suat2, num16 != 0);
                    dataRowView["thue_nt"] = (object)SysFunc.Round(num17, StartUpTrans.M_ROUND_NT);
                }
            }
            this.UpdateTotalHT();
        }

        private void Chksua_tien_Checked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tien.Value = true;
        }

        private void Chksua_tien_Unchecked(object sender, RoutedEventArgs e)
        {
            this.IsCheckedSua_tien.Value = false;
            bool? isChecked = this.Chksua_tien.IsChecked;
            if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0 || !sender.GetType().Name.Equals("CheckBox"))
                return;
            this.UpdateTotalChkSua_tien();
            this.CalculateTyGia();
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
                if (num2 != new Decimal(0) && num3 != new Decimal(0))
                {
                    Decimal num4 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND_NT);
                    StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tien_nt2"] = (object)num4;
                    Decimal num5 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["tl_ck"], new Decimal(0));
                    if (num5 != new Decimal(0))
                    {
                        Decimal num6 = SysFunc.Round(num4 * num5 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck_nt"] = (object)num6;
                        Decimal num7 = SysFunc.Round(num6 * num1, StartUpTrans.M_ROUND);
                        if (num7 != new Decimal(0))
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["ck"] = (object)num7;
                    }
                    Decimal num8 = this.ParseDecimal(StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue_suati"], new Decimal(0));
                    if (num8 != new Decimal(0))
                    {
                        Decimal num6 = SysFunc.Round(num4 * num8 / new Decimal(100), StartUpTrans.M_ROUND_NT);
                        StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue_nt"] = (object)num6;
                        Decimal num7 = SysFunc.Round(num6 * num1, StartUpTrans.M_ROUND);
                        if (num7 != new Decimal(0))
                            StartUpTrans.DsTrans.Tables[1].DefaultView[index]["thue"] = (object)num7;
                    }
                }
            }
            this.UpdateTotalHT();
        }

        public override string GetLanguageString(string code, string language)
        {
            return StartUp.GetLanguageString(code, language);
        }

        private void Chkthue_ck0_Checked(object sender, RoutedEventArgs e)
        {
            if (this.IsCheckedSua_tien.Value)
                return;
            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
            {
                Decimal num1 = this.ParseDecimal(dataRowView["tien_nt2"], new Decimal(0));
                Decimal num2 = this.ParseDecimal(dataRowView["tien2"], new Decimal(0));
                Decimal num3 = this.ParseDecimal(dataRowView["thue_suati"], new Decimal(0)) / new Decimal(100);
                Decimal num4 = SysFunc.Round(num1 * num3, StartUpTrans.M_ROUND_NT);
                Decimal num5 = SysFunc.Round(num2 * num3, StartUpTrans.M_ROUND);
                dataRowView["thue_nt"] = (object)num4;
                dataRowView["thue"] = (object)num5;
            }
            this.UpdateTotalHT();
        }

        private void Chkthue_ck0_Unchecked(object sender, RoutedEventArgs e)
        {
            if (this.IsCheckedSua_tien.Value)
                return;
            foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
            {
                Decimal num1 = this.ParseDecimal(dataRowView["tien_nt2"], new Decimal(0));
                Decimal num2 = this.ParseDecimal(dataRowView["tien2"], new Decimal(0));
                Decimal num3 = this.ParseDecimal(dataRowView["ck_nt"], new Decimal(0));
                Decimal num4 = this.ParseDecimal(dataRowView["ck"], new Decimal(0));
                Decimal num5 = this.ParseDecimal(dataRowView["thue_suati"], new Decimal(0)) / new Decimal(100);
                Decimal num6 = SysFunc.Round((num1 - num3) * num5, StartUpTrans.M_ROUND_NT);
                Decimal num7 = SysFunc.Round((num2 - num4) * num5, StartUpTrans.M_ROUND);
                dataRowView["thue_nt"] = (object)num6;
                dataRowView["thue"] = (object)num7;
            }
            this.UpdateTotalHT();
        }

        private void FormMain_EditModeEnded(object sender, string menuItemName, RoutedEventArgs e)
        {
            if (StartUpTrans.DsTrans.Tables[0].DefaultView.Count <= 0)
                return;
            if (!menuItemName.Equals("btnSave"))
                this.LoadDataDu13();
            this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
        }

        private void txtTk_du_voi_Tk_thue_PreviewGotFocus(
          object sender,
          KeyboardFocusChangedEventArgs e)
        {
            if (this.IsCheckedSua_HT_Thue.Value)
                return;
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Return);
        }

        private void GrdCt_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (FormTrans.currActionTask != ActionTask.Add || (StartUpTrans.DsTrans.Tables[1].DefaultView.Count != 1 || !(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["dien_giaii"].ToString() == string.Empty)))
                return;
            StartUpTrans.DsTrans.Tables[1].DefaultView[0]["dien_giaii"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row["dien_giai"];
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

        private void txtma_bp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (StartUp.M_BP_BH.Equals("0"))
                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Return);
            else if (this.txtma_bp.RowResult != null)
            {
                if (this.M_LAN.ToUpper().Equals("V"))
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp"] = (object)this.txtma_bp.RowResult["ten_bp"].ToString();
                else
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp2"] = (object)this.txtma_bp.RowResult["ten_bp2"].ToString();
            }
            else if (this.M_LAN.ToUpper().Equals("V"))
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp"] = (object)"";
            else
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_bp2"] = (object)"";
        }

        private void txtHan_ck_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtHan_ck.RowResult != null)
            {
                if (this.txtHan_tt.Value == DBNull.Value || this.txtHan_tt.nValue == new Decimal(0))
                    this.txtHan_tt.Value = this.txtHan_ck.RowResult["han_tt"];
                if (this.M_LAN == "V")
                    this.tblten_han_ck.Text = this.txtHan_ck.RowResult["ten_thck"].ToString().Trim();
                else
                    this.tblten_han_ck.Text = this.txtHan_ck.RowResult["ten_thck2"].ToString().Trim();
            }
            else
                this.tblten_han_ck.Text = "";
        }

        private void Post(int ispostck)
        {
            string format = "exec [dbo].{0} @stt_rec,@IsHasPT1";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, "[ARCTHD1-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar, 50).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@IsHasPT1", SqlDbType.Int).Value = ispostck;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private void btnViewPT1_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString().Trim()))
                return;
            if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"].ToString().Trim().Equals("PT1"))
            {
                SqlCommand sqlcmd1 = new SqlCommand();
                sqlcmd1.CommandText = "Select count(1) from ph41 WHERE stt_Rec = @stt_rec_pt";
                sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString();
                if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                {
                    if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu thu không tồn tại, có xóa thông tin phiếu thu trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) != MessageBoxResult.Yes)
                        return;
                    SqlCommand sqlcmd2 = new SqlCommand();
                    sqlcmd2.CommandText = "UPDATE ph21 Set stt_rec_pt = '', so_ct_pt = '', ma_ct_pt = '' WHERE stt_rec = @stt_rec_pt; ";
                    sqlcmd2.CommandText += "UPDATE cttt20 Set tat_toan = 0, stt_rec_tt = '' WHERE stt_rec = @stt_rec";
                    sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    sqlcmd2.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.VarChar)).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                    this.BindingSasObj.ExcuteNonQuery(sqlcmd2);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_pt"] = "";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct_pt"] = "";
                }
                else
                    SysFunc.EditVoucherFromBrowse(this.BindingSasObj, "PT1", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString(), Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), this.BindingSasObj.M_ProcessName);
            }
            else
            {
                SqlCommand sqlcmd1 = new SqlCommand();
                sqlcmd1.CommandText = "Select count(1) from ph51 WHERE stt_Rec = @stt_rec_pt";
                sqlcmd1.Parameters.Add(new SqlParameter("@stt_rec_pt", SqlDbType.VarChar)).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_pt"].ToString();
                if ((int)this.BindingSasObj.ExcuteScalar(sqlcmd1) == 0)
                {
                    if (ExMessageBox.Show(693, StartupBase.SasObj, "Phiếu thu không tồn tại, có xóa thông tin phiếu thu trên hóa đơn không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.Yes)
                    {
                        SqlCommand sqlcmd2 = new SqlCommand();
                        sqlcmd2.CommandText = "UPDATE ph21 Set stt_rec_pt = '', so_ct_pt = '', ma_ct_pt = '' WHERE stt_rec = @stt_rec_pt; ";
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

        private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (this.IsSequenceSave)
                return;
            InputLanguageManager.Current.CurrentInputLanguage = new CultureInfo("en-US");
            if (Keyboard.IsKeyDown(Key.Return) && (Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt)))
            {
                TextBox textBox = sender as TextBox;
                textBox.SelectedText = Environment.NewLine;
                ++textBox.SelectionStart;
                textBox.SelectionLength = 0;
                e.Handled = true;
            }
            else if (Keyboard.IsKeyDown(Key.Return))
            {
                (this.Toolbar.FindName("btnSave") as SasVoucherLib.ToolBarButton).Focus();
                e.Handled = true;
            }
        }

        private void txtNgay_lhd_GotFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct.Value == DBNull.Value)
                return;
            if (this.txtNgay_lhd.Value == DBNull.Value)
                this.txtNgay_lhd.Value = this.txtNgay_ct.Value;
        }

        private void txtsd_hddt_yn_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtsd_hddt_yn.Text))
                return;
            this.txtsd_hddt_yn.Text = "0";
        }

        private void GrdCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!FrmArcthd1.IsInEditMode.Value)
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

        private void btnAddInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (FrmArcthd1.IsInEditMode.Value)
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
                    SasCusEinvoice.UpdateInvoice updateInvoice = new SasCusEinvoice.UpdateInvoice(StartupBase.SasObj, ds.Copy(), "BKAV", "HD1");
                    ds.Clear();
                    ds.Dispose();
                    msg = updateInvoice.BKAV_Create(out sGUID, out sInvoiceForm, out sInvoiceSerial, out sInvoiceNo);
                    frmWaiting.Set(300);
                    if (msg.Length > 0)
                    {
                        MessageBox.Show(msg,"Thong bao");
                    }
                    else
                    {
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["inv_guid"] = sGUID.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_hddt"] = sInvoiceNo.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri_hddt"] = sInvoiceSerial.ToString().Trim();
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["mau_hddt"] = sInvoiceForm.ToString().Trim();

                        SqlCommand sqlcmd1 = new SqlCommand();
                        sqlcmd1.CommandText = "Update PH21 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CTTT20 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                        sqlcmd1.CommandText += "; \n Update CT00 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
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
            if (FrmArcthd1.IsInEditMode.Value)
                return;

            FrmWaiting frmWaiting = new FrmWaiting(600);
            try
            {
                frmWaiting.Show();
                DataTable dterror = new DataTable();
                dterror.Columns.Add("tag", typeof(string));
                dterror.Columns.Add("so_ct", typeof(string));
                dterror.Columns.Add("error", typeof(string));
                frmWaiting.Set(100);
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
                    frmWaiting.Set(200);
                    if (ds != null)
                    {
                        SasCusEinvoice.UpdateInvoice updateInvoice = new SasCusEinvoice.UpdateInvoice(StartupBase.SasObj, ds.Copy(), "BKAV", "HD1");
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
                            sqlcmd1.CommandText = "Update PH21 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CTTT20 SET inv_guid = @sGUID, mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
                            sqlcmd1.CommandText += "; \n Update CT00 SET mau_hddt= @sInvoiceForm, so_seri_hddt = @sInvoiceSerial, so_ct_hddt = @sInvoiceNo WHERE stt_rec = '" + voucherIds + "'";
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
            if (FrmArcthd1.IsInEditMode.Value)
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

    }
}
