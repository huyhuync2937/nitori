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
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
//using ArapLib;
namespace SODNBG1
{
    public partial class FrmPoctpna : FormTrans
    {
        private DataTable dtMa_ncc;
        public static string KeyFilter = "";

        public static int iRow = 0;
        public static int OldiRow = 0;
        public static string[] FieldCk = new string[0]
        {
    
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
        public FrmPoctpna()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.C_QS = this.txtMa_qs;
            this.C_NgayHT = this.txtNgay_ct;         
            this.C_So_ct = this.txtSo_ct;
            this.dtMa_ncc = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT CAST(0 as BIT) as tag, ma_kh, ten_kh, ten_kh2,e_mail, ma_dvcs FROM dmkh")).Tables[0];

        }

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.BindingSasObj = StartupBase.SasObj;
                StartUp.M_AR_CK = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_CK"));
                StartUp.M_AR_TT = (int)Convert.ToInt16(this.BindingSasObj.GetOption(this.stt_mau_temlate.ToString(), "M_AR_TT"));
                FormTrans.currActionTask = ActionTask.View;
                if (StartUpTrans.DsTrans.Tables[0].Rows.Count > 1)
                    FrmPoctpna.iRow = StartUpTrans.DsTrans.Tables[0].Rows.Count - 1;
                FrmPoctpna.IsInEditMode = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsInEditMode");
                this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Ma_nt0");
                this.Voucher_Lan0 = (CodeValueBindingObject)this.FormMain.FindResource((object)"Voucher_Lan0");
                this.IsCheckedSua_tien = (CodeValueBindingObject)this.FormMain.FindResource((object)"IsCheckedSua_tien");
                this.Ty_Gia_ValueChange = (CodeValueBindingObject)this.FormMain.FindResource((object)"Ty_Gia_ValueChange");

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
                    this.Voucher_Lan0.Value = this.M_LAN.Trim().Equals("V");
                }
                this.Voucher_Ma_nt0.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
                this.Voucher_Ma_nt0.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0);
               
                if (StartUp.M_AR_CK == 0)
                {
                   
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
            this.gridlayout50.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;
            this.GrdLayout10.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;          
            this.GrdLayout21.DataContext = (object)StartUpTrans.DsTrans.Tables[0].DefaultView;

            this.GrdCt.DataSource = (IEnumerable)StartUpTrans.DsTrans.Tables[1].DefaultView;
            // this.txtStatus.ItemsSource = (IEnumerable)StartUpTrans.tbStatus.DefaultView;

            if (StartUpTrans.tbStatus.DefaultView.Count != 1)
                return;
            //   this.txtStatus.IsEnabled = false;
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_ncc.Focus()));
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
                row["ty_giaf"] = !row["ma_nt"].ToString().Trim().Equals(StartUpTrans.M_ma_nt0.Trim()) ? (object)StartUp.GetRates(row["ma_nt"].ToString().Trim(), Convert.ToDateTime(row["ngay_ct"]).Date) : (object)1;
                row["status"] = (object)"1";
           
             
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_ncc.Focus()));
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
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.txtMa_ncc.Focus()));
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
            StartUp.HiddenFieldIsSetted = true;
            DataTable dataTable = StartUpTrans.DsTrans.Tables[0].Copy();
            dataTable.Rows.RemoveAt(0);
            FormView formView = new FormView(StartupBase.SasObj, dataTable.DefaultView, StartUpTrans.DsTrans.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
            formView.ListFieldSum = "t_tt_nt;t_tt";
            formView.frmBrw.Title = SysFunc.Cat_Dau(this.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
            FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "SODNBG1_4";
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
                dataRow["so_luong"] = (object)0;            
          
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

                                break;
                            }
                            break;
                        case "so_luong":
                            try
                            {
                                if (e.Editor.Value == DBNull.Value)
                                    e.Cell.Record.Cells["so_luong"].Value = (object)0;
                                this.Sum_ALL();
                                break;
                            }
                            catch (Exception ex)
                            {
                                ErrorLog.CatchMessage(ex);
                                break;
                            }
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

                    this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                    this.GrdCt.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                    if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
                    {
                        TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                        if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                            return;
                    }
                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ncc"].ToString()))
                    {
                        int num = (int)ExMessageBox.Show(2230, StartupBase.SasObj, "Chưa vào mã nhà cung cấp!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.txtMa_ncc.Focus();
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

                        }
                        this.Sum_ALL();
                        Decimal result1 = new Decimal(0);
                        Decimal result2 = new Decimal(0);
                        Decimal num = this.ParseDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"], new Decimal(0));
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien_nt)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result1);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[1].Compute("sum(tien)", "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'").ToString(), out result2);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien_nt"] = (object)result1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"] = (object)result2;

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
                        SysFunc.SetStrSearch(StartupBase.SasObj, "ph110", ref table);
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

                    if (!DataProvider.UpdateCtTable(StartupBase.SasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), LocalTable2, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                    {
                        int num1 = (int)ExMessageBox.Show(2285, StartupBase.SasObj, "Lưu không thành công, kiểm tra lại dữ liệu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                string cmdText = "SELECT MAX(so_ct) FROM ph110 WHERE ma_qs='" + ma_qs.Trim() + "';";
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
      

        private void txtTy_gia_GotFocus(object sender, RoutedEventArgs e)
        {
            if (!this.Voucher_Ma_nt0.Value)
                return;
           
            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
        }

        private void txtTy_gia_LostFocus(object sender, RoutedEventArgs e)
        {
           
        }

        public void TyGiaValueChange()
        {
           
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
            this.txtSo_ct.MaxLength = ListFunc.GetLengthColumn(ListFunc.GetSqlTableFieldList(StartupBase.SasObj, "v_PH110"), "so_ct");
        }
        private void Sum_ALL()
        {
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
            StartUpTrans.DsTrans.Tables[1].AcceptChanges(); 
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_so_luong"] = (object)this.ParseDecimal((object)StartUpTrans.DsTrans.Tables[1].Compute("sum(so_luong)", StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter).ToString(), new Decimal(0));
        }

        private void IsVisibilityFieldsXamDataGrid(string ma_nt)
        {          
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
            this.ChangeLanguage();
        }

        private void IsVisibilityFieldsXamDataGridBySua_Tien()
        {

        }

        private void PhanBoThueInCT()
        {
           
        }

        private void PhanBo()
        {
           
        }

        private void txttong_cp_nt_LostFocus(object sender, RoutedEventArgs e)
        {
            if (FormTrans.currActionTask == ActionTask.Delete || FormTrans.currActionTask == ActionTask.View)
                return;
            this.IsVisibilityFieldsXamDataGridBySua_Tien();
        }

        private void txttong_cp_LostFocus(object sender, RoutedEventArgs e)
        {
           
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

            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck"] = (object)"";
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_thck2"] = (object)"";

        }

        private void Post()
        {
            string format = "exec [dbo].{0} @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 2 ? string.Format(format, (object)"[SODNBG1-Post]") : string.Format(format, (object)StartUpTrans.Post_store[2]));
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
        private void txtSobaogia_GotFocus(object sender, RoutedEventArgs e)
        {

        }

        //private void txtMa_ncc_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        //{
        //   if(!string.IsNullOrEmpty(this.txtMa_ncc.Text) && this.txtMa_ncc.RowResult!=null)
        //    {
        //        if(StartupBase.M_LAN.Equals("V"))
        //        {
        //            this.tblTenncc.Text = this.txtMa_ncc.RowResult["ten_kh"].ToString();
        //        }
        //        else
        //        {
        //            this.tblTenncc.Text = this.txtMa_ncc.RowResult["ten_kh2"].ToString();
        //        }
        //        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dia_chi"] = (object)this.txtMa_ncc.RowResult["dia_chi"];
        //    }    
        //}
       

        private void BtnChondonhang_Click(object sender, RoutedEventArgs e)
        {
            FrmLoc Locdonhang = new FrmLoc();
             Locdonhang.ShowDialog();
            if (StartUp.isOk)
            {
                if (StartUp.HDBData.Tables[1].DefaultView.Count > 0)
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_yeucau"] = StartUp.HDBData.Tables[1].DefaultView[0]["ma_hd"].ToString().Trim();
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
                            dataRow["dvt1"] = r["dvt1"];                           
                            dataRow["so_luong"] = r["so_luong1"];                              
                            dataRow["ghi_chu"] = "";
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
        private void txtMa_ncc_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.F11)
                return;
            this.OnF11();
        }
        private void OnF11()
        {
            string text = this.txtMa_ncc.Text;
            char[] chArray = new char[1] { ',' };
            foreach (object obj in text.Split(chArray))
            {
                foreach (DataRow dataRow in this.dtMa_ncc.Select(string.Format("ma_kh = '{0}'", obj)))
                    dataRow["tag"] = (object)true;
            }
            DataTable dataTable = this.dtMa_ncc.Copy();
            COTKTH2Dvcs cotktH2Dvcs = new COTKTH2Dvcs();
            //dataTable.DefaultView.RowFilter = !string.IsNullOrEmpty(this.txtMa_dvcs.Text.Trim()) ? "ma_dvcs LIKE '" + this.txtMa_dvcs.Text + "'" : "1=1";
            cotktH2Dvcs.GrdCt.DataSource = (IEnumerable)dataTable.DefaultView;
            if (StartupBase.M_LAN != "V")
                cotktH2Dvcs.Title = "Supplier list";
            bool? nullable = cotktH2Dvcs.ShowDialog();
            if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                return;
            DataRow[] dataRowArray = dataTable.Select("tag = 1");
            string str = "";
            foreach (DataRow dataRow in dataRowArray)
                str = str + (str == "" ? "" : ",") + dataRow["ma_kh"].ToString().Trim();
            this.txtMa_ncc.Text = str;
        }
        private void btnSendEmail_Click(object sender, RoutedEventArgs e)
        {
     
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

                    FrmPoctpna.KeyFilter = string.IsNullOrEmpty(this.txtMa_ncc.Text)
                        ? "1=1"
                        : "dbo.InList(ma_kh, '" + this.txtMa_ncc.Text.Trim() + "', ',') = 1";

                    //SqlCommand sql = new SqlCommand("SELECT e_mail FROM dmkh WHERE " + (string.IsNullOrEmpty(FrmPoctpna.KeyFilter) ? "" : " and " + FrmPoctpna.KeyFilter));

                    SqlCommand sql = new SqlCommand("SELECT tk_portal FROM dmkh WHERE " + FrmPoctpna.KeyFilter);

                    DataTable table = StartupBase.SasObj.ExcuteReader(sql).Tables[0];
                    if (table == null || table.Rows.Count == 0)
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
    }
}
