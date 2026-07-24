using ArapLib;
using Infragistics.Windows.DataPresenter;
using SasControls;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using SasLib;
using System;
using System.Linq;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace CACTPC1
{
    public partial class FrmTim : FormFilter
    {
        public new static readonly DependencyProperty SasObjProperty = DependencyProperty.Register(nameof(SasObj), typeof(SasObject), typeof(FrmTim), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));

        public FrmTim()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.M_LAN = StartUpTrans.M_LAN;
        }

        public string M_LAN { get; set; }

        public SasObject SasObj
        {
            get
            {
                return (SasObject)this.GetValue(FrmTim.SasObjProperty);
            }
            set
            {
                this.SetValue(FrmTim.SasObjProperty, (object)value);
            }
        }

        public FrmTim(SasObject _SasObj, string _filterID, string _tableList)
        {
            this.InitializeComponent();
            this.SasObj = _SasObj;
            this.BindingSasObj = _SasObj;
            this.GridSearch.filterID = _filterID;
            this.GridSearch.tableList = _tableList;
            this.GridSearch.SasObj = _SasObj;
            this.Title = SysFunc.Cat_Dau(this.Title);
        }

        private void FrmTim_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void FrmTim_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMa_qs.Filter = !(StartUpTrans.Ma_ct == "BN1") ? "ma_cts like '%PC1%' and status = 1" : "ma_cts like '%BN1%' and status = 1";
            this.txtloc_nsd.Value = StartUpTrans.DmctInfo["m_loc_nsd"] == DBNull.Value ? (object)0 : StartUpTrans.DmctInfo["m_loc_nsd"];
            this.txtNgay_ct1.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct1");
            this.txtNgay_ct2.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct2");
            this.txtNgay_ct1.Focus();
            this.txtMa_kh.SearchInit();
            this.txtTk_co.SearchInit();
            this.txtMaDVCS.SearchInit();
            this.lblTenDVCS.Text = this.txtMaDVCS.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMaDVCS.RowResult["ten_dvcs"].ToString() : this.txtMaDVCS.RowResult["ten_dvcs2"].ToString());
            this.lbltenkh.Text = this.txtMa_kh.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtMa_kh.RowResult["ten_kh"].ToString() : this.txtMa_kh.RowResult["ten_kh2"].ToString());
            this.lbltentk.Text = this.txtTk_co.RowResult == null ? "" : (StartupBase.SasObj.GetOption("M_LAN").ToString() == "V" ? this.txtTk_co.RowResult["ten_tk"].ToString() : this.txtTk_co.RowResult["ten_tk2"].ToString());
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               try
               {
                   DataSet dataSet = StartupBase.SasObj.ExcuteReader(new SqlCommand("select CAST(count(1) as Numeric(16, 0)) from dmdvcs"));
                   if (dataSet == null || dataSet.Tables.Count <= 0 || dataSet.Tables[0].Rows.Count <= 0)
                       return;
                   this.txtMaDVCS.IsReadOnly = FNum.ToDec(dataSet.Tables[0].Rows[0][0]) <= new Decimal(1);
               }
               catch (Exception ex)
               {
                   Debug.WriteLine(ex.Message);
               }
           }), DispatcherPriority.Background);
            DataView defaultView = this.SasObj.GetPostInfo(StartUpTrans.Ma_ct).DefaultView.ToTable().DefaultView;
            DataRow row = defaultView.Table.NewRow();
            row["ten_post"] = row["ten_act"] = (object)"Tất cả";
            row["ten_post2"] = row["ten_act2"] = (object)"All";
            defaultView.Table.Rows.Add(row);
            this.txtStatus.ItemsSource = (IEnumerable)defaultView.Table.AsEnumerable().OrderBy<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("ma_post"))).AsDataView<DataRow>();
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtStatus.SelectedIndex = 0), DispatcherPriority.Background);
            if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
                this.txtStatus.DisplayMemberPath = "ten_post";
            else
                this.txtStatus.DisplayMemberPath = "ten_post2";
        }

        private string GetPhFilterExpr()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = "1=1 ";
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
                str = str + " and ngay_ct >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
                str = str + " and ngay_ct <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
                str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
                str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtMa_kh.Text))
                str = str + " and ma_kh = " + this.ConvertDataToSql((object)this.txtMa_kh.Text.Trim(), typeof(string));
            if (!string.IsNullOrEmpty(this.txtTk_co.Text))
                str = str + " and tk like " + this.ConvertDataToSql((object)(this.txtTk_co.Text.Trim() + "%"), typeof(string));
            if (Convert.ToInt16(this.txtloc_nsd.Value) == (short)1)
                str = str + " and [user_id] = " + (object)StartUpTrans.M_User_Id;
            if (!string.IsNullOrEmpty(this.txtMaDVCS.Text))
                str = str + " and ma_dvcs LIKE '" + this.txtMaDVCS.Text.Trim() + "%'";
            if (!string.IsNullOrEmpty(this.txtMa_qs.Text))
                str = str + " and ma_qs LIKE '" + this.txtMa_qs.Text.Trim() + "%'";
            if (this.txtStatus.Value != null && !string.IsNullOrEmpty(this.txtStatus.Value.ToString().Trim()))
                str = str + " and status = '" + this.txtStatus.Value.ToString().Trim() + "'";
            if (!SysFunc.CheckPermission(this.SasObj, ActionTask.View, StartupBase.Menu_Id))
                str = str + " AND user_id0 = " + this.SasObj.UserInfo.Rows[0]["user_id"].ToString();
            if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
                str = str + " and " + this.GridSearch.arrStrFilter[0];
            return str;
        }

        private string GetCtFilterExpr()
        {
            string str = "1=1";
            if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[1]))
                str = str + " and " + this.GridSearch.arrStrFilter[1];
            return str;
        }

        private string GetCtgtFilterExpr()
        {
            string str = "1=1";
            if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[2]))
                str = str + " and " + this.GridSearch.arrStrFilter[2];
            return str;
        }

        public string ConvertDataToSql(object value, Type ValueType)
        {
            string str;
            switch (ValueType.ToString())
            {
                case "System.String":
                    str = string.Format("'{0}'", (object)(value as string).Replace("'", "'"));
                    break;
                case "System.DateTime":
                    str = string.Format("'{0}'", (object)((DateTime)value).ToString("yyyyMMdd"));
                    break;
                default:
                    str = string.Format("'{0}'", value);
                    break;
            }
            return str;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)) && !(Keyboard.FocusedElement as TextBoxAutoComplete).ParentControl.CheckLostFocus() || !this.IsHopLe())
                    return;
                this.SasObj.SetSysvar("M_ngay_ct1", (object)this.txtNgay_ct1.dValue);
                this.SasObj.SetSysvar("M_ngay_ct2", (object)this.txtNgay_ct2.dValue);
                bool flag = false;
                this.GridSearch._GenerateSQLString();
                this.GridSearch.GrdSearch.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                StartUp.TransFilterCmd.Parameters["@PhFilter"].Value = (object)this.GetPhFilterExpr();
                StartUp.TransFilterCmd.Parameters["@CtFilter"].Value = (object)this.GetCtFilterExpr();
                StartUp.TransFilterCmd.Parameters["@GtFilter"].Value = (object)this.GetCtgtFilterExpr();
                StartUp.TransFilterCmd.Parameters["@sl_ct"].Value = (object)-1;
                DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                string str1 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tt"))).Sum().Value.ToString(this.SasObj.GetOption("M_IP_TIEN").ToString());
                string str2 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tt_nt"))).Sum().Value.ToString(this.SasObj.GetOption("M_IP_TIEN_NT").ToString());
                int count1 = dataSet.Tables[0].Rows.Count;
                if (count1 > 0)
                {
                    flag = true;
                    int num = (int)ExMessageBox.Show(825, StartupBase.SasObj, "Có [" + (object)count1 + "] chứng từ. Tổng phát sinh  [" + str2 + "] / [" + str1 + "]", "SIS.VN", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(830, StartupBase.SasObj, "Không có chứng từ nào như vậy! ", "SIS.VN", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                if (flag)
                {
                    string strBrowse = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[0];
                    string strBrowseCt = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[1];
                    FormView formView = new FormView(this.SasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, strBrowse, strBrowseCt, "stt_rec");
                    FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
                    formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() + ". Ky " : StartUpTrans.CommandInfo["bar2"].ToString() + ". Period ") + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                    formView.ListFieldSum = "t_tt_nt;t_tt";
                    formView.frmBrw.LanguageID = "CACTPC1_6";
                    formView.ShowDialog();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    int count2 = StartUpTrans.DsTrans.Tables[0].Rows.Count;
                    int count3 = StartUpTrans.DsTrans.Tables[1].Rows.Count;
                    int count4 = StartUpTrans.DsTrans.Tables[2].Rows.Count;
                    for (int index = count2 - 1; index >= 1; --index)
                        StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(index);
                    for (int index = 0; index < count3; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.RemoveAt(0);
                    for (int index = 0; index < count4; ++index)
                        StartUpTrans.DsTrans.Tables[2].Rows.RemoveAt(0);
                    int count5 = dataSet.Tables[0].Rows.Count;
                    for (int index = 0; index < count5; ++index)
                        StartUpTrans.DsTrans.Tables[0].Rows.Add(dataSet.Tables[0].Rows[index].ItemArray);
                    int count6 = dataSet.Tables[1].Rows.Count;
                    for (int index = 0; index < count6; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);
                    int count7 = dataSet.Tables[2].Rows.Count;
                    for (int index = 0; index < count7; ++index)
                        StartUpTrans.DsTrans.Tables[2].Rows.Add(dataSet.Tables[2].Rows[index].ItemArray);
                    if (formView.DataGrid.ActiveRecord != null)
                    {
                        int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
                        if (index == -1)
                            index = 0;
                        string str3 = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                        FrmCACTPC1.iRow = index + 1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
                    }
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private bool IsHopLe()
        {
            if (string.IsNullOrEmpty(this.txtNgay_ct1.Value.ToString()) || !this.txtNgay_ct1.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(835, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "SIS.VN", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (string.IsNullOrEmpty(this.txtNgay_ct2.Value.ToString()) || !this.txtNgay_ct2.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(840, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "SIS.VN", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct2.Focus();
                this.txtNgay_ct2.SelectAll();
                return false;
            }
            if (!(this.txtNgay_ct1.dValue > this.txtNgay_ct2.dValue))
                return true;
            int num1 = (int)ExMessageBox.Show(845, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtNgay_ct1.Focus();
            this.txtNgay_ct1.SelectAll();
            return false;
        }

        private void txtloc_nsd_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtloc_nsd.Text))
                return;
            this.txtloc_nsd.Value = StartUpTrans.DmctInfo["m_loc_nsd"] == DBNull.Value ? (object)0 : StartUpTrans.DmctInfo["m_loc_nsd"];
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()))
            {
                this.lbltenkh.Text = "";
            }
            else
            {
                if (this.txtMa_kh.RowResult == null)
                    return;
                this.lbltenkh.Text = !StartUpTrans.M_LAN.ToUpper().Equals("V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
            }
        }

        private void txtTk_co_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtTk_co.Text.Trim()))
            {
                this.lbltentk.Text = "";
            }
            else
            {
                if (this.txtTk_co.RowResult == null)
                    return;
                this.lbltentk.Text = !StartUpTrans.M_LAN.ToUpper().Equals("V") ? this.txtTk_co.RowResult["ten_tk2"].ToString() : this.txtTk_co.RowResult["ten_tk"].ToString();
            }
        }

        private void txtMaDVCS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaDVCS.RowResult != null)
                this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
            else
                this.lblTenDVCS.Text = "";
        }

        private void txtMa_qs_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_qs.Text.Trim()))
            {
                this.lbltenqs.Text = "";
            }
            else
            {
                if (this.txtMa_qs.RowResult == null)
                    return;
                this.lbltenqs.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_qs.RowResult["ten_qs2"].ToString() : this.txtMa_qs.RowResult["ten_qs"].ToString();
            }
        }

    }
}
