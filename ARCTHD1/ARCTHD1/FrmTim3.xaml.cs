using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
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
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace ARCTHD1
{
    public partial class FrmTim3 : FormFilter
    {

        public FrmTim3(SasObject _SasObj, string _filterID, string _tableList)
        {
            this.InitializeComponent();
            this.BindingSasObj = _SasObj;
            this.GridSearch.filterID = _filterID;
            this.GridSearch.tableList = _tableList;
            this.GridSearch.SasObj = _SasObj;
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmTim_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void FrmTim_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtloc_nsd.Value = StartUpTrans.DmctInfo["m_loc_nsd"] == DBNull.Value ? (object)0 : StartUpTrans.DmctInfo["m_loc_nsd"];
            this.txtNgay_ct1.Value = (object)(DateTime)this.BindingSasObj.GetSysvar("M_ngay_ct1");
            this.txtNgay_ct2.Value = (object)(DateTime)this.BindingSasObj.GetSysvar("M_ngay_ct2");
            this.txtMa_kh.SearchInit();
            this.txtMaDVCS.SearchInit();
            this.txtTk_no.SearchInit();
            DataView defaultView = this.BindingSasObj.GetPostInfo(StartUpTrans.Ma_ct).DefaultView;
            DataRow row = defaultView.Table.NewRow();
            row["ten_post"] = (object)"Tất cả";
            row["ten_post2"] = (object)"All";
            defaultView.Table.Rows.Add(row);
            this.txtStatus.ItemsSource = (IEnumerable)defaultView.Table.AsEnumerable().OrderBy<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("ma_post"))).AsDataView<DataRow>();
            this.txtStatus.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.txtStatus.SelectedIndex != -1)
                   return;
               this.txtStatus.SelectedIndex = 0;
           }));
            if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
            {
                if (this.txtMa_kh.RowResult != null)
                    this.lbltenkh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
                if (this.txtMaDVCS.RowResult != null)
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
                if (this.txtTk_no.RowResult != null)
                    this.lbltentk.Text = this.txtTk_no.RowResult["ten_tk"].ToString();
                this.txtStatus.DisplayMemberPath = "ten_post";
            }
            else
            {
                if (this.txtMa_kh.RowResult != null)
                    this.lbltenkh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
                if (this.txtMaDVCS.RowResult != null)
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
                if (this.txtTk_no.RowResult != null)
                    this.lbltentk.Text = this.txtTk_no.RowResult["ten_tk2"].ToString();
                this.txtStatus.DisplayMemberPath = "ten_post2";
            }
            this.txtMa_qs.IsFocus = true;
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
            if (!string.IsNullOrEmpty(this.txtTk_no.Text))
                str = str + " and ma_nx like " + this.ConvertDataToSql((object)(this.txtTk_no.Text.Trim() + "%"), typeof(string));
            if (Convert.ToInt16(this.txtloc_nsd.Value) == (short)1)
                str = str + " and [user_id] = " + (object)StartUpTrans.M_User_Id;
            if (!string.IsNullOrEmpty(this.txtMa_qs.Text))
                str = str + " and ma_qs LIKE '" + this.txtMa_qs.Text.Trim() + "%'";
            if (!string.IsNullOrEmpty(this.txtMa_bp.Text))
                str = str + " and ma_bp LIKE '" + this.txtMa_bp.Text.Trim() + "%'";
            if (this.txtStatus.Value != null && !string.IsNullOrEmpty(this.txtStatus.Value.ToString().Trim()))
                str = str + " and status = '" + this.txtStatus.Value.ToString().Trim() + "'";
            if (!string.IsNullOrEmpty(this.txtMaDVCS.Text))
                str = str + " and ma_dvcs LIKE '" + this.txtMaDVCS.Text.Trim() + "%'";
            if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                str = str + " and  AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
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
                this.BindingSasObj.SetSysvar("M_ngay_ct1", (object)this.txtNgay_ct1.dValue);
                this.BindingSasObj.SetSysvar("M_ngay_ct2", (object)this.txtNgay_ct2.dValue);
                bool flag = false;
                this.GridSearch._GenerateSQLString();
                this.GridSearch.GrdSearch.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
                StartUp.TransFilterCmd.Parameters["@PhFilter"].Value = (object)this.GetPhFilterExpr();
                StartUp.TransFilterCmd.Parameters["@CtFilter"].Value = (object)this.GetCtFilterExpr();
                StartUp.TransFilterCmd.Parameters["@GtFilter"].Value = (object)"";
                StartUp.TransFilterCmd.Parameters["@sl_ct"].Value = (object)-1;
                DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                string str1 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tt"))).Sum().Value.ToString(this.BindingSasObj.GetOption("M_IP_TIEN").ToString());
                string str2 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tt_nt"))).Sum().Value.ToString(this.BindingSasObj.GetOption("M_IP_TIEN_NT").ToString());
                int count1 = dataSet.Tables[0].Rows.Count;
                if (count1 > 0)
                {
                    flag = true;
                    int num = (int)ExMessageBox.Show(490, StartupBase.SasObj, "Có [" + (object)count1 + "] chứng từ. Tổng phát sinh  [" + str2 + "] / [" + str1 + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(495, StartupBase.SasObj, "Không có chứng từ nào như vậy! ", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                if (flag)
                {
                    string str3 = "";
                    string str4 = "";
                    if (StartUpTrans.M_LAN.Equals("V"))
                    {
                        if (StartUpTrans.CommandInfo["Vbrowse2"] != null)
                        {
                            string str5 = StartUpTrans.CommandInfo["Vbrowse2"].ToString();
                            str3 = str5.Split('|')[0];
                            str4 = str5.Split('|')[1];
                        }
                    }
                    else if (StartUpTrans.CommandInfo["Ebrowse2"] != null)
                    {
                        string str5 = StartUpTrans.CommandInfo["Ebrowse2"].ToString();
                        str3 = str5.Split('|')[0];
                        str4 = str5.Split('|')[1];
                    }
                    if (StartUp.M_AR_CK == 0)
                    {
                        str3 = StartUp.EditFields(str3);
                        str4 = StartUp.EditFields(str4);
                    }
                    FormView formView = new FormView(this.BindingSasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, str3, str4, "stt_rec");
                    formView.frmBrw.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() + ". Kỳ " : StartUpTrans.CommandInfo["bar2"].ToString() + ". Period ") + this.txtNgay_ct1.Text + " -> " + this.txtNgay_ct2.Text;
                    formView.ListFieldSum = "t_tt_nt;t_tt";
                    FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
                    formView.frmBrw.LanguageID = "ARCTHD1_7";
                    formView.ShowDialog();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[1].Rows[0]["stt_rec"].ToString() + "'";
                    DataTable table1 = StartUpTrans.DsTrans.Tables[0];
                    DataTable table2 = StartUpTrans.DsTrans.Tables[1];
                    int count2 = StartUpTrans.DsTrans.Tables[0].Rows.Count;
                    int count3 = StartUpTrans.DsTrans.Tables[1].Rows.Count;
                    for (int index = count2 - 1; index >= 1; --index)
                        StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(index);
                    for (int index = 0; index < count3; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.RemoveAt(0);
                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                    int count4 = dataSet.Tables[0].Rows.Count;
                    for (int index = 0; index < count4; ++index)
                        StartUpTrans.DsTrans.Tables[0].Rows.Add(dataSet.Tables[0].Rows[index].ItemArray);
                    StartUpTrans.DsTrans.Tables[0].AcceptChanges();
                    int count5 = dataSet.Tables[1].Rows.Count;
                    for (int index = 0; index < count5; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        if (FrmArcthd1.iRow > dataSet.Tables[0].Rows.Count - 1)
                            FrmArcthd1.iRow = dataSet.Tables[0].Rows.Count - 1;
                        StartUp.DataFilter(StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["stt_rec"].ToString());
                    }
                    if (formView.DataGrid.ActiveRecord != null)
                    {
                        int index = (formView.DataGrid.ActiveRecord as DataRecord).Index;
                        if (index >= 0)
                        {
                            string stt_rec = (formView.DataGrid.DataSource as DataView)[index]["stt_rec"].ToString();
                            FrmArcthd1.iRow = index + 1;
                            StartUp.DataFilter(stt_rec);
                        }
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
                int num = (int)ExMessageBox.Show(500, StartupBase.SasObj, "Hãy nhập Từ ngày!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Value.ToString()) && this.txtNgay_ct2.IsValueValid)
                return true;
            int num1 = (int)ExMessageBox.Show(505, StartupBase.SasObj, "Hãy nhập Đến ngày!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtNgay_ct2.Focus();
            this.txtNgay_ct2.SelectAll();
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
                this.lbltenkh.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
            }
        }

        private void txtTk_no_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtTk_no.Text.Trim()))
            {
                this.lbltentk.Text = "";
            }
            else
            {
                if (this.txtTk_no.RowResult == null)
                    return;
                this.lbltentk.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtTk_no.RowResult["ten_tk2"].ToString() : this.txtTk_no.RowResult["ten_tk"].ToString();
            }
        }

        private void txtMaDVCS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaDVCS.RowResult != null)
            {
                if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
                else
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
            }
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

        private void txtMa_bp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_bp.Text.Trim()))
            {
                this.lbltenbp.Text = "";
            }
            else
            {
                if (this.txtMa_bp.RowResult == null)
                    return;
                this.lbltenbp.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_bp.RowResult["ten_bp2"].ToString() : this.txtMa_bp.RowResult["ten_bp"].ToString();
            }
        }
    }
}
