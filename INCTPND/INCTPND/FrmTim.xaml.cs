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
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using System.Linq;

namespace INCTPND
{
    public partial class FrmTim : FormFilter
    {
        public new static readonly DependencyProperty SasObjProperty = DependencyProperty.Register(nameof(SasObj), typeof(SasObject), typeof(FrmTim), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));

        public FrmTim()
        {
            this.InitializeComponent();
        }

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
        }

        private void FrmTim_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct1.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct1");
            this.txtNgay_ct2.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct2");
            this.txtUser.Text = StartUp.M_loc_nsd.ToString().Trim();
            this.txtMaDVCS.SearchInit();
            this.txtMa_kh.SearchInit();
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
                if (this.txtMaDVCS.RowResult != null)
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
                if (this.txtMa_kh.RowResult != null)
                    this.tblten_kh.Text = this.txtMa_kh.RowResult["ten_kh"].ToString();
                this.txtStatus.DisplayMemberPath = "ten_post";
            }
            else
            {
                if (this.txtMaDVCS.RowResult != null)
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
                if (this.txtMa_kh.RowResult != null)
                    this.tblten_kh.Text = this.txtMa_kh.RowResult["ten_kh2"].ToString();
                this.txtStatus.DisplayMemberPath = "ten_post2";
            }
            this.txtNgay_ct1.Focus();
        }

        private void FrmTim_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
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
                StartUp.TransFilterCmd.Parameters["@GtFilter"].Value = (object)"";
                StartUp.TransFilterCmd.Parameters["@sl_ct"].Value = (object)-1;
                DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                string str1 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tien"))).Sum().Value.ToString(StartupBase.SasObj.GetOption("M_IP_TIEN").ToString());
                string str2 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_tien_nt"))).Sum().Value.ToString(StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString());
                int count1 = dataSet.Tables[0].Rows.Count;
                if (count1 > 0)
                {
                    flag = true;
                    int num = (int)ExMessageBox.Show(580, StartupBase.SasObj, "Có [" + (object)count1 + "] chứng từ. Tổng phát sinh  [" + str2 + "] / [" + str1 + "]", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(585, StartupBase.SasObj, "Không có chứng từ nào như vậy! ", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                if (flag)
                {
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
                    FormView formView = new FormView(this.SasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, strBrowse, strBrowseCt, "stt_rec");
                    formView.ListFieldSum = "t_tien_nt;t_tien";
                    formView.frmBrw.Title = (StartUpTrans.M_LAN.Equals("V") ? "Phieu nhap kho. Ky " : "Receipt. Period ") + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                    FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
                    formView.frmBrw.LanguageID = "INCTPND_6";
                    formView.ShowDialog();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    int count2 = StartUpTrans.DsTrans.Tables[0].Rows.Count;
                    int count3 = StartUpTrans.DsTrans.Tables[1].Rows.Count;
                    for (int index = count2 - 1; index >= 1; --index)
                        StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(index);
                    for (int index = 0; index < count3; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.RemoveAt(0);
                    int count4 = dataSet.Tables[0].Rows.Count;
                    for (int index = 0; index < count4; ++index)
                        StartUpTrans.DsTrans.Tables[0].Rows.Add(dataSet.Tables[0].Rows[index].ItemArray);
                    int count5 = dataSet.Tables[1].Rows.Count;
                    for (int index = 0; index < count5; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        if (FrmINCTPND.iRow > dataSet.Tables[0].Rows.Count - 1)
                            FrmINCTPND.iRow = dataSet.Tables[0].Rows.Count - 1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["stt_rec"].ToString() + "'";
                    }
                    if (formView.DataGrid.ActiveRecord != null)
                    {
                        int dataItemIndex = (formView.DataGrid.ActiveRecord as DataRecord).DataItemIndex;
                        if (dataItemIndex >= 0)
                        {
                            string str3 = (formView.DataGrid.DataSource as DataView)[dataItemIndex]["stt_rec"].ToString();
                            FrmINCTPND.iRow = dataItemIndex + 1;
                            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
                            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str3 + "'";
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

        private void txtMa_kh_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMa_kh.RowResult == null)
                this.tblten_kh.Text = "";
            else
                this.tblten_kh.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMa_kh.RowResult["ten_kh"].ToString() : this.txtMa_kh.RowResult["ten_kh2"].ToString();
        }

        private void txtUser_TextChanged(object sender, RoutedPropertyChangedEventArgs<string> e)
        {
            if (!string.IsNullOrEmpty(this.txtUser.Text.ToString()))
                return;
            this.txtUser.Text = StartUp.M_loc_nsd.ToString().Trim();
        }

        private string GetPhFilterExpr()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = "";
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
                str = str + "  ngay_ct >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
                str = str + " and ngay_ct <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtMa_qs.Text))
                str = str + " and ma_qs LIKE '" + this.txtMa_qs.Text.Trim() + "%'";
            if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
                str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
                str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtMa_kh.Text))
                str = str + " and ma_kh = " + this.ConvertDataToSql((object)this.txtMa_kh.Text.Trim(), typeof(string));
            if (!string.IsNullOrEmpty(this.txtma_gd.Text))
                str = str + " and ma_gd = " + this.ConvertDataToSql((object)this.txtma_gd.Text.Trim(), typeof(string));
            if (this.txtStatus.Value != null && !string.IsNullOrEmpty(this.txtStatus.Value.ToString().Trim()))
                str = str + " and status = '" + this.txtStatus.Value.ToString().Trim() + "'";
            if (!string.IsNullOrEmpty(this.txtMaDVCS.Text))
                str = str + " and ma_dvcs LIKE '" + this.txtMaDVCS.Text.Trim() + "%'";
            if (Convert.ToInt16(this.txtUser.Value) == (short)1)
                str = str + " and [user_id] = " + (object)StartUpTrans.M_User_Id;
            if (!SysFunc.CheckPermission(this.SasObj, ActionTask.View, StartupBase.Menu_Id))
                str = str + " AND user_id0 = " + this.SasObj.UserInfo.Rows[0]["user_id"].ToString();
            if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
                str = str + " and " + this.GridSearch.arrStrFilter[0];
            return str;
        }

        private string GetCtFilterExpr()
        {
            string str = "1=1";
            if (!string.IsNullOrEmpty(this.txtTk_co.Text))
                str = str + " and ma_nx_i like " + this.ConvertDataToSql((object)(this.txtTk_co.Text.Trim() + "%"), typeof(string));
            if (!string.IsNullOrEmpty(this.txtMa_kho.Text.Trim()))
                str = str + " and ma_kho_i like " + this.ConvertDataToSql((object)(this.txtMa_kho.Text.Trim() + "%"), typeof(string));
            if (!string.IsNullOrEmpty(this.txtMa_vt.Text))
                str = str + " and ma_vt like " + this.ConvertDataToSql((object)(this.txtMa_vt.Text.Trim() + "%"), typeof(string));
            if (!string.IsNullOrEmpty(this.GridSearch.arrStrFilter[1]))
                str = str + " and " + this.GridSearch.arrStrFilter[1];
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

        private bool IsHopLe()
        {
            if (string.IsNullOrEmpty(this.txtNgay_ct1.Value.ToString()))
            {
                int num = (int)ExMessageBox.Show(590, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Value.ToString()) && !(this.txtNgay_ct1.dValue > this.txtNgay_ct2.dValue))
                return true;
            int num1 = (int)ExMessageBox.Show(595, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtNgay_ct2.Focus();
            this.txtNgay_ct2.SelectAll();
            return false;
        }

        private void txtMaDVCS_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMaDVCS.RowResult != null)
                this.lblTenDVCS.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtMaDVCS.RowResult["ten_dvcs"].ToString() : this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
            else
                this.lblTenDVCS.Text = "";
        }

        private void txtMa_gd_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtma_gd.RowResult != null)
                this.tblTen_gd.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtma_gd.RowResult["ten_gd"].ToString() : this.txtma_gd.RowResult["ten_gd2"].ToString();
            else
                this.tblTen_gd.Text = "";
        }

        private void txtMa_vt_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_vt.Text.Trim()))
            {
                this.tblten_vt.Text = "";
            }
            else
            {
                if (this.txtMa_vt.RowResult == null)
                    return;
                this.tblten_vt.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_vt.RowResult["ten_vt2"].ToString() : this.txtMa_vt.RowResult["ten_vt"].ToString();
            }
        }

        private void txtMa_kho_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_kho.Text.Trim()))
            {
                this.tblten_kho.Text = "";
            }
            else
            {
                if (this.txtMa_kho.RowResult == null)
                    return;
                this.tblten_kho.Text = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? this.txtMa_kho.RowResult["ten_kho2"].ToString() : this.txtMa_kho.RowResult["ten_kho"].ToString();
            }
        }

        private void txtTk_co_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtTk_co.RowResult == null)
                this.tblten_tk.Text = "";
            else
                this.tblten_tk.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtTk_co.RowResult["ten_tk"].ToString() : this.txtTk_co.RowResult["ten_tk2"].ToString();
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
