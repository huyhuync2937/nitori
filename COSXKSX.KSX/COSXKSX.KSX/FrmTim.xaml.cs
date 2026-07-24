using Infragistics.Windows.DataPresenter;
using SasControls;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using SasLib;
using System;
using System.Collections;
using System.Data;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Data.SqlClient;

namespace COSXKSX.KSX
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

        private void FrmTim_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void FrmTim_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct1.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct1");
            this.txtNgay_ct2.Value = (object)(DateTime)this.SasObj.GetSysvar("M_ngay_ct2");
            this.txtUser.Text = StartUp.M_loc_nsd.ToString().Trim();
            this.txtMa_dvcs.SearchInit();
            if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
            {
                if (this.txtMa_dvcs.RowResult != null)
                    this.tblten_dvcs.Text = this.txtMa_dvcs.RowResult["ten_dvcs"].ToString();
            }
            else if (this.txtMa_dvcs.RowResult != null)
                this.tblten_dvcs.Text = this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString();
            DataView defaultView = this.BindingSasObj.GetPostInfo(StartUpTrans.Ma_ct).DefaultView;
            DataRow row = defaultView.Table.NewRow();
            row["ten_post"] = (object)"Tất cả";
            row["ten_post2"] = (object)"All";
            defaultView.Table.Rows.Add(row);
            this.cb_status.ItemsSource = (IEnumerable)defaultView.Table.AsEnumerable().OrderBy<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("ma_post"))).AsDataView<DataRow>();
            this.cb_status.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.cb_status.SelectedIndex != -1)
                   return;
               this.cb_status.SelectedIndex = 0;
           }));
            this.cb_status.DisplayMemberPath = !StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V") ? "ten_post2" : "ten_post";
            this.txtMa_qs.IsFocus = true;
        }

        private string GetPhFilterExpr()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = " 1=1 ";
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
                str = str + " and ngay_ksx >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
                str = str + " and ngay_ksx <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
                str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
                str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (this.cb_status.Value != null && !string.IsNullOrEmpty(this.cb_status.Value.ToString().Trim()))
                str = str + " and status = '" + this.cb_status.Value.ToString().Trim() + "'";
            if (!string.IsNullOrEmpty(this.txtMa_dvcs.Text))
                str = str + " and ma_dvcs =  " + this.ConvertDataToSql((object)this.txtMa_dvcs.Text.Trim(), typeof(string));
            if (!SysFunc.CheckPermission(this.SasObj, ActionTask.View, StartupBase.Menu_Id))
                str = str + " and user_id0 = " + this.SasObj.UserInfo.Rows[0]["user_id"].ToString();
            if (Convert.ToInt16(this.txtUser.Value) == (short)1)
                str = str + " and [user_id] = " + (object)StartUpTrans.M_User_Id;
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
                string phfilter = this.GetPhFilterExpr();
                string ctfilter = this.GetCtFilterExpr();

                StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablePH + " where ma_ct = @ma_ct and " + phfilter + " order by ngay_ksx ");
                StartUp.TransFilterCmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                DataTable tblph = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                tblph.TableName = "tblPH1";
                StartUp.TransFilterCmd = new SqlCommand("select stt_rec as Column1,*,stt_rec as stt_rec1 from " + StartUp.v_tablesanpham + " where " + ctfilter);
                DataTable tblsanpham = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                tblsanpham.TableName = "tblsanpham1";
                StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablenguyenlieu + " where " + ctfilter);
                DataTable tblnguyenlieu = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                tblnguyenlieu.TableName = "tblnguyenlieu1";
                StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablemaymoc + " where " + ctfilter);
                DataTable tblmaymoc = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                tblmaymoc.TableName = "tblmaymoc1";
                StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablenguonluc + " where " + ctfilter);
                DataTable tblnguonluc = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                tblnguonluc.TableName = "tblnguonluc1";

                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(tblph.Copy());
                dataSet.Tables.Add(tblsanpham.Copy());
                dataSet.Tables.Add(tblnguyenlieu.Copy());
                dataSet.Tables.Add(tblmaymoc.Copy());
                dataSet.Tables.Add(tblnguonluc.Copy());

                dataSet.Tables[0].DefaultView.Sort = "ngay_ksx asc, so_ct asc";
                int count1 = dataSet.Tables[0].Rows.Count;
                if (count1 > 0)
                {
                    flag = true;
                    int num = (int)ExMessageBox.Show(2340, StartupBase.SasObj, "Có [" + (object)count1 + "] kế hoạch sản xuất.", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(2345, StartupBase.SasObj, "Không có kế hoạch sản xuất nào như vậy! ", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
                    formView.ListFieldSum = "t_tt_nt;t_tt";
                    string str1;
                    if (StartUpTrans.M_LAN.Equals("V"))
                    {
                        str1 = "Ke hoach san xuat. Ky " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                        if (!string.IsNullOrEmpty(this.tblten_dvcs.Text.Trim()))
                            str1 = str1 + " - Don vi: " + this.tblten_dvcs.Text.Trim();
                    }
                    else
                    {
                        str1 = "Manufacturing order. Period " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                        if (!string.IsNullOrEmpty(this.tblten_dvcs.Text.Trim()))
                            str1 = str1 + " - Unit: " + this.tblten_dvcs.Text.Trim();
                    }
                    formView.frmBrw.Title = str1;
                    FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
                    formView.frmBrw.LanguageID = "COSXLSX_KSX_5";
                    formView.ShowDialog();
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                    int count2 = StartUpTrans.DsTrans.Tables[0].Rows.Count;
                    int count3 = StartUpTrans.DsTrans.Tables[1].Rows.Count;
                    int count4 = StartUpTrans.DsTrans.Tables[2].Rows.Count;
                    int count5 = StartUpTrans.DsTrans.Tables[3].Rows.Count;
                    int count6 = StartUpTrans.DsTrans.Tables[4].Rows.Count;

                    for (int index = count2 - 1; index >= 1; --index)
                        StartUpTrans.DsTrans.Tables[0].Rows.RemoveAt(index);
                    for (int index = 0; index < count3; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.RemoveAt(0);
                    for (int index = 0; index < count4; ++index)
                        StartUpTrans.DsTrans.Tables[2].Rows.RemoveAt(0);
                    for (int index = 0; index < count5; ++index)
                        StartUpTrans.DsTrans.Tables[3].Rows.RemoveAt(0);
                    for (int index = 0; index < count6; ++index)
                        StartUpTrans.DsTrans.Tables[4].Rows.RemoveAt(0);

                    int count7 = dataSet.Tables[0].Rows.Count;
                    for (int index = 0; index < count7; ++index)
                        StartUpTrans.DsTrans.Tables[0].Rows.Add(dataSet.Tables[0].Rows[index].ItemArray);

                    int count8 = dataSet.Tables[1].Rows.Count;
                    for (int index = 0; index < count8; ++index)
                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);

                    int count9 = dataSet.Tables[2].Rows.Count;
                    for (int index = 0; index < count9; ++index)
                        StartUpTrans.DsTrans.Tables[2].Rows.Add(dataSet.Tables[2].Rows[index].ItemArray);

                    int count10 = dataSet.Tables[3].Rows.Count;
                    for (int index = 0; index < count10; ++index)
                        StartUpTrans.DsTrans.Tables[3].Rows.Add(dataSet.Tables[3].Rows[index].ItemArray);

                    int count11 = dataSet.Tables[4].Rows.Count;
                    for (int index = 0; index < count11; ++index)
                        StartUpTrans.DsTrans.Tables[4].Rows.Add(dataSet.Tables[4].Rows[index].ItemArray);

                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        if (FrmPoctpna.iRow > dataSet.Tables[0].Rows.Count - 1)
                            FrmPoctpna.iRow = dataSet.Tables[0].Rows.Count - 1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ksx asc, so_ct asc";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0 ASC";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.Sort = "stt_rec0 ASC";
                        StartUpTrans.DsTrans.Tables[3].DefaultView.Sort = "stt_rec0 ASC";
                        StartUpTrans.DsTrans.Tables[4].DefaultView.Sort = "stt_rec0 ASC";

                    }
                    if (formView.DataGrid.ActiveRecord != null)
                    {
                        int dataItemIndex = (formView.DataGrid.ActiveRecord as DataRecord).DataItemIndex;
                        if (dataItemIndex >= 0)
                        {
                            string str2 = (formView.DataGrid.DataSource as DataView)[dataItemIndex]["stt_rec"].ToString();
                            FrmPoctpna.iRow = dataItemIndex + 1;
                            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str2 + "'";
                            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str2 + "'";
                            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + str2 + "'";
                            StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + str2 + "'";
                            StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + str2 + "'";
                            StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ksx asc, so_ct asc";
                            StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0 ASC";
                            StartUpTrans.DsTrans.Tables[2].DefaultView.Sort = "stt_rec0 ASC";
                            StartUpTrans.DsTrans.Tables[3].DefaultView.Sort = "stt_rec0 ASC";
                            StartUpTrans.DsTrans.Tables[4].DefaultView.Sort = "stt_rec0 ASC";
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
            if (string.IsNullOrEmpty(this.txtNgay_ct1.Value.ToString()))
            {
                int num = (int)ExMessageBox.Show(2350, StartupBase.SasObj, "Thưa ngài! Hãy nhập từ ngày!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Value.ToString()))
                return true;
            int num1 = (int)ExMessageBox.Show(2355, StartupBase.SasObj, "Thưa ngài! Hãy nhập đến ngày!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtNgay_ct2.Focus();
            this.txtNgay_ct2.SelectAll();
            return false;
        }

        private void txtUser_TextChanged(object sender, RoutedPropertyChangedEventArgs<string> e)
        {
            if (!string.IsNullOrEmpty(this.txtUser.Text.ToString()))
                return;
            this.txtUser.Text = StartUp.M_loc_nsd.ToString().Trim();
        }

        private void txtMa_dvcs_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMa_dvcs.RowResult == null)
                this.tblten_dvcs.Text = "";
            else if (StartUpTrans.M_LAN.Equals("V"))
                this.tblten_dvcs.Text = this.txtMa_dvcs.RowResult["ten_dvcs"].ToString();
            else
                this.tblten_dvcs.Text = this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString();
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
