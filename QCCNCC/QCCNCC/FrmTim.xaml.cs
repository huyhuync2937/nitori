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
using System.Data;
using System.Windows;
using System.Windows.Input;

namespace QCCNCC
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
            else
            {
                if (this.txtMa_dvcs.RowResult != null)
                    this.tblten_dvcs.Text = this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString();

            }
            if (!string.IsNullOrEmpty(StartUp.Ma_ct))
                this.txtMa_qs.Filter = "ma_cts like '" + StartUp.Ma_ct + "%' and status =1";
            this.txtMa_qs.IsFocus = true;
        }

        private string GetPhFilterExpr()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = "";
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
                str = str + "  ngay_ct >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
                str = str + " and ngay_ct <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
                str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
                str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";

            if (!string.IsNullOrEmpty(this.txtMa_qs.Text))
                str = str + " and ma_qs LIKE '" + this.txtMa_qs.Text.Trim() + "%'";
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
                StartUp.TransFilterCmd.Parameters["@sl_ct"].Value = (object)-1;
                DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                string str1 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_so_luong"))).Sum().Value.ToString(StartupBase.SasObj.GetOption("M_IP_TIEN").ToString());
                string str2 = dataSet.Tables[0].AsEnumerable().Select<DataRow, Decimal?>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_so_luong1"))).Sum().Value.ToString(StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString());
                int count1 = dataSet.Tables[0].Rows.Count;
                if (count1 > 0)
                {
                    flag = true;
                    int num = (int)ExMessageBox.Show(2130, StartupBase.SasObj, string.Format("Có {0} chứng từ. Tổng số lượng  {1} / {2}", count1, str2, str1), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    int num1 = (int)ExMessageBox.Show(2345, StartupBase.SasObj, "Không có chứng từ nào như vậy! ", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                if (flag)
                {
                    if (StartUp.M_AR_CK == 0 && !StartUp.HiddenFieldIsSetted)
                    {
                        StartUp.stringBrowse1 = StartUp.EditCkFields(StartUp.stringBrowse1);
                        StartUp.stringBrowse2 = StartUp.EditCkFields(StartUp.stringBrowse2);
                    }
                    StartUp.HiddenFieldIsSetted = true;
                    FormView formView = new FormView(this.SasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, StartUp.stringBrowse1, StartUp.stringBrowse2, "stt_rec");
                    formView.ListFieldSum = "t_so_luong;t_so_luong1";
                    string str3;
                    if (StartUpTrans.M_LAN.Equals("V"))
                    {
                        str3 = "Hop dong don ban hang. Ky " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                        if (!string.IsNullOrEmpty(this.tblten_dvcs.Text.Trim()))
                            str3 = str3 + " - Don vi: " + this.tblten_dvcs.Text.Trim();
                    }
                    else
                    {
                        str3 = "Sales contract. Period " + this.txtNgay_ct1.Text + " - " + this.txtNgay_ct2.Text;
                        if (!string.IsNullOrEmpty(this.tblten_dvcs.Text.Trim()))
                            str3 = str3 + " - Unit: " + this.tblten_dvcs.Text.Trim();
                    }
                    formView.frmBrw.Title = str3;
                    FreeCodeFieldLib.InitFreeCodeField(StartupBase.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
                    formView.frmBrw.LanguageID = "QCCNCC_5";
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
                        if (FrmPoctpna.iRow > dataSet.Tables[0].Rows.Count - 1)
                            FrmPoctpna.iRow = dataSet.Tables[0].Rows.Count - 1;
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["stt_rec"].ToString() + "'";
                    }
                    if (formView.DataGrid.ActiveRecord != null)
                    {
                        int dataItemIndex = (formView.DataGrid.ActiveRecord as DataRecord).DataItemIndex;
                        if (dataItemIndex >= 0)
                        {
                            string str4 = (formView.DataGrid.DataSource as DataView)[dataItemIndex]["stt_rec"].ToString();
                            FrmPoctpna.iRow = dataItemIndex + 1;
                            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + str4 + "'";
                            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + str4 + "'";
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
                int num = (int)ExMessageBox.Show(2350, StartupBase.SasObj, "Thưa ngài! Hãy nhập từ ngày!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Value.ToString()))
                return true;
            int num1 = (int)ExMessageBox.Show(2355, StartupBase.SasObj, "Thưa ngài! Hãy nhập đến ngày!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
