using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace Incd1
{
    public partial class FrmLoc : FormFilter
    {
        public FrmLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            SysFunc.LoadIcon((Window)this);
            this.txtVtTonkho.Text = StartUp.parameters[0];
            this.txtCt_dc.Text = StartUp.parameters[1];
            this.txtMa_kho.SearchInit();
            this.txtMaterial.SearchInit();
            this.txtten_kho.Text = this.txtMa_kho.RowResult == null ? "" : (StartupBase.M_LAN.Equals("V") ? this.txtMa_kho.RowResult["ten_kho"].ToString() : this.txtMa_kho.RowResult["ten_kho2"].ToString());
            this.txtTenVT.Text = this.txtMaterial.RowResult == null ? "" : (StartupBase.M_LAN.Equals("V") ? this.txtMaterial.RowResult["ten_vt"].ToString() : this.txtMaterial.RowResult["ten_vt2"].ToString());
            this.TxtStartDateTime.Focus();
            this.GridSearch.SasObj = StartupBase.SasObj;
            this.GridSearch.tableList = StartUp.TableList;
        }

        public string GetFilter(out string filterSD)
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "SELECT link from dmlink where rtrim(ma_link) like 'code'";
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            string link = dataSet.Tables[0].Rows[0]["link"].ToString().Trim();
            string str = "1=1";
            filterSD = "1=1";
            if (!string.IsNullOrEmpty(this.txtMa_kho.Text))
            {
                str = str + "  AND ma_kho LIKE '" + this.txtMa_kho.Text + "%'";
                ref string local = ref filterSD;
                local = local + "  AND ma_kho LIKE '" + this.txtMa_kho.Text + "%'";
            }
            if (!string.IsNullOrEmpty(this.txtMaterial.Text))
            {
                str = str + "  AND ma_vt LIKE '" + this.txtMaterial.Text + "%'";
                ref string local = ref filterSD;
                local = local + "  AND ma_vt LIKE '" + this.txtMaterial.Text + "%'";
            }
            if (!string.IsNullOrEmpty(this.txtMaDVCS.Text))
            {
                //str = str + " and ma_dvcs like  '" + this.txtMaDVCS.Text + "%'";
                //ref string local = ref filterSD;
                //local = local + " and ma_kho in (Select ma_kho From SQL04.SISERP2022_NITORI_QLKHO.dbo.dmkho Where ma_dvcs like '" + this.txtMaDVCS.Text + "%') ";
                str = str + " and ma_dvcs like  '" + "NITORI" + "%'";
                ref string local = ref filterSD;
                local = local + " and ma_kho in (Select ma_kho From "+ link+".dbo.dmkho Where ma_dvcs like '" + "NITORI" + "%') ";
            }
            this.GridSearch._GenerateSQLString();
            if (this.GridSearch.arrStrFilter != null && !string.IsNullOrEmpty(this.GridSearch.arrStrFilter[0]))
            {
                str = str + " and " + this.GridSearch.arrStrFilter[0];
                ref string local = ref filterSD;
                local = local + " and " + this.GridSearch.arrStrFilter[0];
            }
            return str;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            try
            {
                int.TryParse(this.txtCt_dc.Text, out StartUp.m_Tinh_dc);
                if (!this.TxtStartDateTime.IsValueValid)
                {
                    int num = (int)ExMessageBox.Show(355, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.TxtStartDateTime.Focus();
                    this.TxtStartDateTime.SelectAll();
                }
                else if (!this.TxtEndDateTime.IsValueValid)
                {
                    int num = (int)ExMessageBox.Show(360, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.TxtEndDateTime.Focus();
                    this.TxtEndDateTime.SelectAll();
                }
                else if (this.TxtStartDateTime.Value == null || this.TxtStartDateTime.Value == DBNull.Value)
                {
                    int num = (int)ExMessageBox.Show(365, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.TxtStartDateTime.Focus();
                    this.TxtStartDateTime.SelectAll();
                }
                else if (this.TxtEndDateTime.Value == null || this.TxtEndDateTime.Value == DBNull.Value)
                {
                    int num = (int)ExMessageBox.Show(370, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.TxtEndDateTime.Focus();
                    this.TxtEndDateTime.SelectAll();
                }
                else if ((DateTime)this.TxtStartDateTime.Value > (DateTime)this.TxtEndDateTime.Value)
                {
                    int num = (int)ExMessageBox.Show(380, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.TxtEndDateTime.Focus();
                    this.TxtEndDateTime.SelectAll();
                }
                else
                {
                    StartUp.DataSourceReport.Tables.Clear();
                    DataTable table = new DataTable("tbInfo");
                    table.Columns.Add("StartDateTime");
                    table.Columns.Add("EndDateTime");
                    table.Columns.Add("Ma_nt0");
                    table.Columns.Add("Ma_Kho");
                    table.Columns.Add("Ten_Kho");
                    table.Columns.Add("tk_vt_dmvt", typeof(string));
                    table.Columns.Add("ten_tk_vt_dmvt", typeof(string));
                    table.Columns.Add("ten_tk_vt_dmvt2", typeof(string));
                    string str1 = "";
                    string str2 = "";
                    string str3 = "";
                    foreach (Record record in (IEnumerable<Record>)this.GridSearch.GrdSearch.Records)
                    {
                        DataRowView dataItem = (record as DataRecord).DataItem as DataRowView;
                        if (dataItem["column_id"].ToString().Trim() == "tk_vt_dmvt")
                            str1 = dataItem["giatri1"].ToString().Trim();
                    }
                    if (!string.IsNullOrEmpty(str1))
                    {
                        SqlCommand sqlcmd = new SqlCommand();
                        sqlcmd.CommandText = "SELECT ten_tk,ten_tk2 from dmtk where rtrim(tk) like '" + str1 + "'";
                        DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                        if (dataSet.Tables[0].Rows.Count == 1)
                        {
                            str2 = dataSet.Tables[0].Rows[0]["ten_tk"].ToString().Trim();
                            str3 = dataSet.Tables[0].Rows[0]["ten_tk2"].ToString().Trim();
                        }
                    }
                    table.Rows.Add((object)this.TxtStartDateTime.Text, (object)this.TxtEndDateTime.Text, (object)StartupBase.SasObj.GetOption("M_MA_NT0").ToString(), (object)this.txtMa_kho.Text.Trim().ToString(), (object)this.txtten_kho.Text.Trim().ToString(), (object)str1, (object)str2, (object)str3);
                    StartUp.DataSourceReport.Tables.Add(table);
                    string filterSD;
                    string filter = this.GetFilter(out filterSD);
                    this.Hide();
                    bool KReport = true;
                    if (this.cbmau_bc.Value.ToString().Equals("1"))
                        KReport = false;
                    StartUp.QueryData(true, this.TxtStartDateTime.Value, this.TxtEndDateTime.Value, filter, filterSD, this.txtVtTonkho.Text, KReport, 1);
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtma_kho_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_kho.Text.Trim() == "")
                this.txtCt_dc.Text = "0";
            else
                this.txtCt_dc.Text = "1";
            this.txtten_kho.Text = "";
            if (this.txtMa_kho.RowResult == null)
                return;
            this.txtten_kho.Text = StartupBase.M_LAN.Equals("V") ? this.txtMa_kho.RowResult["ten_kho"].ToString() : this.txtMa_kho.RowResult["ten_kho2"].ToString();
        }

        private void txtMaterial_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtTenVT.Text = "";
            if (this.txtMaterial.RowResult == null)
                return;
            this.txtTenVT.Text = StartupBase.M_LAN.Equals("V") ? this.txtMaterial.RowResult["ten_vt"].ToString() : this.txtMaterial.RowResult["ten_vt2"].ToString();
        }

        private void txtVtTonkho_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtVtTonkho.Text))
                return;
            this.txtVtTonkho.Value = (object)"*";
        }

    }
}
