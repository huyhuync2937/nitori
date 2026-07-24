using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;

namespace COHDMTH43
{
    public class StartUp : StartupBase
    {
        public static DataSet dsReport = new DataSet();
        public static string hdTuNgay_ = "  -  -  ";
        public static string hdDenNgay_ = "  -  -  ";
        public static string ctTuNgay_ = "  -  -  ";
        public static string ctDenNgay_ = "  -  -  ";
        public static string tableList = "v_COHDMTH43";
        private static SqlCommand cmd = new SqlCommand();
        public static string g_strCondition = string.Empty;
        public static string g_strFilter = string.Empty;
        public static string g_ps_ck = "*";
        public static SasFormBrowes.FormBrowse oBrowse;
        public static SasFormBrowes.FormBrowse oBrowse_ct;
        public static FrmLoc _frmLoc;
        public static DataRow commandInfo;
        public static DateTime M_NGAY_KY1;
        public static DateTime M_NGAY_CT1;
        public static DateTime M_NGAY_CT2;
        public static string M_MA_DVCS;
        public static string M_IP_TIEN;
        public static string M_IP_TIEN_NT;
        public static string M_IP_SL;
        public static string M_ma_nt0;
        public static object g_hdTuNg;
        public static object g_hdDenNg;
        public static object g_ctTuNg;
        public static object g_ctDenNg;

        public override void Run()
        {
            StartupBase.Namespace = "COHDMTH43";
            try
            {
                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUp.M_IP_TIEN = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
                StartUp.M_IP_SL = StartupBase.SasObj.GetOption("M_IP_SL").ToString();
                StartUp.M_IP_TIEN_NT = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
                StartUp.M_NGAY_KY1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_NGAY_CT1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_CT1");
              //  StartUp.M_NGAY_CT2 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_CT2");
                StartUp.M_MA_DVCS = StartupBase.SasObj.DmdvcsInfo.Rows[0][0].ToString();
                StartUp.commandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                if (StartUp.commandInfo == null)
                    return;
                StartUp._frmLoc = new FrmLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void CallGridReport(
          bool isFirstLoad,
          object hdTuNg,
          object hdDenNg,
          object ctTuNg,
          object ctDenNg,
          string strCondition,
          string strFilter,
          int mau_bc,
          string ps_ck)
        {
            StartUp.g_hdTuNg = hdTuNg;
            StartUp.g_hdDenNg = hdDenNg;
            StartUp.g_ctTuNg = ctTuNg;
            StartUp.g_ctDenNg = ctDenNg;
            StartUp.g_strCondition = strCondition;
            StartUp.g_strFilter = strFilter;
            StartUp.g_ps_ck = ps_ck;
            if (isFirstLoad)
            {
                StartUp.cmd.CommandText = "Exec " + StartUp.commandInfo["store_proc"] + " @hdTuNg, @dhDenNg, @ctTuNg, @ctDenNg, @condition, @filter, @ps_ck,@filternt";
                StartUp.cmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdTuNg);
                StartUp.cmd.Parameters.Add("@dhDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdDenNg);
                StartUp.cmd.Parameters.Add("@ctTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(ctTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)ctTuNg);
                StartUp.cmd.Parameters.Add("@ctDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(ctDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)ctDenNg);
                StartUp.cmd.Parameters.Add("@condition", SqlDbType.VarChar, 4000).Value = (object)strCondition;
                StartUp.cmd.Parameters.Add("@filter", SqlDbType.VarChar, 4000).Value = (object)strFilter;
                StartUp.cmd.Parameters.Add("@ps_ck", SqlDbType.VarChar).Value = (object)ps_ck;
                StartUp.cmd.Parameters.Add("@filternt", SqlDbType.VarChar).Value = (object)mau_bc;
                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbDetail";
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, StartUp.fieldShow(mau_bc, 0));
                StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
                StartUp.oBrowse.F5 += new SasFormBrowes.FormBrowse.GridKeyUp_F5(StartUp.oBrowse_F5);
                StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
            }
            else
            {
                StartUp.dsReport.Tables.Clear();
                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbDetail";
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
                StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                StartUp.oBrowse.UpdateSumaryFields();
            }
            if (!isFirstLoad)
                return;
            StartUp.oBrowse.frmBrw.LanguageID = "COHDMTH43_1";
            StartUp.oBrowse.ShowDialog();
            StartUp._frmLoc.Close();
        }

        private static void oBrowse_F5(object sender, EventArgs e)
        {
            try
            {
                if (StartUp.oBrowse.ActiveRecord == null || StartUp.oBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
                    return;
                string str1 = string.Empty;
                string str2 = "";
                string empty = string.Empty;
                if (!(StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord is DataRecord activeRecord))
                    return;
                if (activeRecord != null)
                {
                    DataRowView dataItem = activeRecord.DataItem as DataRowView;
                    if (string.IsNullOrEmpty(dataItem["so_ct"].ToString().Trim()))
                    {
                        int index = activeRecord.Index;
                        str1 = StartUp.dsReport.Tables["tbDetail"].Rows[index]["ma_hdm"].ToString().Trim();
                        str2 = StartUp.dsReport.Tables["tbDetail"].Rows[index]["ma_vt"].ToString().Trim();
                    }
                    else
                    {
                        str1 = dataItem["so_ct"].ToString().Trim();
                        str2 = activeRecord.Cells["ma_vt"].Value.ToString().Trim();
                    }
                }
                DataSet dataSet1 = new DataSet();
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "Exec COHDMTH43_detail @ma_hdm, @ma_vt,@hdTuNg, @dhDenNg, @ctTuNg, @ctDenNg,  @condition, @filter";
                sqlcmd.Parameters.Add("@ma_hdm", SqlDbType.VarChar).Value = (object)str1;
                sqlcmd.Parameters.Add("@ma_vt", SqlDbType.VarChar).Value = (object)str2;
                sqlcmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartUp.g_hdTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartUp.g_hdTuNg);
                sqlcmd.Parameters.Add("@dhDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartUp.g_hdDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartUp.g_hdDenNg);
                sqlcmd.Parameters.Add("@ctTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartUp.g_ctTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartUp.g_ctTuNg);
                sqlcmd.Parameters.Add("@ctDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartUp.g_ctDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartUp.g_ctDenNg);
                sqlcmd.Parameters.Add("@condition", SqlDbType.VarChar, 4000).Value = (object)StartUp.g_strCondition;
                sqlcmd.Parameters.Add("@filter", SqlDbType.VarChar, 4000).Value = (object)StartUp.g_strFilter;
                DataSet dataSet2 = StartupBase.SasObj.ExcuteReader(sqlcmd);
                StartUp.oBrowse_ct = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataSet2.Tables[0].DefaultView, StartUp.fieldShow(StartUp._frmLoc.cbMau_bc.SelectedIndex, 1));
                StartUp.oBrowse_ct.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.oBrowse_ct.frmBrw.Title = (StartupBase.M_LAN.Equals("V") ? "Chi tiet thuc hien don hang: " : "Detailed performance of the order: ") + str1;
                StartUp.oBrowse_ct.frmBrw.LanguageID = "COHDMTH43_2";
                StartUp.oBrowse_ct.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static Decimal DecimalTryParse(object value)
        {
            Decimal result = new Decimal(0);
            if (value != null)
                Decimal.TryParse(value.ToString(), out result);
            return result;
        }

        public static void oBrowse_Esc(object sender, EventArgs e)
        {
        }

        public static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.commandInfo["rep_file"].ToString());
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.dsReport, "tbDetail");
            reportManager.Preview(StartUp.dsReport);
            SysFunc.ResetFilter(ref StartUp.dsReport, "tbDetail");
        }

        public static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_ctTuNg, StartUp.g_ctDenNg, StartUp.g_strCondition, StartUp.g_strFilter, StartUp._frmLoc.cbMau_bc.SelectedIndex, StartUp.g_ps_ck);
        }

        public static string fieldShow(int kindreport, int isDetail)
        {
            string empty = string.Empty;
            string[] strArray1 = new string[2];
            string str;
            switch (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper())
            {
                case "V":
                    string[] strArray2;
                    if (kindreport == 0)
                        strArray2 = StartUp.commandInfo["Vbrowse1"].ToString().Split('|');
                    else
                        strArray2 = StartUp.commandInfo["Vbrowse2"].ToString().Split('|');
                    str = isDetail != 0 ? strArray2[1] : strArray2[0];
                    break;
                default:
                    string[] strArray3;
                    if (kindreport == 0)
                        strArray3 = StartUp.commandInfo["Ebrowse1"].ToString().Split('|');
                    else
                        strArray3 = StartUp.commandInfo["Ebrowse2"].ToString().Split('|');
                    str = isDetail != 0 ? strArray3[1] : strArray3[0];
                    break;
            }
            return str;
        }

        public static DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "tableInfo";
            string str1 = "ĐƠN HÀNG TỪ NGÀY: " + StartUp.hdTuNgay_ + " ĐẾN NGÀY: " + StartUp.hdDenNgay_;
            string str2 = "PHÁT SINH TỪ NGÀY: " + StartUp.ctTuNgay_ + " ĐẾN NGÀY: " + StartUp.ctDenNgay_;
            string str3 = " ORDER DATE: " + StartUp.hdTuNgay_ + " TO DATE: " + StartUp.hdDenNgay_;
            string str4 = " ARISING FROM: " + StartUp.ctTuNgay_ + " TO DATE: " + StartUp.ctDenNgay_;
            dataTable.Columns.Add(new DataColumn("hdNgDenNg", typeof(string))
            {
                DefaultValue = (object)str1
            });
            dataTable.Columns.Add(new DataColumn("ctNgDenNg", typeof(string))
            {
                DefaultValue = (object)str2
            });
            dataTable.Columns.Add(new DataColumn("hdNgDenNg_Eng", typeof(string))
            {
                DefaultValue = (object)str3
            });
            dataTable.Columns.Add(new DataColumn("ctNgDenNg_Eng", typeof(string))
            {
                DefaultValue = (object)str4
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }
    }
}
