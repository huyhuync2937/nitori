using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;

namespace QCSN
{
    public class StartUp : StartupBase
    {
        public static DataSet DataSourceReport = new DataSet();
        public static string sqlTableView = "v_QCSN";
        public static string SqlTableKey = "stt_rec";
        public static string SqlTableObjectName = "ten_kh";
        private static SqlCommand cmd1 = new SqlCommand();
        private static SqlCommand cmd = new SqlCommand();
        public static string TableName = "ct70";
        public static DataTable dtInfo;
        private static SasFormBrowes.FormBrowse oBrowse;
        public static DataRow CommandInfo;
        public static DateTime M_ngay_ct0;
        public static string M_ma_nt0;
        private static FormLoc _frmLoc;

        public override void Run()
        {
            StartupBase.Namespace = "QCSN";
            try
            {
                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                DateTime now1 = DateTime.Now;
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.M_ngay_ct0 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.dtInfo = new DataTable();
                StartUp.dtInfo.TableName = "TbInfo";
                StartUp.dtInfo.Columns.Add("StartDate");
                StartUp.dtInfo.Columns.Add("EndDate");
                if (StartUp.CommandInfo == null)
                    return;
                StartUp._frmLoc = new FormLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString());
                if (StartupBase.M_LAN != "V")
                    StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
                DateTime now2 = DateTime.Now;
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        public static void CallGridVouchers(
          object StartDate,
          object EndDate,
         
          string MaVT)
        {
            try
            {
                string strBrowse1 = "";
                string strBrowseCt1 = "";
                string strBrowse2 = "";
                string strBrowseCt2 = "";
                string[] strArray1 = StartUp.CommandInfo["store_proc"].ToString().Split('|');
                StartUp.cmd = new SqlCommand();
                StartUp.cmd.CommandText = "Exec " + strArray1[0] + " @hdTuNg, @hdDenNg, @ma_vt";
                StartUp.cmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartDate.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartDate);
                StartUp.cmd.Parameters.Add("@hdDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(EndDate.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)EndDate);
                StartUp.cmd.Parameters.Add("@ma_vt", SqlDbType.NVarChar).Value = (object)MaVT;
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable1 = dataSet.Tables[0].Copy();
                dataTable1.TableName = "tbMain";
                DataTable dataTable2 = dataSet.Tables[1].Copy();
                dataTable2.TableName = "tbDetail";
                
                DataTable table1 = dataTable1.Copy();
                DataTable table2 = dataTable2.Copy();

                table1.TableName = "tbPh";
              
                table2.TableName = "tbCt";
           
                StartUp.DataSourceReport = new DataSet();
                StartUp.DataSourceReport.Tables.Add(table1);
                StartUp.DataSourceReport.Tables.Add(table2);
                StartUp.DataSourceReport.Tables.Add(StartUp.dtInfo.Copy());
         
                string[] strArray2 = StartUp.CommandInfo["VBrowse1"].ToString().Trim().Split('|');
                string[] strArray3 = StartUp.CommandInfo["VBrowse2"].ToString().Trim().Split('|');
                if (StartupBase.M_LAN != "V")
                {
                    strArray2 = StartUp.CommandInfo["EBrowse1"].ToString().Trim().Split('|');
                    strArray3 = StartUp.CommandInfo["EBrowse2"].ToString().Trim().Split('|');
                }
               
                strBrowse1 = dataTable2.Rows[0]["HeaderString"].ToString();
                strBrowseCt1 ="";
                strBrowse2 = strArray3[0];
                strBrowseCt2 = strArray3[1];
               
                StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable1.DefaultView,  strBrowse1);
                StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
                StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar"].ToString());
                StartUp.oBrowse.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(StartUp.oBrowse_Esc);
                StartUp.oBrowse.DataGrid.Loaded += new RoutedEventHandler(StartUp.DataGrid_Loaded);
                if (StartupBase.M_LAN != "V")
                    StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartUp.CommandInfo["bar2"].ToString());
                StartUp.oBrowse.frmBrw.LanguageID = "QCSNBrowse";
                StartUp.oBrowse.ShowDialog();
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.InnerException.Message);
            }
        }

        private static void DataGrid_Loaded(object sender, RoutedEventArgs e)
        {
            StartUp.oBrowse.DataGrid.Dispatcher.BeginInvoke((Delegate)new Action(() => StartUp.oBrowse.DataGrid.Focus()), DispatcherPriority.Background);
        }

        public static void QueryData(
          bool isFirstLoad,
          object StartDate,
          object EndDate,
          string MaVT
        )
        {
            try
            {
                if (isFirstLoad)
                {
                    StartUp.CallGridVouchers(StartDate, EndDate, MaVT);
                }
                else
                {
                    StartUp.DataSourceReport.Tables["tbInfo"].Copy();
                    StartUp.DataSourceReport = new DataSet();
                    DataSet dataSet = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                    DataTable dataTable = dataSet.Tables[0].Copy();
                    dataTable.TableName = "tbMain";
                    DataTable table1 = dataSet.Tables[1].Copy();
                    table1.TableName = "tbDetail";
                    DataTable table2 = dataTable.Copy();
                    DataTable table3 = table1.Copy();
                    table2.TableName = "tbPh";
                    table2.Columns.Add("KHInfo", typeof(string), "TRIM(ma_kh)+' - '+ TRIM(ten_kh)");
                    table3.TableName = "tbCt";
                    table3.Columns.Add("VTInfo", typeof(string), "TRIM(ma_vt)+' - '+ TRIM(ten_vt)");
                    StartUp.DataSourceReport = new DataSet();
                    StartUp.DataSourceReport.Tables.Add(table2);
                    StartUp.DataSourceReport.Tables.Add(table3);
                    StartUp.DataSourceReport.Tables.Add(StartUp.dtInfo.Copy());
                    DataRelation relation = new DataRelation("Stt_rec_Relation", table2.Columns["stt_rec"], table3.Columns["stt_rec"], false);
                    StartUp.DataSourceReport.Relations.Add(relation);
                    StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
                    StartUp.DataSourceReport.Tables.Add(table1);
                    StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                    StartUp.oBrowse.UpdateSumaryFields();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            int result;
            string filter = StartUp._frmLoc.GetFilter();
            StartUp.QueryData(false, StartUp._frmLoc.TxtStartDateTime.Value, StartUp._frmLoc.TxtEndDateTime.Value, StartUp._frmLoc.txtMaVT.Text.Trim());
        }

        private static void oBrowse_Esc(object sender, EventArgs e)
        {
            StartUp.oBrowse.frmBrw.Close();
            if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                return;
            Application.Current.Shutdown();
        }

        private static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.CommandInfo["rep_file"].ToString());
            SysFunc.DSCopyWithFilter((XamDataGrid)StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.DataSourceReport, "tbPh");
            reportManager.Preview(StartUp.DataSourceReport);
        }

        public static string GetTableShow(bool KindReport)
        {
            string empty = string.Empty;
            string upper = StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper();
            string str1 = "ngay_ct;Ma_ct0;so_ct;ma_kh;";
            string str2 = !upper.Equals("V") ? str1 + "ten_kh2;dien_giai;tk;tk_du;" : str1 + "ten_kh;dien_giai;tk;tk_du;";
            string str3 = !KindReport ? str2 + "ps_no_nt;ps_co_nt;ma_vv;ma_phi;" : str2 + "ps_no;ps_co;ma_vv;ma_phi;";
            return !upper.Equals("V") ? str3 + "ten_tk2;ten_tk2_du;ma_ct;ma_dvcs" : str3 + "ten_tk;ten_tk_du;ma_ct;ma_dvcs";
        }

        public static int GetCountDVCS()
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandType = CommandType.Text;
            sqlcmd.CommandText = "Select count(ma_dvcs) From dmdvcs";
            return int.Parse(StartupBase.SasObj.ExcuteScalar(sqlcmd).ToString());
        }
    }
}
