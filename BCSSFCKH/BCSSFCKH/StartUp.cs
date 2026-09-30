using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using Microsoft.SqlServer.Server;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace BCSSFCKH
{
    public class StartUp : StartupBase
    {
        public static DataSet dsReport = new DataSet();
        public static string hdTuNgay_ = "  -  -  ";
        public static string hdDenNgay_ = "  -  -  ";
        public static string ctTuNgay_ = "  -  -  ";
        public static string ctDenNgay_ = "  -  -  ";
        public static string tableList = "v_Incd1";
        private static SqlCommand cmd = new SqlCommand();
        public static int kindStyleReport = -1;
        public static string g_strCondition = string.Empty;
        public static string g_strFilter = string.Empty;
        public static string g_ps_ck = "*";
        public static SasFormBrowes.FormBrowse oBrowse;
        public static SasFormBrowes.FormBrowse oBrowse_ct;
        public static FrmLoc _frmLoc;
        public static DataRow commandInfo;
        public static DateTime M_NGAY_KY1;
        public static DateTime M_NGAY_CT1;
        public static string M_MA_DVCS;
        public static string M_IP_TIEN;
        public static string M_IP_TIEN_NT;
        public static string M_IP_SL;
        public static string M_ma_nt0;
        public static object g_hdTuNg;
        public static object g_hdDenNg;
        public static object g_ctTuNg;
        public static object g_ctDenNg;
        public static object g_loaiNvl;
        public static string[] CanChangeValueFields;
        public static int m_user_id = 0;

        public override void Run()
        {
            StartupBase.Namespace = "BCSSFCKH";
            try
            {

                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUp.M_IP_TIEN = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
                StartUp.M_IP_SL = StartupBase.SasObj.GetOption("M_IP_SL").ToString();
                StartUp.M_IP_TIEN_NT = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
                StartUp.M_NGAY_KY1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_NGAY_CT1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_CT1");
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
        public static string fieldShow(int kindReport, int isDetail)
        {
            string str = string.Empty;
            string[] strArray1 = new string[2];
            switch (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper())
            {
                case "V":
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray2 = StartUp.commandInfo["Vbrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray2[1] : strArray2[0];
                            break;
                        case 2:
                            string[] strArray3 = StartUp.commandInfo["Vbrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray3[1] : strArray3[0];
                            break;
                    }
                    break;
                default:
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray4 = StartUp.commandInfo["Ebrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray4[1] : strArray4[0];
                            break;
                        case 2:
                            string[] strArray5 = StartUp.commandInfo["Ebrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray5[1] : strArray5[0];
                            break;
                    }
                    break;
            }
            return str;
        }
        public static void CallGridReport( bool isFirstLoad,object hdTuNg,object hdDenNg,object txtMavt)
        {
            StartUp.g_hdTuNg = hdTuNg;
            StartUp.g_hdDenNg = hdDenNg;
            StartUp.g_loaiNvl = txtMavt;

            if (isFirstLoad)
            {
                StartUp.cmd.CommandText = "Exec " + StartUp.commandInfo["store_proc"] + " @hdTuNg, @dhDenNg,@ma_vt";
                StartUp.cmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdTuNg);
                StartUp.cmd.Parameters.Add("@dhDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdDenNg);
                StartUp.cmd.Parameters.Add("@ma_vt", SqlDbType.VarChar).Value = string.IsNullOrEmpty(txtMavt.ToString()) ? (object)"" : (string)txtMavt;

                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbtong";
                dataTable.TableName = "tbDetail";
                StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, StartUp.fieldShow(1, 0));
                //StartUp.oBrowse.F3 += new SasFormBrowes.FormBrowse.GridKeyUp_F3(StartUp.oBrowse_F3);
                StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);

                object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport");
                //if (name != null)
                //{
                //    ToolBar toolBar = name as ToolBar;
                //    //for (int i = toolBar.Items.Count - 1; i > 0; i--)
                //    //{
                //    //    if ((toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnRefresh" && (toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnExport")
                //    //    {
                //    //        toolBar.Items.Remove((toolBar.Items[i] as SasControls.ToolBarButton));
                //    //    }
                //    //}
                //    SasControls.ToolBarButton toolBarButton3 = new SasControls.ToolBarButton();
                //    toolBarButton3.BorderBrush = (Brush)Brushes.Transparent;
                //    toolBarButton3.Name = "btnXoa";
                //    toolBarButton3.Text = "Tạo lệnh sản xuất";
                //    toolBarButton3.ToolTip = "F2";
                //    toolBarButton3.ImagePath = "Images\\UpdateSearch.png";
                //    toolBarButton3.Click += new RoutedEventHandler(ToolBarButtonF2_Click);
                //    toolBar.Items.Insert(1, toolBarButton3);

                //}

                //StartUp.oBrowse.F5 += new SasFormBrowes.FormBrowse.GridKeyUp_F5(StartUp.oBrowse_F5);
                StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);

                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
                //StartUp.oBrowse.SetRowColorByTag("sx2", "0", Colors.Black, true);
                //StartUp.oBrowse.SetRowColorByTag("sx3", "1", Colors.Red, false);
            }
            else
            {
                StartUp.dsReport.Tables.Clear();
                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbDetail";
                StartUp.dsReport.Tables[0].TableName = "tbtong";
                dataTable.TableName = "tbDetail";
                StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
                StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                StartUp.oBrowse.UpdateSumaryFields();
            }
            if (!isFirstLoad)
                return;
            StartUp.oBrowse.frmBrw.LanguageID = "Incd1";
            StartUp.oBrowse.ShowDialog();
            StartUp._frmLoc.Close();
        }



        public static bool CheckValidSoct(SasObject SasObj, string ma_qs, string so_ct, string stt_rec)
        {
            string format = "{0}";
            SqlCommand sqlCommand = new SqlCommand(string.Format(format, "CheckValidSoct"));
            sqlCommand.Parameters.Add(new SqlParameter("@ma_qs", ma_qs));
            sqlCommand.Parameters.Add(new SqlParameter("@so_ct", so_ct));
            sqlCommand.Parameters.Add(new SqlParameter("@stt_rec", stt_rec));
            sqlCommand.CommandType = CommandType.StoredProcedure;
            return (bool)SasObj.ExcuteScalar(sqlCommand);
        }

        public static string GetNewSoct(SasObject SasObj, string ma_qs, bool isUpdateNewSoCt)
        {
            try
            {
                string cmdText;
                //if (isUpdateNewSoCt)
                //{
                //    string format = "EXEC  {0} '" + ma_qs.Trim() + "'";
                //    cmdText = string.Format(format, "[GetNewSoct#2]");
                //}
                //else if (Convert.ToInt16(SasObj.GetOption("M_AUTO_SOCT").ToString()) == 1)
                //{
                cmdText = "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'";
                string update = "UPDATE dmqs SET so_ct = so_ct + 1 WHERE ma_qs = '" + ma_qs.Trim() + "'";

                //}
                //else
                //{
                //    string text;
                //    string format2 = (text = "EXEC  {0} '" + ma_qs.Trim() + "'");
                //    cmdText = string.Format(format2, "GetNewSoct");
                //}

                StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(update));
                DataTable dataTable = SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0];
                if (dataTable.Rows.Count > 0)
                {
                    DataRow dataRow = dataTable.Rows[0];
                    if (dataRow[1] != null && dataRow[1] != DBNull.Value)
                    {
                        string value = dataRow[1].ToString();
                        return string.Format(dataRow[0].ToString(), Convert.ToDouble(value));
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }

            return "";
        }
        //private static void oBrowse_F3(object sender, EventArgs e)
        //{
        //    DataView dataView = StartUp.oBrowse.DataGrid.DataSource as DataView;

        //    DataTable distinctValues = dataView.ToTable(true, "chon");

        //    DataRowView[] array = (from DataRowView x in distinctValues.DefaultView
        //                           where (bool)x["chon"]
        //                           select x).ToArray();
        //    if (array.Length <= 0)
        //    {
        //        int num = (int)ExMessageBox.Show(110, StartupBase.SasObj, "Phải đánh dấu trước khi sửa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        //        return;
        //    }
        //    UpdateFieldHandler.Execute();
        //}
        public static void oBrowse_Esc(object sender, EventArgs e)
        {
        }

        public static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.commandInfo["rep_file"].ToString(), StartUp.kindStyleReport);
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.dsReport, "tbDetail");
            reportManager.Preview(StartUp.dsReport);
            SysFunc.ResetFilter(ref StartUp.dsReport, "tbDetail");
        }

        public static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            try
            {
                // Get all rows with id=1 and recalculate subsequent rows with the same ma_may
                
                bool kindReport = true;
            
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_loaiNvl);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

  
    }
}
