using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;

namespace Inctpxd
{
    public class StartUp : StartUpTrans
    {
        public static string M_Tilte = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static string stringBrowse1 = "";
        public static string stringBrowse2 = "";
        public static string tableList = "v_PH84;v_CT84";
        public static string M_CHK_DATE_YN = "";
        public static int M_ROUND_SL = 0;
        public static string M_SL0_NTXT_CK = "0";
        public static SqlCommand TransFilterCmd;
        public static string M_CHK_TON_VT;
        public static string M_MA_THUE;
        public static string M_PHONE;
        private FrmInctpxd _Form;
        public static DateTime M_ngay_ct0;
        public static DateTime ngay_gia_px;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static string M_QL_LO_CK = "0";

        public override void Run()
        {
            StartupBase.Namespace = "Inctpxd";
            StartUpTrans.Ma_ct = "PXD";
            StartUpTrans.filterId = "INCTPXD";
            StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
            StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
            StartUpTrans.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
            StartUpTrans.M_ROUND = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND"));
            StartUpTrans.M_ROUND_NT = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_NT"));
            StartUpTrans.M_ROUND_GIA = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
            StartUpTrans.M_ROUND_GIA_NT = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
            StartUp.M_ROUND_SL = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_SL"));
            StartUp.M_SL0_NTXT_CK = StartupBase.SasObj.GetSysvar("M_SL0_NTXT_CK").ToString().Trim();
            StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
            StartUp.M_CHK_TON_VT = StartupBase.SasObj.GetOption("M_CHK_TON_VT").ToString();
            StartUp.M_MA_THUE = StartupBase.SasObj.GetOption("M_MA_THUE").ToString();
            StartUp.M_PHONE = StartupBase.SasObj.GetOption("M_PHONE").ToString();
            StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
            StartUp.ngay_gia_px = DateTime.Parse(StartupBase.SasObj.GetOption("M_NGAY_GIA_PX").ToString());
            StartUp.M_QL_LO_CK = StartupBase.SasObj.GetOption("M_QL_LO_CK").ToString().Trim();
            StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
            StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
            StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
            StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
            StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
            StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
            StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
            int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
            this._Form = new FrmInctpxd();
            if (StartUpTrans.CommandInfo == null)
            {
                int num = (int)ExMessageBox.Show(1115, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else if (StartUpTrans.DmctInfo == null)
            {
                int num = (int)ExMessageBox.Show(1120, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else
            {
                if (!StartUpTrans.CommandInfo["ma_phan_he"].ToString().Trim().Equals("IN"))
                    Application.Current.Shutdown();
                string[] strArray1;
                if (!StartUpTrans.M_LAN.Equals("V"))
                    strArray1 = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                else
                    strArray1 = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                string[] strArray2 = strArray1;
                if (strArray2 != null)
                {
                    StartUp.stringBrowse1 = strArray2[0];
                    StartUp.stringBrowse2 = strArray2[1];
                }
                StartUp.M_Tilte = StartUpTrans.M_LAN.Equals("V") ? SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString()) : SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString());
                string str = " AND 1=1";
                if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                    str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                {
                    SelectTime selectTime = new SelectTime();
                    selectTime.LanguageID = "INCTPXDSelTime";
                    SysFunc.LoadIcon((Window)selectTime);
                    selectTime.ShowDialog();
                    if (selectTime.IsOK)
                    {
                        StartUp.M_NGAY_BAT_DAU = new DateTime?((DateTime)selectTime.M_NGAY_CT1);
                        StartUp.M_NGAY_KET_THUC = new DateTime?((DateTime)selectTime.M_NGAY_CT2);
                        str += string.Format(" AND ngay_ct BETWEEN '{0:yyyyMMdd}' AND '{1:yyyyMMdd}'", (object)StartUp.M_NGAY_BAT_DAU, (object)StartUp.M_NGAY_KET_THUC);
                    }
                }
                string format = "exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter,@sl_ct;";
                StartUp.TransFilterCmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 0 ? string.Format(format, (object)"[LoadVoucher]") : string.Format(format, (object)StartUpTrans.Process_Store[0]));
                StartUp.TransFilterCmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                StartUp.TransFilterCmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)((StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) ? "1=1 AND ma_dvcs = '" + StartupBase.SasObj.M_ma_dvcs.Trim() + "'" : "stt_rec = '" + StartUpTrans.Editing_Stt_Rec + "'") + str);
                StartUp.TransFilterCmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                StartUp.TransFilterCmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                StartUp.TransFilterCmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)(StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) ? StartUpTrans.M_sl_ct0 : 1);
                StartUpTrans.DsTrans = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ct asc, so_ct asc";
                StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)string.Empty;
                row["ma_nt"] = (object)StartUpTrans.M_ma_nt0;
                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row, 0);
                StartUpTrans.tbStatus = StartupBase.SasObj.GetPostInfo(StartUpTrans.Ma_ct);
                this._Form.Title = StartUp.M_Tilte;
                SysFunc.LoadIcon((Window)this._Form);
                this._Form.StartUpMain = (StartUpTrans)this;
                this._Form.ShowDialog();
            }
        }

        public static Decimal GetRates(string _ma_nt, DateTime _ngay)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("select [dbo].[GetRates](@ma_nt,@ngay_ct)");
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 3).Value = (object)_ma_nt;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.VarChar, 8).Value = (object)string.Format("{0:yyyyMMdd}", (object)_ngay);
                object obj = StartupBase.SasObj.ExcuteScalar(sqlcmd);
                return Convert.ToDecimal(obj.Equals((object)DBNull.Value) ? (object)0 : obj);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return new Decimal(0);
        }

        public static void DataFilter(string stt_rec)
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
        }

        public static string GetLanguageString(string code, string language)
        {
            return code == "M_MA_NT" ? (StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0) ? "" : StartUpTrans.DsTrans.Tables[0].Rows[FrmInctpxd.iRow]["ma_nt"].ToString()) : (code == "M_MA_NT0" ? StartUpTrans.M_ma_nt0 : code);
        }

        public static DataSet CheckData()
        {
            string format = "exec [dbo].{0} @status, @stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[INCTPXD-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static DataTable GetDmnt()
        {
            SqlCommand sqlcmd = new SqlCommand("Select *,@ma_nt as ma_nt0, @Type as read_num_type from dmnt");
            sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 3).Value = (object)StartUpTrans.M_ma_nt0;
            sqlcmd.Parameters.Add("@Type", SqlDbType.Char, 1).Value = StartupBase.SasObj.GetOption("M_READ_NUM");
            DataTable dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
            dataTable.TableName = "TableNTInfo";
            if (StartUpTrans.DsTrans.Tables.IndexOf("TableNTInfo") >= 0)
                StartUpTrans.DsTrans.Tables.Remove("TableNTInfo");
            return dataTable;
        }

        public static bool IsTkMe(string tk)
        {
            bool flag = false;
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec CheckIsTkMe @tk");
                sqlcmd.Parameters.Add("@tk", SqlDbType.Char).Value = (object)tk;
                if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd) > 0)
                    flag = true;
            }
            catch (SqlException ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return flag;
        }

        public static DataTable GetPN(string ma_vt, string ma_kho, object ngay_ct)
        {
            DataTable dataTable = new DataTable();
            try
            {
                SqlCommand sqlcmd = new SqlCommand("Exec GetPN @ma_vt,@ma_kho, @ngay_ct");
                sqlcmd.Parameters.Add("@ma_vt", SqlDbType.VarChar).Value = (object)ma_vt;
                sqlcmd.Parameters.Add("@ma_kho", SqlDbType.VarChar).Value = (object)ma_kho;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.DateTime).Value = ngay_ct;
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                if (dataSet != null || dataSet.Tables.Count > 0)
                    dataTable = dataSet.Tables[0].Copy();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return dataTable;
        }
    }
}
