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

namespace APCTPN1
{
    public class StartUp : StartUpTrans
    {
        public static int M_loc_nsd = 0;
        public static string titleWindow = string.Empty;
        public static string M_ma_ms = string.Empty;
        public static int so_dong_in = 0;
        public static string M_CHK_DATE_YN = "";
        public static string M_MAU_THUE_CK = "0";
        public static SqlCommand TransFilterCmd;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static string M_User_name = String.Empty;
        public static string Tablename = "PH31";
        public override void Run()
        {
            StartupBase.Namespace = "APCTPN1";
            try
            {
                Debug.WriteLine(string.Format("#2: {0}", (object)DateTime.Now.ToString()));
                StartUpTrans.Ma_ct = "PN1";
                StartUpTrans.filterId = "APCTPN1";
                StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
                StartUp.M_ma_ms = StartupBase.SasObj.GetOption("M_MA_MS").ToString().Trim();
                StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                StartUp.M_User_name = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                StartUpTrans.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
                StartUpTrans.M_CHK_HD_VAO = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_CHK_HD_VAO").ToString());
                StartUp.M_MAU_THUE_CK = StartupBase.SasObj.GetOption("M_MAU_THUE_CK").ToString().Trim();
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                int.TryParse(StartUpTrans.DmctInfo["m_loc_nsd"].ToString(), out StartUp.M_loc_nsd);
                StartUpTrans.filterView = string.Format("{0};{1};{2}", (object)StartUpTrans.DmctInfo["v_phdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctgtdbf"].ToString().Trim());
                StartUp.so_dong_in = StartupBase.SasObj.GetDmctInfo(StartUpTrans.Ma_ct).Rows[0]["so_dong_in"] != DBNull.Value ? (int)Convert.ToInt16(StartupBase.SasObj.GetDmctInfo(StartUpTrans.Ma_ct).Rows[0]["so_dong_in"]) : 0;
                if (StartUpTrans.CommandInfo == null)
                {
                    int num = (int)ExMessageBox.Show(250, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                        return;
                    Application.Current.Shutdown();
                }
                else if (StartUpTrans.DmctInfo == null)
                {
                    int num = (int)ExMessageBox.Show((int)byte.MaxValue, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                        return;
                    Application.Current.Shutdown();
                }
                else
                {
                    FrmAPCTPN1 frmApctpN1 = new FrmAPCTPN1();
                    StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
                    StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
                    StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                    StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                    StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                    int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                    string str = " AND 1=1";
                    if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                        str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                    StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                    if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                    {
                        SelectTime selectTime = new SelectTime();
                        selectTime.LanguageID = "APCTPN1SelTime";
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
                    StartUp.TransFilterCmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"";
                    StartUp.TransFilterCmd.Parameters.Add("@sl_ct", SqlDbType.Int).Value = (object)(StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) ? StartUpTrans.M_sl_ct0 : 1);
                    Debug.WriteLine(string.Format("#3: {0}", (object)DateTime.Now.ToString()));
                    StartUpTrans.DsTrans = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ct asc, so_ct asc";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.Sort = "stt_rec0";
                    DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row["stt_rec"] = (object)string.Empty;
                    row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row, 0);
                    StartUpTrans.tbStatus = StartupBase.SasObj.GetPostInfo(StartUpTrans.Ma_ct);
                    frmApctpN1.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
                    Debug.WriteLine(string.Format("#4: {0}", (object)DateTime.Now.ToString()));
                    frmApctpN1.StartUpMain = (StartUpTrans)this;
                    SysFunc.LoadIcon((Window)frmApctpN1);
                    frmApctpN1.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void DeleteVoucher(string _stt_rec)
        {
            try
            {
                string format = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 11 ? string.Format(format, (object)"[DeleteVoucher]") : string.Format(format, (object)StartUpTrans.Process_Store[11]));
                sqlcmd.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static double GetRates(string _ma_nt, DateTime _ngay)
        {
            try
            {
                string format = "select [dbo].{0}(@ma_nt,@ngay_ct);";
                SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 12 ? string.Format(format, (object)"[GetRates]") : string.Format(format, (object)StartUpTrans.Process_Store[12]));
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 3).Value = (object)_ma_nt;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.VarChar, 8).Value = (object)string.Format("{0:yyyyMMdd}", (object)_ngay);
                object obj = StartupBase.SasObj.ExcuteScalar(sqlcmd);
                return Convert.ToDouble(obj.Equals((object)DBNull.Value) ? (object)0 : obj);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return 0.0;
        }

        public static DataSet CheckData()
        {
            string format = "exec [dbo].{0} @status, @stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[APCTPN1-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static string GetLanguageString(string code, string language)
        {
            return code == "M_MA_NT" ? (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0) ? StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString() : "") : (code == "M_MA_NT0" ? StartUpTrans.M_ma_nt0 : code);
        }

        public static DataSet GetPhieuht(string stt_rec)
        {
            SqlCommand sqlcmd = new SqlCommand("InPhieuht");
            sqlcmd.CommandType = CommandType.StoredProcedure;
            sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = (object)stt_rec;
            DataSet _ds = StartupBase.SasObj.ExcuteReader(sqlcmd);
            StartUp.GetDmnt(_ds);
            return _ds;
        }

        public static void GetDmnt(DataSet _ds)
        {
            SqlCommand sqlcmd = new SqlCommand("Select *,@ma_nt as ma_nt0, @Type as read_num_type from dmnt");
            sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 3).Value = (object)StartUpTrans.M_ma_nt0;
            sqlcmd.Parameters.Add("@Type", SqlDbType.Char, 1).Value = StartupBase.SasObj.GetOption("M_READ_NUM");
            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
            table.TableName = "TableNTInfo";
            if (_ds.Tables.IndexOf("TableNTInfo") >= 0)
                _ds.Tables.Remove("TableNTInfo");
            _ds.Tables.Add(table);
        }

        public static bool IsTkMe(string tk)
        {
            bool flag = false;
            try
            {
                string format = "exec {0} @tk;";
                SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 13 ? string.Format(format, (object)"[CheckIsTkMe]") : string.Format(format, (object)StartUpTrans.Process_Store[13]));
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

        public static bool CheckExistHDVao(
          string stt_rec,
          string so_ct0,
          string so_seri0,
          string ngay_ct0,
          string ma_so_thue)
        {
            string format = "Exec {0} @stt_rec, @so_ct0, @so_seri0, @ngay_ct0, @ma_so_thue;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 14 ? string.Format(format, (object)"[CheckExistsHDVao]") : string.Format(format, (object)StartUpTrans.Process_Store[14]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)stt_rec;
            sqlcmd.Parameters.Add("@so_ct0", SqlDbType.VarChar).Value = (object)so_ct0;
            sqlcmd.Parameters.Add("@so_seri0", SqlDbType.VarChar).Value = (object)so_seri0;
            sqlcmd.Parameters.Add("@ngay_ct0", SqlDbType.VarChar).Value = (object)ngay_ct0;
            sqlcmd.Parameters.Add("@ma_so_thue", SqlDbType.VarChar).Value = (object)ma_so_thue;
            return (int)StartupBase.SasObj.ExcuteScalar(sqlcmd) == 1;
        }
    }
}
