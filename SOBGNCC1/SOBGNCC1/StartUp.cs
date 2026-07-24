using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows;

namespace SOBGNCC1
{
    public class StartUp : StartUpTrans
    {
        public static string M_User_name = string.Empty;
        public static string titleWindow = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static int M_so_lien = 0;
        public static int M_loc_nsd = 0;
        public static int M_MA_MS = 0;
        public static string M_CHK_DATE_YN = "";
        public static string M_CHK_TON_VT = string.Empty;
        public static int so_dong_in = 0;
        public static string M_BP_BH = string.Empty;
        public static string tableList = "v_PH111;v_CT111";
        public static int M_AR_TT = 1;
        public static string stringBrowse1 = "";
        public static string stringBrowse2 = "";
        public static string stringBrowse3 = "";
        public static string stringBrowse4 = "";
        public static string storeproc = "";
        public static bool HiddenFieldIsSetted = false;
        public static bool isOk = false;
        public static DataSet HDBData = null;
        private static string[] hiddenCkFields = new string[6]
        {
      "t_ck_nt",
      "t_ck",
      "tl_ck",
      "ck_nt",
      "tk_ck",
      "ck"
        };
        public static SqlCommand TransFilterCmd;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DateTime M_ngay_ct0;
        public static int M_AR_CK;
        public override void Run()
        {
            StartupBase.Namespace = "SOBGNCC1";
            try
            {
                StartUpTrans.Ma_ct = "BGC";
                StartUpTrans.filterId = "SOBGNCC1";
                FrmPoctpna frmPoctpna = new FrmPoctpna();
                StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                StartUp.M_User_name = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
                StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
                StartUpTrans.M_ROUND_GIA = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
                StartUpTrans.M_ROUND_GIA_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
                StartUpTrans.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                StartUpTrans.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
                StartUp.M_MA_MS = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_MA_MS").ToString());
                StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                if (StartUpTrans.CommandInfo == null)
                {
                    int num1 = (int)ExMessageBox.Show(2360, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else if (StartUpTrans.DmctInfo == null)
                {
                    int num2 = (int)ExMessageBox.Show(2365, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm dịnh sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    string[] strArray;
                    if (StartUpTrans.M_LAN.Equals("V"))
                        strArray = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                    else
                        strArray = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                    if (strArray != null)
                    {
                        StartUp.stringBrowse1 = strArray[0];
                        StartUp.stringBrowse2 = strArray[1];
                    }
                    if (StartUpTrans.M_LAN.Equals("V"))
                        strArray = StartUpTrans.CommandInfo["Vbrowse1"].ToString().Split('|');
                    else
                        strArray = StartUpTrans.CommandInfo["Ebrowse1"].ToString().Split('|');
                    if (strArray != null)
                    {
                        StartUp.stringBrowse3 = strArray[0];
                        StartUp.stringBrowse4 = strArray[1];
                    }
                    StartUp.storeproc = StartUpTrans.CommandInfo["store_proc"].ToString();
                    StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
                    StartUp.M_BP_BH = StartUpTrans.DmctInfo["m_bp_bh"].ToString();
                    StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
                    StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                    StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                    StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                    int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                    int.TryParse(StartUpTrans.DmctInfo["so_lien"].ToString(), out StartUp.M_so_lien);
                    int.TryParse(StartUpTrans.DmctInfo["m_loc_nsd"].ToString(), out StartUp.M_loc_nsd);
                    int.TryParse(StartUpTrans.DmctInfo["so_dong_in"].ToString(), out StartUp.so_dong_in);
                    string str = " AND 1=1";
                    if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                        str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                    StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                    if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                    {
                        SelectTime selectTime = new SelectTime();
                        selectTime.LanguageID = "SOBGNCC1_4";
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
                    StartUpTrans.DsTrans = DataProvider.FillCommand(StartupBase.SasObj, StartUp.TransFilterCmd);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ct asc, so_ct asc";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0 ASC";
                    DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row["stt_rec"] = (object)string.Empty;
                    row["ma_nt"] = (object)StartUpTrans.M_ma_nt0;
                    StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row, 0);
                    StartUpTrans.tbStatus = DataProvider.FillCommand(StartupBase.SasObj, new SqlCommand("Select * from dmPost where ma_ct like '%" + StartUpTrans.Ma_ct + "%'")).Tables[0];
                    frmPoctpna.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
                    SysFunc.LoadIcon((Window)frmPoctpna);
                    frmPoctpna.StartUpMain = (StartUpTrans)this;
                    frmPoctpna.ShowInTaskbar = true;
                    frmPoctpna.ShowDialog();
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

        public static int UpdateRates(string _ma_nt, DateTime _ngay, Decimal _ty_gia)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("exec [dbo].[SetRates] @ma_nt,@ngay_ct,@ty_gia,@user_id");
                sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 3).Value = (object)_ma_nt;
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.VarChar, 8).Value = (object)string.Format("{0:yyyyMMdd}", (object)_ngay);
                sqlcmd.Parameters.Add("@ty_gia", SqlDbType.Decimal).Value = (object)_ty_gia;
                sqlcmd.Parameters.Add("@user_id", SqlDbType.VarChar).Value = (object)StartUpTrans.M_User_Id;
                return StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return 0;
        }

        public static void DataFilter(string stt_rec)
        {
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
        }

        public static string GetLanguageString(string code, string language)
        {
            return code == "M_MA_NT" ? (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0) ? StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString() : "") : (code == "M_MA_NT0" ? StartUpTrans.M_ma_nt0 : code);
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

        public static DataSet CheckData()
        {
            string format = "exec [dbo].{0} @status, @stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[SOBGNCC1-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static string EditCkFields(string values)
        {
            return string.Join(";", ((IEnumerable<string>)((IEnumerable<string>)values.Split(';')).ToArray<string>()).Select<string, string>((Func<string, string>)(x =>
            {
                if (x == "")
                    return x;
                string[] properties = x.Split(':');
                return ((IEnumerable<string>)StartUp.hiddenCkFields).Any<string>((Func<string, bool>)(a => a == properties[0])) ? x + ":IV" : x;
            })).ToArray<string>());
        }
    }
}
