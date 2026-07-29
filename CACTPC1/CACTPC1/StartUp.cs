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

namespace CACTPC1
{
    public class StartUp : StartUpTrans 
    {
        public static ActionTask currActionTask = ActionTask.None;
        public static string titleWindow = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static string M_ma_ms = string.Empty;
        public static string[] M_Gd_2Tg_List = "4,5,6,7,8,9".Split(',');
        public static string M_CHK_DATE_YN = "";
        public static string M_IP_TIEN_NT = "#,0.00";
        public static string M_IP_TIEN = "#,0";
        public static Decimal so_ct0 = new Decimal(0);
        public static string dien_giai0 = "";
        public static string M_XL_CL_TGGS = "0";
        public static string M_CA_THEO_DOI_PT = "0";
        public static SqlCommand TransFilterCmd;
        public static int M_KT_CHI_TQ;
        public new static int M_CHK_HD_VAO;
        public static string M_MAU_THUE_CK = "0";
        public static int M_CHK_ZERO;
        public static string[] M_TK_TAI_QUY;
        public static string M_IP_TIEN_HD;
        public static string M_IP_TY_GIA;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DateTime M_ngay_ct0;
        public static DataTable dtRegInfo;        
        public static string M_SUA_MANT_PTC = "0";
        public static DataSet HDBData = null;
        public static bool isOk = false;

        public static DataRow[] PO = null;
        public override void Run()
        {
            StartupBase.Namespace = "CACTPC1";
            try
            {
                StartupBase.SasObj.SynchroFile(".", "CatgLib.dll");
                StartUpTrans.Ma_ct = "BN1"; 
                StartUpTrans.filterId = "CACTPC1";
                try
                {
                    StartUp.M_SUA_MANT_PTC = StartupBase.SasObj.GetOption("M_SUA_MANT_PTC").ToString().Trim();
                }
                catch
                {
                    StartUp.M_SUA_MANT_PTC = "0";
                }
                StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
                StartUpTrans.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                StartUp.M_ma_ms = StartupBase.SasObj.GetOption("M_MA_MS").ToString().Trim();
                StartUp.M_Gd_2Tg_List = StartupBase.SasObj.GetSysvar("M_GD_2TG_LIST").ToString().Trim().Split(',');
                StartUp.M_ngay_ct0 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_KT_CHI_TQ = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_KT_CHI_TQ").ToString());
                StartUp.M_CHK_HD_VAO = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_CHK_HD_VAO").ToString());
                StartUp.M_MAU_THUE_CK = StartupBase.SasObj.GetOption("M_MAU_THUE_CK").ToString().Trim();
                StartUp.M_CHK_ZERO = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_CHK_ZERO").ToString());
                StartUp.M_TK_TAI_QUY = StartupBase.SasObj.GetOption("M_TK_TAI_QUY").ToString().Trim().Split(',');
                StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                StartUpTrans.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
                StartUp.M_IP_TIEN_HD = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
                StartUp.M_IP_TY_GIA = StartupBase.SasObj.GetOption("M_IP_TY_GIA").ToString();
                StartUp.M_XL_CL_TGGS = StartupBase.SasObj.GetSysvar("M_XL_CL_TGGS").ToString().Trim();
                StartUp.M_CA_THEO_DOI_PT = StartupBase.SasObj.GetSysvar("M_CA_THEO_DOI_PT").ToString().Trim();
                StartUp.M_IP_TIEN_NT = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString() ?? "";
                StartUp.M_IP_TIEN = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() ?? "";
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                if (StartUpTrans.CommandInfo == null)
                {
                    int num = (int)ExMessageBox.Show(850, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                        return;
                    Application.Current.Shutdown();
                }
                else
                {
                    DataRow commandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartUpTrans.CommandInfo["menu_copy"].ToString().Trim());
                    if (commandInfo != null)
                    {
                        StartUpTrans.CommandInfo["vbrowse2"] = commandInfo["vbrowse2"];
                        StartUpTrans.CommandInfo["ebrowse2"] = commandInfo["ebrowse2"];
                    }
                    StartUpTrans.Ma_ct = StartUpTrans.CommandInfo["ma_ct"].ToString().Trim();
                    StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                    if (StartUpTrans.DmctInfo == null)
                    {
                        int num = (int)ExMessageBox.Show(855, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                            return;
                        Application.Current.Shutdown();
                    }
                    else
                    {
                        StartUpTrans.filterId = StartUpTrans.Ma_ct;
                        StartUpTrans.filterView = string.Format("{0};{1};{2}", (object)StartUpTrans.DmctInfo["v_phdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctgtdbf"].ToString().Trim());
                        StartUp.dtRegInfo = StartupBase.SasObj.GetRegInfo();
                        StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
                        StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
                        StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                        StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                        StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                        FrmCACTPC1 frmCactpC1 = new FrmCACTPC1();
                        int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                        string str = " AND 1=1";
                        if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                            str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                        StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                        if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                        {
                            SelectTime selectTime = new SelectTime();
                            selectTime.LanguageID = "CACTPC1SelTime";
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
                        StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                        StartUpTrans.DsTrans.Tables[2].DefaultView.Sort = "stt_rec0";
                        DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                        row["stt_rec"] = (object)string.Empty;
                        row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                        StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row, 0);
                        StartUpTrans.tbStatus = StartupBase.SasObj.GetPostInfo(StartUpTrans.Ma_ct);
                        SysFunc.LoadIcon((Window)frmCactpC1);
                        frmCactpC1.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
                        frmCactpC1.StartUpMain = (StartUpTrans)this;
                        frmCactpC1.ShowDialog();
                    }
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
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Trim().Equals("1"))
                {
                    string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString().Trim();
                    string format = "exec [dbo].[ApttpbDel] '{0}','{1}','{2}','{3}';\n";
                    string str2 = "";
                    foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[1].DefaultView)
                        str2 += string.Format(format, (object)_stt_rec, (object)StartUpTrans.Ma_ct, (object)dataRowView["tk_i"].ToString().Trim(), (object)str1);
                    Debug.WriteLine(str2);
                    SqlCommand sqlcmd = new SqlCommand(str2);
                    StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                }
                string format1 = "exec [dbo].{0} @cMa_ct,@stt_rec;";
                SqlCommand sqlcmd1 = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 11 ? string.Format(format1, (object)"[DeleteVoucher]") : string.Format(format1, (object)StartUpTrans.Process_Store[11]));
                sqlcmd1.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd1.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd1);
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
                SqlCommand sqlcmd = new SqlCommand("select [dbo].[GetRates](@ma_nt,@ngay_ct)");
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 7 ? string.Format(format, (object)"[CACTPC1-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[7]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static bool CheckQuy(string tk, DateTime ngay)
        {
            bool flag = true;
            try
            {
                SqlCommand sqlcmd = new SqlCommand("EXEC [dbo].[Caso1] @tk, @ngay_ct1,@ngay_ct2");
                sqlcmd.Parameters.Add("@tk", SqlDbType.VarChar, 50).Value = (object)tk;
                sqlcmd.Parameters.Add("@ngay_ct1", SqlDbType.SmallDateTime).Value = (object)StartUp.M_ngay_ct0;
                sqlcmd.Parameters.Add("@ngay_ct2", SqlDbType.SmallDateTime).Value = (object)ngay;
                DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[1];
                if (Convert.ToDecimal(table.Rows[0]["no_ck"]) - Convert.ToDecimal(table.Rows[0]["co_ck"]) < Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["t_tien"]))
                    flag = false;
            }
            catch (Exception ex)
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
            SqlCommand sqlcmd = new SqlCommand("Exec CheckExistsHDVao @stt_rec, @so_ct0, @so_seri0, @ngay_ct0, @ma_so_thue");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)stt_rec;
            sqlcmd.Parameters.Add("@so_ct0", SqlDbType.VarChar).Value = (object)so_ct0;
            sqlcmd.Parameters.Add("@so_seri0", SqlDbType.VarChar).Value = (object)so_seri0;
            sqlcmd.Parameters.Add("@ngay_ct0", SqlDbType.VarChar).Value = (object)ngay_ct0;
            sqlcmd.Parameters.Add("@ma_so_thue", SqlDbType.VarChar).Value = (object)ma_so_thue;
            return (int)StartupBase.SasObj.ExcuteScalar(sqlcmd) == 1;
        }

        public static void In()
        {
            try
            {
                (!(StartUpTrans.Ma_ct == "PC1") ? (Form)new FrmInBN1() : (Form)new FrmIn()).ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
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

        private static void oBrowse_F7(object sender, EventArgs e)
        {
            new FrmIn().ShowDialog();
        }

        private static void oBrowse_Esc(object sender, EventArgs e)
        {
            (sender as SasFormBrowes.FormBrowse).frmBrw.Close();
        }

        public static string GetLanguageString(string code, string language)
        {
            return code == "M_MA_NT" ? (!StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0) ? StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString() : "") : (code == "M_MA_NT0" ? StartUpTrans.M_ma_nt0 : code);
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

        public static Decimal ToDec(object value)
        {
            return value == DBNull.Value || value == null ? new Decimal(0) : (Decimal)value;
        }

        public static void InsertInfoBank(ref DataSet ds)
        {
            string str = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk"].ToString();
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "SELECT * FROM dmtknh WHERE tk LIKE '" + str + "'";
            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
            table.TableName = "tbInfoNH";
            if (table.Rows.Count == 0)
                table.Rows.Add(table.NewRow());
            ds.Tables.Add(table);
        }

        public static DataRow GetRateCL(string ma_nt, DateTime ngay_hd, DateTime ngay_ct)
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "EXEC [dbo].[CACTPT1-GetRateCL] @Ma_nt, @Ngay_hd, @Ngay_ct";
            sqlcmd.Parameters.Add("@Ma_nt", SqlDbType.VarChar).Value = (object)ma_nt;
            sqlcmd.Parameters.Add("@Ngay_hd", SqlDbType.SmallDateTime).Value = (object)ngay_hd;
            sqlcmd.Parameters.Add("Ngay_ct", SqlDbType.SmallDateTime).Value = (object)ngay_ct;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            return dataSet.Tables[0].Rows.Count == 0 ? (DataRow)null : dataSet.Tables[0].Rows[0];
        }
    }
}
