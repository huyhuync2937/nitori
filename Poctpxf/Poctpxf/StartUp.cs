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

namespace Poctpxf
{
    public class StartUp : StartUpTrans
    {
        public static ActionTask currActionTask = ActionTask.None;
        public static string titleWindow = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static string M_CHK_TON_VT = string.Empty;
        public static string M_ma_ms = string.Empty;
        public static string stringBrowse1 = "";
        public static string stringBrowse2 = "";
        public static string M_CHK_DATE_YN = "";
        public static int M_loc_nsd = 0;
        public static string M_SD_HDDT = string.Empty;
        public static string M_SD_TOKEN = string.Empty;
        public static SqlCommand TransFilterCmd;
        public static bool IsQLHD;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DataSet HDBData = null;

        public static DataRow[] KhoNG = null;


        public override void Run()
        {
            StartupBase.Namespace = "Poctpxf";
            StartUpTrans.Ma_ct = "PXF";
            StartUpTrans.filterId = "Poctpxf";
            StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
            StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
            StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            StartUpTrans.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
            StartUp.M_CHK_TON_VT = StartupBase.SasObj.GetOption("M_CHK_TON_VT").ToString();
            StartUp.M_ma_ms = StartupBase.SasObj.GetOption("M_MA_MS").ToString().Trim();
            StartUpTrans.M_ROUND = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND"));
            StartUpTrans.M_ROUND_NT = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_NT"));
            StartUpTrans.M_ROUND_GIA = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
            StartUpTrans.M_ROUND_GIA_NT = (int)Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
            StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
            StartUp.M_SD_HDDT = StartupBase.SasObj.GetOption("M_SD_HDDT").ToString().Trim();
            if (StartUp.M_SD_HDDT.Equals(string.Empty))
                StartUp.M_SD_HDDT = "0";
            StartUp.M_SD_TOKEN = StartupBase.SasObj.GetOption("M_SD_TOKEN").ToString().Trim();
            if (StartUp.M_SD_TOKEN.Equals(string.Empty))
                StartUp.M_SD_TOKEN = "0";
            SqlCommand sqlcmd = new SqlCommand("SELECT COUNT(*) FROM v_dmmauhd WHERE ma_ct_qs LIKE @ma_ct");
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)("%" + StartUpTrans.Ma_ct + "%");
            StartUp.IsQLHD = Convert.ToDecimal(StartupBase.SasObj.ExcuteScalar(sqlcmd)) != new Decimal(0);
            StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
            StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
            int.TryParse(StartUpTrans.DmctInfo["m_loc_nsd"].ToString(), out StartUp.M_loc_nsd);
            StartUpTrans.filterView = string.Format("{0};{1}", (object)StartUpTrans.DmctInfo["v_phdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctdbf"].ToString().Trim());
            if (StartUpTrans.CommandInfo == null)
            {
                int num = (int)ExMessageBox.Show(1185, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else if (StartUpTrans.DmctInfo == null)
            {
                int num = (int)ExMessageBox.Show(1190, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else
            {
                StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
                StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
                StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                FrmPoctpxf frmPoctpxf = new FrmPoctpxf();
                string str = " AND 1=1";
                if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                    str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                {
                    SelectTime selectTime = new SelectTime();
                    selectTime.LanguageID = "PoctpxfSelTime";
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
                DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                row["stt_rec"] = (object)string.Empty;
                row["ma_nt"] = StartUpTrans.DmctInfo["ma_nt"];
                StartUpTrans.DsTrans.Tables[0].Rows.InsertAt(row, 0);
                string[] strArray = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                if (StartUpTrans.M_LAN != "V")
                    strArray = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                if (strArray != null)
                {
                    StartUp.stringBrowse1 = strArray[0];
                    StartUp.stringBrowse2 = strArray[1];
                }
                StartUpTrans.tbStatus = DataProvider.FillCommand(StartupBase.SasObj, new SqlCommand("Select * from dmPost where ma_ct like '%" + StartUpTrans.Ma_ct + "%'")).Tables[0];
                SysFunc.LoadIcon((Window)frmPoctpxf);
                frmPoctpxf.StartUpMain = (StartUpTrans)this;
                frmPoctpxf.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString());
                if (StartUpTrans.M_LAN != "V")
                    frmPoctpxf.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString());
                frmPoctpxf.ShowDialog();
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
        }

        public static void refresh(string _stt_rec)
        {
            SqlCommand sqlcmd = new SqlCommand("SELECT mau_hddt,sd_hddt_yn,so_seri_hddt,so_ct_hddt,tinh_trang_hddt,sl_in,[status] FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec = '" + _stt_rec + "'");
            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            if (table.Rows.Count == 0)
                return;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["mau_hddt"] = (object)table.DefaultView[0]["mau_hddt"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sd_hddt_yn"] = (object)table.DefaultView[0]["sd_hddt_yn"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_seri_hddt"] = (object)table.DefaultView[0]["so_seri_hddt"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct_hddt"] = (object)table.DefaultView[0]["so_ct_hddt"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"] = (object)table.DefaultView[0]["tinh_trang_hddt"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)table.DefaultView[0]["sl_in"].ToString();
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"] = (object)table.DefaultView[0]["status"].ToString();
            StartUpTrans.DsTrans.Tables[0].AcceptChanges();
        }

        public static Decimal GetTonKho(
          string ma_kho,
          string ma_vt,
          string stt_rec,
          object ngay_ct)
        {
            Decimal num = new Decimal(0);
            try
            {
                string str = "ma_kho = '" + ma_kho + "'" + " AND ma_vt = '" + ma_vt + "'";
                SqlCommand sqlcmd = new SqlCommand("exec CheckTonXuatAm 1, @ngay_ct,@Stt_rec, @advance");
                sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.SmallDateTime).Value = ngay_ct;
                sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.Char).Value = (object)stt_rec;
                sqlcmd.Parameters.Add("@advance", SqlDbType.VarChar).Value = (object)str;
                DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
                if (table.Rows.Count > 0)
                    num = StartUp.ParseDecimal(table.Rows[0]["ton00"], new Decimal(0));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return num;
        }

        public static Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = defaultvalue;
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        public DateTime GetEndDate(int year, int month)
        {
            DateTime dateTime = new DateTime();
            switch (month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    dateTime = new DateTime(year, month, 31);
                    break;
                case 2:
                    dateTime = year % 4 != 0 || year % 100 != 0 ? new DateTime(year, month, 28) : new DateTime(year, month, 29);
                    break;
                case 4:
                case 6:
                case 9:
                case 11:
                    dateTime = new DateTime(year, month, 30);
                    break;
            }
            return dateTime;
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

        public static DataTable GetGiaPN(
          string ngay_ct,
          string ma_vt,
          string ma_kho,
          string ma_kh)
        {
            string format = "Exec {0} @ngay_ct, @ma_vt, @ma_kho, @ma_kh;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 11 ? string.Format(format, (object)"[POCTPXF-PN]") : string.Format(format, (object)StartUpTrans.Process_Store[11]));
            sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.VarChar, 8).Value = (object)ngay_ct;
            sqlcmd.Parameters.Add("@ma_vt", SqlDbType.Char, 16).Value = (object)ma_vt;
            sqlcmd.Parameters.Add("@ma_kho", SqlDbType.Char, 8).Value = (object)ma_kho;
            sqlcmd.Parameters.Add("@ma_kh", SqlDbType.Char, 8).Value = (object)ma_kh;
            return StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
        }

        public static DataSet CheckData()
        {
            string format = "exec [dbo].{0} @status, @stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[POCTPXF-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
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
    }
}
