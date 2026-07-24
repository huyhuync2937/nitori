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
using System.Diagnostics;
using System.Linq;
using System.Windows;

namespace Socthda
{
    public class StartUp : StartUpTrans
    {
        public static string M_Tilte = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static int M_BP_BH = 0;
        public static int M_UPDATE_GIA2 = 0;
        public static string M_CHK_DATE_YN = "";
        public static string M_SD_HDDT = string.Empty;
        public static string M_SD_TOKEN = string.Empty;
        public static string stringBrowse1 = "";
        public static string stringBrowse2 = "";
        public static string stringBrowse3 = "";
        public static string stringBrowse4 = "";
        public static bool HiddenFieldIsSetted = false;
        private static string[] hiddenCkFields = new string[6]
        {
      "t_ck_nt",
      "t_ck",
      "tl_ck",
      "ck_nt",
      "tk_ck",
      "ck"
        };
        private static string[] hiddenKmFields = new string[2]
        {
      "km_ck",
      "tk_km_i"
        };

        private static string[] hiddenGiaVonFields = new string[8]
        {
      "gia",
      "gia_nt",
      "tien",
      "tien_nt",
      "t_tien",
      "t_tien_nt",
        "tk_vt",
        "tk_gv"
        };
        private static string[] hiddenGiaBanFields = new string[20]
        {
          "gia2",
          "gia_nt2",
          "tien2",
          "tien_nt2",
          "t_tien2",
          "t_tien_nt2",
           "tk_dt",
            "thue_suat",
            "t_thue_nt",
            "t_thue",
            "ma_nt",
            "ty_giaf",
            "t_thue_nt",
            "t_thue",
            "t_tt",
            "t_tt_nt",
            "tk_thue_co",
            "tk_thue_no",
            "ma_nx",
            "ma_thue"
        };

        public static SqlCommand TransFilterCmd;
        public static int M_KM_CK;
        public static int M_AR_CK;
        public static int M_THUE_KM_CK;
        public static string M_CHK_TON_VT;
        public static string M_MA_THUE;
        public static string M_PHONE;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        private FrmSocthda _Form;
        public static DateTime M_ngay_ct0;
        public static bool IsQLHD;
        public static DateTime ngay_gia_px;
        public static DataTable dtRegInfo;
        public static DataTable CTKMTable = null;
        public static DataTable VattuKMHTHTable = null;      
        public static string M_USER_NAME = string.Empty;
        public static string M_QL_LO_CK = "0";

        public override void Run()
        {
            StartupBase.Namespace = "Socthda";
            StartUpTrans.Ma_ct = "HDA";
            StartUpTrans.filterId = "SOCTHDA";
            StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString().Trim().ToUpper();
            StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
            StartUpTrans.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
            StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
            StartUp.M_USER_NAME = (StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString());
            StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
            StartUp.M_CHK_TON_VT = StartupBase.SasObj.GetOption("M_CHK_TON_VT").ToString();
            StartUp.M_MA_THUE = StartupBase.SasObj.GetOption("M_MA_THUE").ToString();
            StartUp.M_PHONE = StartupBase.SasObj.GetOption("M_PHONE").ToString();
            StartUp.M_QL_LO_CK = StartupBase.SasObj.GetOption("M_QL_LO_CK").ToString().Trim();
            StartUp.M_THUE_KM_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_THUE_KM_CK"));
            StartUp.M_UPDATE_GIA2 = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_UPDATE_GIA2"));
            StartUp.M_SD_HDDT = StartupBase.SasObj.GetOption("M_SD_HDDT").ToString().Trim();
            if (StartUp.M_SD_HDDT.Equals(string.Empty))
                StartUp.M_SD_HDDT = "0";
            StartUp.M_SD_TOKEN = StartupBase.SasObj.GetOption("M_SD_TOKEN").ToString().Trim();
            if (StartUp.M_SD_TOKEN.Equals(string.Empty))
                StartUp.M_SD_TOKEN = "0";
            StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
            StartUp.ngay_gia_px = DateTime.Parse(StartupBase.SasObj.GetOption("M_NGAY_GIA_PX").ToString());
            StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
            StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
            if (StartUpTrans.CommandInfo == null)
            {
                int num = (int)ExMessageBox.Show(835, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else if (StartUpTrans.DmctInfo == null)
            {
                int num = (int)ExMessageBox.Show(840, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else
            {
                StartUp.dtRegInfo = StartupBase.SasObj.GetRegInfo();
                StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
                StartUpTrans.M_ong_ba = StartUpTrans.DmctInfo["m_ong_ba"].ToString();
                StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                int.TryParse(StartUpTrans.DmctInfo["m_bp_bh"].ToString(), out StartUp.M_BP_BH);
                string[] strArray = (string[])null;
                if (StartUpTrans.M_LAN.Trim().Equals("V"))
                {
                    if (StartUpTrans.CommandInfo["Vbrowse2"].ToString().Trim() != "")
                        strArray = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                }
                else if (StartUpTrans.CommandInfo["Ebrowse2"].ToString().Trim() != "")
                    strArray = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                if (strArray != null)
                {
                    StartUp.stringBrowse1 = strArray[0];
                    StartUp.stringBrowse2 = strArray[1];
                    StartUp.stringBrowse3 = strArray[2];
                    StartUp.stringBrowse4 = strArray[3];
                }
                StartUpTrans.filterView = string.Format("{0};{1};{2}", (object)StartUpTrans.DmctInfo["v_phdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctdbf"].ToString().Trim(), (object)StartUpTrans.DmctInfo["v_ctgtdbf"].ToString().Trim());
                this._Form = new FrmSocthda();
                StartUp.M_Tilte = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? StartUpTrans.CommandInfo["bar"].ToString() : StartUpTrans.CommandInfo["bar2"].ToString());
                string str = " AND 1=1";
                if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                    str = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                {
                    SelectTime selectTime = new SelectTime();
                    selectTime.LanguageID = "SOCTHDASelTime";
                    SysFunc.LoadIcon((Window)selectTime);
                    selectTime.ShowDialog();
                    if (selectTime.IsOK)
                    {
                        StartUp.M_NGAY_BAT_DAU = new DateTime?((DateTime)selectTime.M_NGAY_CT1);
                        StartUp.M_NGAY_KET_THUC = new DateTime?((DateTime)selectTime.M_NGAY_CT2);
                        str += string.Format(" AND ngay_ct BETWEEN '{0:yyyyMMdd}' AND '{1:yyyyMMdd}'", (object)StartUp.M_NGAY_BAT_DAU, (object)StartUp.M_NGAY_KET_THUC);
                    }
                    else
                    {
                        if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                            return;
                        Application.Current.Shutdown();
                        return;
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
                SqlCommand sqlcmd = new SqlCommand("SELECT COUNT(*) FROM v_dmmauhd WHERE ma_ct_qs LIKE @ma_ct");
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)("%" + StartUpTrans.Ma_ct + "%");
                StartUp.IsQLHD = Convert.ToDecimal(StartupBase.SasObj.ExcuteScalar(sqlcmd)) != new Decimal(0);
                StartUp.CTKMTable = StartupBase.SasObj.ExcuteReader(new SqlCommand("exec CreateTableCTKM")).Tables[0];
                StartUp.VattuKMHTHTable = StartUp.CTKMTable.Copy();             
                this._Form.Title = StartUp.M_Tilte;
                this._Form.StartUpMain = (StartUpTrans)this;
                SysFunc.LoadIcon((Window)this._Form);
                this._Form.ShowDialog();
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

        private void CreateColumnsKM()
        {
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_sl_km", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_tien_km_nt", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_tien_km", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_thue_km_nt", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_thue_km", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("tien_tc_nt", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("tien_tc", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_tt_km_nt", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
            StartUpTrans.DsTrans.Tables[0].Columns.Add(new DataColumn("t_tt_km", typeof(Decimal))
            {
                DefaultValue = (object)0
            });
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
            string str = "";
            if (StartUp.M_KM_CK == 0)
                str = " AND km_ck = '0'";
            StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'" + str;
            StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
        }

        public static string GetLanguageString(string code, string language)
        {
            return code == "M_MA_NT" ? (StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_nt"].ToString().Equals(StartUpTrans.M_ma_nt0) ? "" : StartUpTrans.DsTrans.Tables[0].Rows[FrmSocthda.iRow]["ma_nt"].ToString()) : (code == "M_MA_NT0" ? StartUpTrans.M_ma_nt0 : code);
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

        public static DataSet CheckData(int mode, int stt_mau_temlate)
        {
            string format = "exec [dbo].{0} @status, @mode, @stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, "[SOCTHDA-CheckData]") : string.Format(format, StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@mode", SqlDbType.Int).Value = (object)mode;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@stt_mau_temlate", SqlDbType.Int).Value = stt_mau_temlate;
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static void DeleteVoucher(
          string _stt_rec,
          string ma_qs,
          ActionTask action,
          bool isNd51)
        {
            SqlCommand sqlcmd1 = new SqlCommand("exec [dbo].[DeleteVoucher] @ma_ct, @stt_rec");
            sqlcmd1.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
            sqlcmd1.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd1);
            if (!isNd51)
                return;
            string format = "exec [dbo].{0} @stt_rec;";
            SqlCommand sqlcmd2 = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 3 ? string.Format(format, (object)"[SOCTHDA-PostCTHHD]") : string.Format(format, (object)StartUpTrans.Post_store[3]));
            if (action == ActionTask.Delete)
            {
                sqlcmd2.CommandText += " DECLARE @ngay_ct smalldatetime;";
                sqlcmd2.CommandText += " DECLARE @so_ct numeric(16, 0), @so_ct1 numeric(16, 0);";
                SqlCommand sqlCommand1 = sqlcmd2;
                sqlCommand1.CommandText = sqlCommand1.CommandText + " SELECT @so_ct = so_ct - 1, @so_ct1 = so_ct1 - 1 FROM dmqs WHERE ma_qs='" + ma_qs.Trim() + "';";
                sqlcmd2.CommandText += " IF @so_ct = @so_ct1";
                SqlCommand sqlCommand2 = sqlcmd2;
                sqlCommand2.CommandText = sqlCommand2.CommandText + "    SELECT @ngay_ct = ngay_qs1 FROM dmqs WHERE ma_qs='" + ma_qs.Trim() + "';";
                sqlcmd2.CommandText += " ELSE";
                SqlCommand sqlCommand3 = sqlcmd2;
                sqlCommand3.CommandText = sqlCommand3.CommandText + "    SELECT TOP 1 @ngay_ct = ngay_ct FROM cthhd WHERE ma_qs='" + ma_qs.Trim() + "' AND so_ct IS NOT NULL ORDER BY ngay_ct DESC;";
                SqlCommand sqlCommand4 = sqlcmd2;
                sqlCommand4.CommandText = sqlCommand4.CommandText + " UPDATE dmqs SET so_ct = @so_ct, ngay_ct = ISNULL(@ngay_ct,ngay_qs1) WHERE ma_qs='" + ma_qs.Trim() + "';";
            }
            sqlcmd2.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd2);
        }

        public static void DeletePT(string _stt_rec, string ma_ct)
        {
            string format = "exec [dbo].{0} @cMa_ct,@stt_rec;";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 11 ? string.Format(format, (object)"[DeleteVoucher]") : string.Format(format, (object)StartUpTrans.Process_Store[11]));
            sqlcmd.Parameters.Add("@cMa_ct", SqlDbType.Char, 3).Value = (object)ma_ct;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
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

        public static void UpdateSl_in(string stt_rec)
        {
            SqlCommand sqlcmd = new SqlCommand("exec [dbo].[Update_sl_in] @stt_rec, @ma_ct, @user_id, @date, @time, @hostname");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
            sqlcmd.Parameters.Add("@user_id", SqlDbType.Decimal).Value = (object)StartUpTrans.M_User_Id;
            SqlParameter sqlParameter = sqlcmd.Parameters.Add("@date", SqlDbType.SmallDateTime);
            DateTime dateTime = DateTime.Now;
            dateTime = dateTime.Date;
            string str = dateTime.ToString("dd-MM-yyyy");
            sqlParameter.Value = (object)str;
            sqlcmd.Parameters.Add("@time", SqlDbType.Char, 8).Value = (object)DateTime.Now.ToString("HH:mm:ss");
            sqlcmd.Parameters.Add("@hostname", SqlDbType.NChar, 100).Value = (object)Environment.MachineName;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        public static DataSet GetHdb(string filter)
        {
            SqlCommand cmd = new SqlCommand("exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @sl_ct");
            cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)"HDB";
            cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)filter;
            cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
            cmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"";
            cmd.Parameters.Add("@sl_ct", SqlDbType.Int).Value = (object)-1;
            DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, cmd);
            for (int index = 0; index < dataSet.Tables[0].Rows.Count; ++index)
            {
                if (dataSet.Tables[0].Rows[index]["ma_nt"].ToString().Trim().ToUpper().Equals(StartUpTrans.M_ma_nt0))
                {
                    dataSet.Tables[0].Rows[index]["t_tien_nt2"] = (object)0;
                    foreach (DataRow dataRow in dataSet.Tables[1].Select("stt_rec LIKE '" + dataSet.Tables[0].Rows[index]["stt_rec"].ToString() + "'"))
                    {
                        dataRow["gia_nt2"] = (object)0;
                        dataRow["tien_nt2"] = (object)0;
                    }
                }
            }
            return dataSet;
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

        public static string EditKmFields(string values)
        {
            return string.Join(";", ((IEnumerable<string>)((IEnumerable<string>)values.Split(';')).ToArray<string>()).Select<string, string>((Func<string, string>)(x =>
         {
             if (x == "")
                 return x;
             string[] properties = x.Split(':');
             return ((IEnumerable<string>)StartUp.hiddenKmFields).Any<string>((Func<string, bool>)(a => a == properties[0])) ? x + ":IV" : x;
         })).ToArray<string>());
        }

        public static string EditGiaVonFields(string values)
        {
            return string.Join(";", ((IEnumerable<string>)((IEnumerable<string>)values.Split(';')).ToArray<string>()).Select<string, string>((Func<string, string>)(x =>
            {
                if (x == "")
                    return x;
                string[] properties = x.Split(':');
                return ((IEnumerable<string>)StartUp.hiddenGiaVonFields).Any<string>((Func<string, bool>)(a => a == properties[0])) ? x + ":IV" : x;
            })).ToArray<string>());
        }
        public static string EditGiaBanFields(string values)
        {
            return string.Join(";", ((IEnumerable<string>)((IEnumerable<string>)values.Split(';')).ToArray<string>()).Select<string, string>((Func<string, string>)(x =>
            {
                if (x == "")
                    return x;
                string[] properties = x.Split(':');
                return ((IEnumerable<string>)StartUp.hiddenGiaBanFields).Any<string>((Func<string, bool>)(a => a == properties[0])) ? x + ":IV" : x;
            })).ToArray<string>());
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
    }
}
