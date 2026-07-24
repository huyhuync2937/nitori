using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace Poctpna
{
    public class StartUp : StartUpTrans
    {
        public static string M_User_name = string.Empty;
        public static string titleWindow = string.Empty;
        public static string M_ma_nt = string.Empty;
        public static int M_so_lien = 0;
        public static int M_loc_nsd = 0;
        public new static int M_IN_HOI_CK = 0;
        public static int M_MA_MS = 0;
        public static string M_CHK_TON_VT = string.Empty;
        public static int so_dong_in = 0;
        public static string stringBrowse1 = "";
        public static string stringBrowse2 = "";
        public static string M_CHK_DATE_YN = "";
        public static string M_MAU_THUE_CK = "0";
        public static string tableList = "v_ph71;v_ct71;v_ct71gt";
        public static SqlCommand TransFilterCmd;
        public static DataTable dtRegInfo;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DateTime M_ngay_ct0;
        public static FrmPoctpna frmPoctpna;
        public static string M_QL_LO_CK = "0";
        public override void Run()
        {
            StartupBase.Namespace = "Poctpna";
            try
            {
                StartUpTrans.Ma_ct = "PNA";
                StartUpTrans.filterId = "POCTPNA";
                StartUp.dtRegInfo = StartupBase.SasObj.GetRegInfo();
                frmPoctpna = new FrmPoctpna();
                StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                StartUp.M_User_name = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
                StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
                StartUpTrans.M_ROUND = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND"));
                StartUpTrans.M_ROUND_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_NT"));
                StartUpTrans.M_ROUND_GIA = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
                StartUpTrans.M_ROUND_GIA_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
                StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                StartUp.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
                StartUp.M_MA_MS = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_MA_MS").ToString());
                StartUpTrans.M_CHK_HD_VAO = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_CHK_HD_VAO").ToString());
                StartUp.M_MAU_THUE_CK = StartupBase.SasObj.GetOption("M_MAU_THUE_CK").ToString().Trim();
                StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
                StartUp.M_QL_LO_CK = StartupBase.SasObj.GetOption("M_QL_LO_CK").ToString().Trim();
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                if (StartUpTrans.CommandInfo == null)
                {
                    int num1 = (int)ExMessageBox.Show(555, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else if (StartUpTrans.DmctInfo == null)
                {
                    int num2 = (int)ExMessageBox.Show(560, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm dịnh sai!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    string[] strArray = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                    if (StartUpTrans.M_LAN != "V")
                        strArray = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                    if (strArray != null)
                    {
                        StartUp.stringBrowse1 = strArray[0];
                        StartUp.stringBrowse2 = strArray[1];
                    }
                    StartUpTrans.M_ngay_lct = StartUpTrans.DmctInfo["m_ngay_lct"].ToString();
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
                        selectTime.LanguageID = "PoctpnaSelTime";
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

                    frmPoctpna.LoadStatus(); // Load tbStatus
                    frmPoctpna.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar"].ToString()).ToString();
                    if (StartUpTrans.M_LAN != "V")
                        frmPoctpna.Title = SysFunc.Cat_Dau(StartUpTrans.CommandInfo["bar2"].ToString()).ToString();
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
            StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + stt_rec + "'";
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[POCTPNA-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static DataSet GetHdm(string filter)
        {
            SqlCommand cmd = new SqlCommand("exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @sl_ct");
            cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)"HDM";
            cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)filter;
            cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"";
            cmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"";
            cmd.Parameters.Add("@sl_ct", SqlDbType.Int).Value = (object)-1;
            DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, cmd);
            for (int index = 0; index < dataSet.Tables[0].Rows.Count; ++index)
            {
                if (dataSet.Tables[0].Rows[index]["ma_nt"].ToString().Trim().ToUpper().Equals(StartUpTrans.M_ma_nt0))
                {
                    dataSet.Tables[0].Rows[index]["t_tien_nt"] = (object)0;
                    foreach (DataRow dataRow in dataSet.Tables[1].Select("stt_rec LIKE '" + dataSet.Tables[0].Rows[index]["stt_rec"].ToString() + "'"))
                    {
                        dataRow["gia_nt"] = (object)0;
                        dataRow["tien_nt"] = (object)0;
                    }
                }
            }
            return dataSet;
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

        public static DataRow Getdmgia0(string ma_vt, string ngay_ct)
        {
            string format = "EXEC dbo.GetDmgia0 @ma_vt, @ngay_nhap";
            SqlCommand sqlcmd = new SqlCommand(format);
            sqlcmd.Parameters.Add("@ma_vt", SqlDbType.Char, 16).Value = (object)ma_vt;
            sqlcmd.Parameters.Add("@ngay_nhap", SqlDbType.Char, 8).Value = (object)ngay_ct;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            return dataSet.Tables[0].Rows.Count == 1 ? dataSet.Tables[0].Rows[0] : (DataRow)null;
        }
    }
}
