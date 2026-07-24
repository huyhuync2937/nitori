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

namespace COSXKSX.KSXW
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
        public static string M_CHK_DATE_YN = "";
        public static string M_CHK_TON_VT = string.Empty;
        public static int so_dong_in = 0;
        public static string M_BP_BH = string.Empty;
        public static string tableList = "v_KSXPHW;v_KSXCTW";
        public static SqlCommand TransFilterCmd;
        public static DataSet DataSource = new DataSet();
        public new static int M_ROUND_GIA;
        public new static int M_ROUND_GIA_NT;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DateTime M_ngay_ct0;
        public static int M_AR_CK;
        public static string v_tablePH = "v_ksxphw";
        public static string v_tablesanpham = "v_ksxsanphamw";
        public static string v_tablenguyenlieu = "v_ksxnguyenlieuw";
        public static string v_tablemaymoc = "v_ksxmaymocw";
        public static string v_tablenguonluc = "v_ksxnguonlucw";
        public static string tablePH = "ksxphw";
        public static string tablesanpham = "ksxsanphamw";
        public static string tablenguyenlieu = "ksxnguyenlieuw";
        public static string tablemaymoc = "ksxmaymocw";
        public static string tablenguonluc = "ksxnguonlucw";    
        public override void Run()
        {
            StartupBase.Namespace = "COSXKSX.KSXW";
            try
            {
                StartUpTrans.Ma_ct = "KSW";
                StartUpTrans.filterId = "COSXKSX";
                FrmPoctpna frmPoctpna = new FrmPoctpna();
                StartUpTrans.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUpTrans.M_User_Id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                StartUp.M_User_name = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
                StartUpTrans.Ws_Id = StartupBase.SasObj.GetOption("M_WS_ID").ToString();
                StartUp.M_ROUND_GIA = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
                StartUp.M_ROUND_GIA_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
                StartUpTrans.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                StartUpTrans.M_MST_CHECK = StartupBase.SasObj.GetOption("M_MST_CHECK").ToString().Trim();
                StartUp.M_IN_HOI_CK = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString());
                StartUp.M_MA_MS = (int)Convert.ToInt16(StartupBase.SasObj.GetOption("M_MA_MS").ToString());
                StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
                StartUpTrans.tbStatus = StartupBase.SasObj.GetPostInfo(StartUpTrans.Ma_ct);
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                if (StartUpTrans.CommandInfo == null)
                {
                    int num1 = (int)ExMessageBox.Show(2360, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    Application.Current.Shutdown(0);
                }
                else if (StartUpTrans.DmctInfo == null)
                {
                    int num2 = (int)ExMessageBox.Show(2365, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm dịnh sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    Application.Current.Shutdown(0);
                }
                else
                {
                    string str1 = "";
                    string str2 = "";
                    string[] strArray;
                    if (StartUpTrans.M_LAN.Equals("V"))
                        strArray = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|');
                    else
                        strArray = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|');
                    if (strArray != null)
                    {
                        str1 = strArray[0];
                        str2 = strArray[1];
                    }
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
                    string str3 = " AND 1=1";
                    if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                        str3 = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                    StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                    if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                    {
                        SelectTime selectTime = new SelectTime();
                        selectTime.LanguageID = "COSXKSX.KSXWSelTime";
                        SysFunc.LoadIcon((Window)selectTime);
                        selectTime.ShowDialog();
                        if (selectTime.IsOK)
                        {
                            StartUp.M_NGAY_BAT_DAU = new DateTime?((DateTime)selectTime.M_NGAY_CT1);
                            StartUp.M_NGAY_KET_THUC = new DateTime?((DateTime)selectTime.M_NGAY_CT2);
                            str3 += string.Format(" AND ngay_ksx BETWEEN '{0:yyyyMMdd}' AND '{1:yyyyMMdd}'", (object)StartUp.M_NGAY_BAT_DAU, (object)StartUp.M_NGAY_KET_THUC);
                        }
                    }
                  
                    StartUpTrans.DsTrans = new DataSet();
                    string filter =(StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) ? "1=1 AND ma_dvcs = '" + StartupBase.SasObj.M_ma_dvcs.Trim() + "'" : "stt_rec = '" + StartUpTrans.Editing_Stt_Rec + "'") + str3;                  
                   
                    StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablePH + " where ma_ct = @ma_ct and " + filter + " order by ngay_ksx ");
                    StartUp.TransFilterCmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                    DataTable tblph = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                    tblph.TableName = "tblPH";
                    StartUpTrans.DsTrans.Tables.Add(tblph.Copy());
                    StartUp.TransFilterCmd = new SqlCommand("select stt_rec as Column1,*,stt_rec as stt_rec1 from " + StartUp.v_tablesanpham);                  
                    DataTable tblsanpham = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                    tblsanpham.TableName = "tblsanpham";
                    StartUpTrans.DsTrans.Tables.Add(tblsanpham.Copy());
                    StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablenguyenlieu);
                    DataTable tblnguyenlieu = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                    tblnguyenlieu.TableName = "tblnguyenlieu";
                    StartUpTrans.DsTrans.Tables.Add(tblnguyenlieu.Copy());
                    StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablemaymoc);
                    DataTable tblmaymoc = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                    tblmaymoc.TableName = "tblmaymoc";
                    StartUpTrans.DsTrans.Tables.Add(tblmaymoc.Copy());
                    StartUp.TransFilterCmd = new SqlCommand("select *,stt_rec as stt_rec1 from " + StartUp.v_tablenguonluc);
                    DataTable tblnguonluc = StartupBase.SasObj.ExcuteReader(StartUp.TransFilterCmd).Tables[0];
                    tblnguonluc.TableName = "tblnguonluc";
                    StartUpTrans.DsTrans.Tables.Add(tblnguonluc.Copy());

                    StartUpTrans.DsTrans.Tables[0].DefaultView.Sort = "ngay_ksx asc, so_ct asc";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0 ASC";
                    StartUpTrans.DsTrans.Tables[2].DefaultView.Sort = "stt_rec0 ASC";
                    StartUpTrans.DsTrans.Tables[3].DefaultView.Sort = "stt_rec0 ASC";
                    StartUpTrans.DsTrans.Tables[4].DefaultView.Sort = "stt_rec0 ASC";

                    DataRow row = StartUpTrans.DsTrans.Tables[0].NewRow();
                    row["stt_rec"] = (object)string.Empty;
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
                string format = "Delete "+StartUp.tablePH+" where stt_rec = @stt_rec;";
                SqlCommand sqlcmd = new SqlCommand(format);
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                format = "Delete " + StartUp.tablesanpham + " where stt_rec = @stt_rec;";
                sqlcmd = new SqlCommand(format);
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                format = "Delete " + StartUp.tablenguyenlieu + " where stt_rec = @stt_rec;";
                sqlcmd = new SqlCommand(format);
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                format = "Delete " + StartUp.tablemaymoc + " where stt_rec = @stt_rec;";
                sqlcmd = new SqlCommand(format);
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)_stt_rec;
                StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                format = "Delete " + StartUp.tablenguonluc + " where stt_rec = @stt_rec;";
                sqlcmd = new SqlCommand(format);
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[COSXKSX.KSXW-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }
    }
}
