using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using System;
using System.Data;
using System.Data.SqlClient;

namespace SasVoucherLib
{
    public class StartUpTrans : StartupBase
    {
        public static DataSet dsDanhmuc = (DataSet)null;
        public static string Ws_Id = string.Empty;
        public static string Ma_ct = string.Empty;
        public static string Editing_Stt_Rec = string.Empty;
        public static DataRow CommandInfo = (DataRow)null;
        public static DataRow DmctInfo = (DataRow)null;
        public static DataTable tbStatus = (DataTable)null;
        public static DataSet DsTrans = (DataSet)null;
        public static string filterId = string.Empty;
        public static string filterView = string.Empty;
        public static string M_ma_nt0 = string.Empty;
        public new static string M_LAN = string.Empty;
        public static string M_MST_CHECK = string.Empty;
        public static int M_User_Id = 0;
        public static string M_ngay_lct = string.Empty;
        public static string M_ong_ba = string.Empty;
        public static string M_trung_so = string.Empty;
        public static int M_sl_ct0 = 0;
        public static int M_IN_HOI_CK = 0;
        public static int M_CHK_HD_VAO = 0;
        public static string hd_thue = "1";
        public static int M_ROUND;
        public static int M_ROUND_NT;
        public static int M_ROUND_GIA;
        public static int M_ROUND_GIA_NT;
        public static DataRow Ma_ct_info;
        public static string[] Check_Valid_Store;
        public static string[] Post_store;
        public static string[] Process_Store;

        public StartUpTrans()
        {
            if (StartupBase.SasObj == null)
                return;
            StartUpTrans.M_LAN = StartupBase.SasObj.GetOption(nameof(M_LAN)).ToString();
            this.UpdatedsDanhmuc();
        }

        public override void SasLoader(object[] parameters, SasObject SasO)
        {
            StartUpTrans.Editing_Stt_Rec = parameters.Length > 1 ? parameters[1].ToString() : string.Empty;
            if (SasO != null)
                StartUpTrans.M_LAN = SasO.GetOption("M_LAN").ToString();
            base.SasLoader(parameters, SasO);
        }

        public static DataTable GetPhIn()
        {
            DataTable dataTable = (DataTable)null;
            try
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "SELECT * FROM phin WHERE ma_ct LIKE @ma_ct AND stt_rec LIKE @stt_rec";
                sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return dataTable;
        }

        public static DataTable Getdmtk(string filter)
        {
            DataTable dataTable = (DataTable)null;
            try
            {
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "SELECT tk FROM dmtk WHERE " + filter;
                dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return dataTable;
        }

        public static void SetPhIn(DataTable tableIn)
        {
            try
            {
                DataProvider.DeleteRow(StartupBase.SasObj, "phin", string.Format("ma_ct LIKE '{0}' AND stt_rec LIKE '{1}'", (object)StartUpTrans.Ma_ct, (object)tableIn.Rows[0]["stt_rec"].ToString().Trim()));
                DataProvider.UpLoadDataTable(StartupBase.SasObj, "phin", tableIn);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static object GetSl_in(string stt_rec)
        {
            string format = "exec [dbo].{0} @ma_ct,@stt_rec";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 6 ? string.Format(format, (object)nameof(GetSl_in)) : string.Format(format, (object)StartUpTrans.Process_Store[6]));
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
            return StartupBase.SasObj.ExcuteScalar(sqlcmd);
        }

        public static object GetSo_lien(DataRecord currentRecord, string stt_rec)
        {
            SqlCommand sqlcmd = new SqlCommand("SELECT MAX(so_lien) FROM phuserininfo WHERE id_report LIKE @id_report AND stt_rec LIKE @stt_rec");
            sqlcmd.Parameters.Add("@id_report", SqlDbType.Int).Value = (object)currentRecord.Cells["id"].Value.ToString().Trim();
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
            object obj2 = StartupBase.SasObj.ExcuteScalar(sqlcmd);
            return obj2 != DBNull.Value ? obj2 : (object)0;
        }

        public static void UpdateSl_in(string stt_rec, string id_report, string so_lien)
        {
            string format = "exec [dbo].{0} @stt_rec, @ma_ct,@user_id,@date,@time,@hostname,@id_report,@so_lien";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 7 ? string.Format(format, (object)"Update_sl_in") : string.Format(format, (object)StartUpTrans.Process_Store[7]));
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char, 3).Value = (object)StartUpTrans.Ma_ct;
            sqlcmd.Parameters.Add("@user_id", SqlDbType.Decimal).Value = (object)StartUpTrans.M_User_Id;
            sqlcmd.Parameters.Add("@date", SqlDbType.VarChar).Value = (object)DateTime.Now.Date.ToString("yyyyMMdd");
            sqlcmd.Parameters.Add("@time", SqlDbType.Char, 8).Value = (object)DateTime.Now.ToString("HH:mm:ss");
            sqlcmd.Parameters.Add("@hostname", SqlDbType.NChar, 100).Value = (object)Environment.MachineName;
            sqlcmd.Parameters.Add("@id_report", SqlDbType.Int).Value = (object)id_report;
            sqlcmd.Parameters.Add("@so_lien", SqlDbType.Int).Value = (object)so_lien;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        public static int CheckQS(string ma_qs, string ngay_ct, int user_id)
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "if not exists (select 1 from dmqs where ngay_qs1 <= '" + ngay_ct + "' and ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\tselect 1";
            sqlcmd.CommandText += "else";
            SqlCommand sqlCommand1 = sqlcmd;
            sqlCommand1.CommandText = sqlCommand1.CommandText + "\tif not exists (select 1 from dmqs where ma_qs in (select ma_qs from dmuserqs where [user_id] = " + (object)user_id + ") and ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            SqlCommand sqlCommand2 = sqlcmd;
            sqlCommand2.CommandText = sqlCommand2.CommandText + "\t    IF EXISTS(select 1 from dmuserqs where ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\t\t    select 2";
            sqlcmd.CommandText += "\telse";
            SqlCommand sqlCommand3 = sqlcmd;
            sqlCommand3.CommandText = sqlCommand3.CommandText + "\t    if not exists (select 1 from dmqs where (ngay_ct <= '" + ngay_ct + "' OR ngay_ct IS NULL)and ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\t\t    select 3";
            sqlcmd.CommandText += "\t    else";
            sqlcmd.CommandText += "\t\t    select 0";
            return Convert.ToInt32(StartupBase.SasObj.ExcuteScalar(sqlcmd));
        }

        public static int CheckQS(string ma_qs, string ngay_ct, int user_id, string so_ct)
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "if not exists (select 1 from dmqs where ngay_qs1 <= '" + ngay_ct + "' and ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\tselect 1";
            sqlcmd.CommandText += "else";
            SqlCommand sqlCommand1 = sqlcmd;
            sqlCommand1.CommandText = sqlCommand1.CommandText + "\tif not exists (select 1 from dmqs where (not EXISTS(SELECT * FROM dmuserqs WHERE [user_id] = " + (object)user_id + " AND EXISTS(SELECT * FROM dmuserqs WHERE dmqs.ma_qs = dmuserqs.ma_qs AND [user_id] = " + (object)user_id + ")) OR NOT EXISTS (SELECT * FROM dmuserqs WHERE [user_id] =  " + (object)user_id + ") OR NOT EXISTS(SELECT 1 FROM dmuserqs WHERE dmqs.ma_qs = ma_qs)) and ma_qs like '" + ma_qs.Replace("'", "''") + "' )\t";
            SqlCommand sqlCommand2 = sqlcmd;
            sqlCommand2.CommandText = sqlCommand2.CommandText + "\tif not exists (select 1 from dmqs where ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\tbegin";
            SqlCommand sqlCommand3 = sqlcmd;
            sqlCommand3.CommandText = sqlCommand3.CommandText + "\t    IF EXISTS(select 1 from dmuserqs where ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\t\t    select 2";
            sqlcmd.CommandText += "\tend";
            sqlcmd.CommandText += "\telse";
            SqlCommand sqlCommand4 = sqlcmd;
            sqlCommand4.CommandText = sqlCommand4.CommandText + "\t    if exists (select 1 from cthhd where (ngay_ct > '" + ngay_ct + "' AND so_ct < '" + so_ct + "')and ma_qs like '" + ma_qs.Replace("'", "''") + "' )";
            sqlcmd.CommandText += "\t\t    select 3";
            sqlcmd.CommandText += "\t    else";
            sqlcmd.CommandText += "\t\t    select 0";
            return Convert.ToInt32(StartupBase.SasObj.ExcuteScalar(sqlcmd));
        }

        public static DataSet CheckQS(
          string lstma_qs,
          string ngay_ct,
          int user_id,
          string lstso_ct1,
          string lstso_ct2,
          string lsttthd)
        {
            string format = "exec [dbo].{0} @lstma_qs,@stt_rec,@ngay_qs,@user_id,@so_ct1,@so_ct2,@lsttthd";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 0 ? string.Format(format, (object)"CheckQSSo_ct") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[0]));
            sqlcmd.Parameters.Add("@lstma_qs", SqlDbType.VarChar).Value = (object)lstma_qs;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@ngay_qs", SqlDbType.Char, 8).Value = (object)ngay_ct;
            sqlcmd.Parameters.Add("@user_id", SqlDbType.Decimal).Value = (object)user_id;
            sqlcmd.Parameters.Add("@so_ct1", SqlDbType.VarChar).Value = (object)lstso_ct1;
            sqlcmd.Parameters.Add("@so_ct2", SqlDbType.VarChar).Value = (object)lstso_ct2;
            sqlcmd.Parameters.Add("@lsttthd", SqlDbType.VarChar).Value = (object)lsttthd;
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static DataSet CheckQS(
          string lstma_qs,
          string lstso_ct1,
          string lstso_ct2,
          string lsttthd)
        {
            string format = "exec [dbo].{0} @lstma_qs,@stt_rec,@so_ct1,@so_ct2,@lsttthd";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 5 ? string.Format(format, (object)"[CheckQSSo_ct];2") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[5]));
            sqlcmd.Parameters.Add("@lstma_qs", SqlDbType.VarChar).Value = (object)lstma_qs;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@so_ct1", SqlDbType.VarChar).Value = (object)lstso_ct1;
            sqlcmd.Parameters.Add("@so_ct2", SqlDbType.VarChar).Value = (object)lstso_ct2;
            sqlcmd.Parameters.Add("@lsttthd", SqlDbType.VarChar).Value = (object)lsttthd;
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static DataSet CheckQS3(
          string lstma_qs,
          string lstso_ct1,
          string lstso_ct2,
          string lsttthd)
        {
            string format = "exec [dbo].{0} @lstma_qs,@stt_rec,@so_ct1,@so_ct2,@lsttthd";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 6 ? string.Format(format, (object)"[CheckQSSo_ct];3") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[6]));
            sqlcmd.Parameters.Add("@lstma_qs", SqlDbType.VarChar).Value = (object)lstma_qs;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            sqlcmd.Parameters.Add("@so_ct1", SqlDbType.VarChar).Value = (object)lstso_ct1;
            sqlcmd.Parameters.Add("@so_ct2", SqlDbType.VarChar).Value = (object)lstso_ct2;
            sqlcmd.Parameters.Add("@lsttthd", SqlDbType.VarChar).Value = (object)lsttthd;
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
        }

        public static DataRow GetQs(string ma_qs)
        {
            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "SELECT * FROM dmqs a LEFT JOIN dmmauhd b ON a.mau_hd = b.mau_hd WHERE ma_qs LIKE '" + ma_qs + "';";
            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            return table.Rows.Count != 0 ? table.Rows[0] : (DataRow)null;
        }

        public static void BalanceMoney(
          DataView dvDetail,
          string foreign_currency,
          string currency,
          Decimal total_foreign_currency,
          Decimal rate)
        {
            Decimal num = SysFunc.Round(total_foreign_currency * rate, StartUpTrans.M_ROUND);
            Decimal result1 = new Decimal(0);
            Decimal.TryParse(dvDetail.ToTable().Compute("sum(" + currency + ")", dvDetail.RowFilter).ToString(), out result1);
            if (!(num != result1))
                return;
            for (int index = 0; index < dvDetail.Count; ++index)
            {
                Decimal result2 = new Decimal(0);
                Decimal result3 = new Decimal(0);
                Decimal.TryParse(dvDetail[index][foreign_currency].ToString(), out result2);
                Decimal.TryParse(dvDetail[index][currency].ToString(), out result3);
                if (result2 != new Decimal(0))
                {
                    dvDetail[index][currency] = (object)(result3 + Math.Abs(num - result1));
                    break;
                }
            }
        }

        public static int CheckPhanBo(string stt_rec)
        {
            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("stt_rec_hd") && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec_hd"].ToString().Trim()))
                return 0;
            int num = 0;
            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_ct"))
            {
                string upper = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString().ToUpper();
                string str1 = "PN1;PN9;HD9;PN2;PN6;HD1;HD2;HD6;PC1;BN1;PT1;BC1;PNG;PNA;PNB;PNC;HDM;HDA;HDB;HDX;PXF;PNF;HD5;HDL";
                string str2 = "HD3;HD4;PN9;HD9;PC1;BN1;PT1;BC1;PXF";
                if (upper.Equals("PK1"))
                {
                    string format1 = "exec [dbo].{0} @stt_rec";
                    SqlCommand sqlcmd1 = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 0 ? string.Format(format1, (object)"InPostCttt20") : string.Format(format1, (object)StartUpTrans.Post_store[0]));
                    sqlcmd1.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                    if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd1) > 0)
                    {
                        string format2 = "exec [dbo].{0} @stt_rec, @tablename";
                        SqlCommand sqlcmd2 = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 1 ? string.Format(format2, (object)"CheckEditVoucher") : string.Format(format2, (object)StartUpTrans.Check_Valid_Store[1]));
                        sqlcmd2.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                        sqlcmd2.Parameters.Add("@tablename", SqlDbType.VarChar).Value = (object)"cttt20";
                        num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd2);
                    }
                    else
                    {
                        string format2 = "exec [dbo].{0} @stt_rec, @tablename";
                        SqlCommand sqlcmd2 = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 2 ? string.Format(format2, (object)"CheckEditVouchertt") : string.Format(format2, (object)StartUpTrans.Check_Valid_Store[2]));
                        sqlcmd2.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                        sqlcmd2.Parameters.Add("@tablename", SqlDbType.VarChar).Value = (object)"cttt20";
                        num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd2);
                    }
                    if (num == 0)
                    {
                        string format2 = "exec [dbo].{0} @stt_rec";
                        SqlCommand sqlcmd2 = new SqlCommand(StartUpTrans.Post_store == null || StartUpTrans.Post_store.Length <= 1 ? string.Format(format2, (object)"InPostCttt30") : string.Format(format2, (object)StartUpTrans.Post_store[1]));
                        sqlcmd2.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                        if ((int)StartupBase.SasObj.ExcuteScalar(sqlcmd2) > 0)
                        {
                            string format3 = "exec [dbo].{0} @stt_rec, @tablename";
                            SqlCommand sqlcmd3 = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 1 ? string.Format(format3, (object)"CheckEditVoucher") : string.Format(format3, (object)StartUpTrans.Check_Valid_Store[1]));
                            sqlcmd3.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                            sqlcmd3.Parameters.Add("@tablename", SqlDbType.VarChar).Value = (object)"cttt30";
                            num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd3);
                        }
                        else
                        {
                            string format3 = "exec [dbo].{0} @stt_rec, @tablename";
                            SqlCommand sqlcmd3 = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 2 ? string.Format(format3, (object)"CheckEditVouchertt") : string.Format(format3, (object)StartUpTrans.Check_Valid_Store[2]));
                            sqlcmd3.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                            sqlcmd3.Parameters.Add("@tablename", SqlDbType.VarChar).Value = (object)"cttt30";
                            num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd3);
                        }
                    }
                }
                else
                {
                    if (str1.Contains(upper))
                    {
                        bool flag = true;
                        if (upper.Equals("HD6") || upper.Equals("PN6"))
                        {
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_gd") && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("2"))
                                flag = false;
                        }
                        else if (upper.Equals("PC1") || upper.Equals("BN1"))
                        {
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_gd") && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("4"))
                                flag = false;
                        }
                        else if ((upper.Equals("PT1") || upper.Equals("BC1")) && (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_gd") && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("4")))
                            flag = false;
                        if (flag)
                        {
                            string format = "exec [dbo].{0} @stt_rec";
                            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 1 ? string.Format(format, (object)"CheckEditVoucher") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[1]));
                            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                            num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd);
                        }
                    }
                    if (str2.Contains(upper) && num == 0)
                    {
                        bool flag = true;
                        if (upper.Equals("PC1") || upper.Equals("BN1"))
                        {
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_gd") && (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("1") || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("4")))
                                flag = false;
                        }
                        else if ((upper.Equals("PT1") || upper.Equals("BC1")) && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_gd") && (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("1") || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_gd"].ToString().Equals("4")))
                            flag = false;
                        if (flag)
                        {
                            string format = "exec [dbo].{0} @stt_rec";
                            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 2 ? string.Format(format, (object)"CheckEditVouchertt") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[2]));
                            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)stt_rec;
                            num = (int)StartupBase.SasObj.ExcuteScalar(sqlcmd);
                        }
                    }
                }
            }
            return num;
        }

        public static DataRow Getdmgia2(string ma_vt, string ngay_ct, string nh_kh3)
        {
            string format = "EXEC dbo.{0} @Nh_kh3, @ma_vt, @ngay_ban";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 8 ? string.Format(format, (object)"GetDmgia2") : string.Format(format, (object)StartUpTrans.Process_Store[8]));
            sqlcmd.Parameters.Add("@nh_kh3", SqlDbType.Char, 16).Value = (object)nh_kh3;
            sqlcmd.Parameters.Add("@ma_vt", SqlDbType.Char, 16).Value = (object)ma_vt;
            sqlcmd.Parameters.Add("@ngay_ban", SqlDbType.Char, 8).Value = (object)ngay_ct;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            return dataSet.Tables[0].Rows.Count == 1 ? dataSet.Tables[0].Rows[0] : (DataRow)null;
        }

        public static short Getloai_tg(string ma_nt)
        {
            SqlCommand sqlcmd = new SqlCommand("SELECT TOP 1 loai_tg FROM dmnt WHERE ma_nt LIKE @ma_nt");
            sqlcmd.Parameters.Add("@ma_nt", SqlDbType.VarChar, 16).Value = (object)ma_nt;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            return dataSet.Tables[0].Rows.Count == 1 ? Convert.ToInt16(dataSet.Tables[0].Rows[0]["loai_tg"]) : (short)1;
        }

        public static int CheckValidSo_ct(string ma_qs, string so_ct)
        {
            if (StartUpTrans.DmctInfo["m_dmqs"].ToString().Equals("0"))
                return 1;
            string format = "EXEC {0} @Ma_qs, @So_ct";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 4 ? string.Format(format, (object)"[CheckValidSoct#3]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[4]));
            sqlcmd.Parameters.Add("@Ma_qs", SqlDbType.VarChar).Value = (object)ma_qs;
            sqlcmd.Parameters.Add("@So_ct", SqlDbType.VarChar).Value = (object)so_ct;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            return dataSet.Tables[0].Rows.Count == 1 ? (int)Convert.ToInt16(dataSet.Tables[0].Rows[0][0]) : 0;
        }

        public static void UpdateTkSd13(int isDel, int isWarning)
        {
        }

        public static string GetFileNameExportWithSignature(DataRowView drv)
        {
            string str1 = string.Empty;
            if (drv != null && StartupBase.SasObj != null)
            {
                DataTable table = drv.DataView.Table;
                if (table.Columns.Contains("ngay_ct") && table.Columns.Contains("so_ct") && table.Columns.Contains("ma_qs"))
                {
                    string str2;
                    if (!table.Columns.Contains("ma_kh"))
                        str2 = string.Format("{0}_{1}_{2}", (object)Convert.ToDateTime(drv["ngay_ct"]).ToString("yyyyMMdd"), (object)SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, drv["ma_qs"].ToString().Trim()), (object)SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, drv["so_ct"].ToString().Trim()));
                    else
                        str2 = string.Format("{0}_{1}_{2}_{3}", (object)Convert.ToDateTime(drv["ngay_ct"]).ToString("yyyyMMdd"), (object)SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, drv["ma_qs"].ToString().Trim()), (object)SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, drv["so_ct"].ToString().Trim()), (object)SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, drv["ma_kh"].ToString().Trim()));
                    str1 = SysFunc.Cat_Dau(str2).Replace(" ", "");
                }
            }
            return str1;
        }

        public static DataTable GetListColumnForeignCurrency()
        {
            SqlCommand sqlcmd = new SqlCommand("select column_name from information_schema.columns where domain_name like 'udt_tien_nt' OR domain_name like 'udt_gia_nt'");
            return StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
        }

        public void UpdatedsDanhmuc()
        {
            try
            {
                foreach (DataRow row in StartupBase.SasObj.DmdmInfo.Select("loaddm = '1'"))
                {
                    string tablename = row["ma_dm"].ToString().Trim() + "!" + row["table_view"].ToString().Trim();
                    DataTable dt = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0}", row["table_view"].ToString().Trim()))).Tables[0];
                    dt.TableName = tablename;
                    if (StartUpTrans.dsDanhmuc == null) StartUpTrans.dsDanhmuc = new DataSet();

                    if (!StartUpTrans.dsDanhmuc.Tables.Contains(dt.TableName))
                    {
                        StartUpTrans.dsDanhmuc.Tables.Add(dt.Copy());
                    }
                }
            }
            catch
            {
                StartUpTrans.dsDanhmuc = null;
            }
        }
    }
}
