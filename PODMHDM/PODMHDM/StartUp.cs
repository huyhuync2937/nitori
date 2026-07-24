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
using System.Windows.Controls;
namespace PODMHDM
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
        public static string tableList = "v_ph71;v_ct71;v_ct71gt";
        public static SqlCommand TransFilterCmd;
        public static DateTime? M_NGAY_BAT_DAU;
        public static DateTime? M_NGAY_KET_THUC;
        public static DateTime M_ngay_ct0;
        public static string storeproc = "";
        public static DataSet HDBData = null;
        public static bool isOk = false;
        public static DataTable dtRegInfo;
        public static DataRow[] dataRowArray;
        public static FormBrowse obrowseBKCT;
        private static SqlCommand sqlcmdBKCT;
        private static string strBrowseBKCT = "";

        public override void Run()
        {
            StartupBase.Namespace = "PODMHDM";
            try
            {              
                StartUpTrans.Ma_ct = "HDM";
                StartUpTrans.filterId = "PODMHDM";
                FrmPoctpna frmPoctpna = new FrmPoctpna();
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
                StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
                StartUpTrans.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUpTrans.DmctInfo = DataLoader.GetSqlFieldValue(StartupBase.SasObj, "dmct", "ma_ct", StartUpTrans.Ma_ct);
                if (StartUpTrans.CommandInfo == null)
                {
                    int num1 = (int)ExMessageBox.Show(1360, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else if (StartUpTrans.DmctInfo == null)
                {
                    int num2 = (int)ExMessageBox.Show(1365, StartupBase.SasObj, "Chưa khai báo chứng từ hoặc chứng từ ngầm dịnh sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    StartUp.dtRegInfo = StartupBase.SasObj.GetRegInfo();

                    storeproc = StartUpTrans.CommandInfo["store_proc"].ToString();
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
                    StartUpTrans.M_trung_so = StartUpTrans.DmctInfo["m_trung_so"].ToString();
                    StartUpTrans.Check_Valid_Store = StartUpTrans.DmctInfo["Check_Valid_Store"].ToString().Trim().Split('|');
                    StartUpTrans.Post_store = StartUpTrans.DmctInfo["Post_store"].ToString().Trim().Split('|');
                    StartUpTrans.Process_Store = StartUpTrans.DmctInfo["Process_Store"].ToString().Trim().Split('|');
                    int.TryParse(StartUpTrans.DmctInfo["m_sl_ct0"].ToString(), out StartUpTrans.M_sl_ct0);
                    int.TryParse(StartUpTrans.DmctInfo["so_lien"].ToString(), out StartUp.M_so_lien);
                    int.TryParse(StartUpTrans.DmctInfo["m_loc_nsd"].ToString(), out StartUp.M_loc_nsd);
                    int.TryParse(StartUpTrans.DmctInfo["so_dong_in"].ToString(), out StartUp.so_dong_in);
                    string str1 = " AND 1=1";
                    if (!SysFunc.CheckPermission(StartupBase.SasObj, ActionTask.View, StartupBase.Menu_Id))
                        str1 = " AND user_id0 = " + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                    StartUp.M_CHK_DATE_YN = StartupBase.SasObj.GetOption("M_CHK_DATE_YN").ToString().Trim();
                    if (StartUpTrans.Editing_Stt_Rec.Equals(string.Empty) && StartUp.M_CHK_DATE_YN == "1")
                    {
                        SelectTime selectTime = new SelectTime();
                        selectTime.LanguageID = "PODMHDMSelTime";
                        SysFunc.LoadIcon((Window)selectTime);
                        selectTime.ShowDialog();
                        if (selectTime.IsOK)
                        {
                            StartUp.M_NGAY_BAT_DAU = new DateTime?((DateTime)selectTime.M_NGAY_CT1);
                            StartUp.M_NGAY_KET_THUC = new DateTime?((DateTime)selectTime.M_NGAY_CT2);
                            str1 += string.Format(" AND ngay_ct BETWEEN '{0:yyyyMMdd}' AND '{1:yyyyMMdd}'", (object)StartUp.M_NGAY_BAT_DAU, (object)StartUp.M_NGAY_KET_THUC);
                        }
                    }
                    string format = "exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter,@sl_ct;";
                    StartUp.TransFilterCmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 0 ? string.Format(format, (object)"[LoadVoucher]") : string.Format(format, (object)StartUpTrans.Process_Store[0]));
                    StartUp.TransFilterCmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                    SqlParameter sqlParameter = StartUp.TransFilterCmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000);
                    string str2;
                    if (!StartUpTrans.Editing_Stt_Rec.Equals(string.Empty))
                        str2 = "stt_rec = '" + StartUpTrans.Editing_Stt_Rec + "'  AND rtrim(ltrim(ma_ct)) = '" + StartUpTrans.Ma_ct.Trim() + "'";
                    else
                        str2 = "1=1 AND ma_dvcs = '" + StartupBase.SasObj.M_ma_dvcs.Trim() + "' AND rtrim(ltrim(ma_ct)) = '" + StartUpTrans.Ma_ct.Trim() + "'";
                    string str3 = str1;
                    string str4 = str2 + str3;
                    sqlParameter.Value = (object)str4;
                    StartUp.TransFilterCmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)(" rtrim(ltrim(ma_ct)) = '" + StartUpTrans.Ma_ct.Trim() + "'");
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
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Check_Valid_Store == null || StartUpTrans.Check_Valid_Store.Length <= 8 ? string.Format(format, (object)"[PODMHDM-CheckData]") : string.Format(format, (object)StartUpTrans.Check_Valid_Store[8]));
            sqlcmd.Parameters.Add("@status", SqlDbType.Char, 1).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"];
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
            return StartupBase.SasObj.ExcuteReader(sqlcmd);
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
        public static void CallGridReportBKCT(bool isFirstLoad)
        {
            try
            {
                string stt_rec = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim();
                string stt_rec0 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec0"].ToString().Trim();


                StartUp.sqlcmdBKCT = new SqlCommand();
                StartUp.sqlcmdBKCT.CommandText = "Gethdmgh";
                StartUp.sqlcmdBKCT.CommandType = CommandType.StoredProcedure;
                StartUp.sqlcmdBKCT.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = stt_rec.ToString().Trim();
                StartUp.sqlcmdBKCT.Parameters.Add("@stt_rec0", SqlDbType.VarChar).Value = stt_rec0.ToString().Trim();


                DataTable dataTable = StartupBase.SasObj.ExcuteReader(StartUp.sqlcmdBKCT).Tables[0].Copy();
                dataTable.TableName = "tbDetail";

                strBrowseBKCT = "chon:80:h=Chọn;id:80:h=Id;ngay_giao:80:h=Ngày giao;dvt:80:h=DVT;packing:80:h=Packing;he_so:80:h=Hệ số;so_luong:80:h=Số lượng";


                StartUp.obrowseBKCT = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, strBrowseBKCT);
                StartUp.obrowseBKCT.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(StartUp.FormBrowse_Esc);
                //FrmPoctpna.obrowseBKCT.frmBrw.PreviewKeyDown += new KeyEventHandler(FrmPoctpna.FrmBrwBKCT_PreviewKeyDown);
                StartUp.obrowseBKCT.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.obrowseBKCT_CTRL_R);
                StartUp.obrowseBKCT.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.obrowseBKCT.frmBrw.Title = SysFunc.Cat_Dau((StartupBase.M_LAN.Equals("V") ? "Chi tiết lịch giao hàng" : "Delivery schedule details"));
                object name = StartUp.obrowseBKCT.frmBrw.ToolBar.FindName("tbReport");
                if (name != null)
                {
                    ToolBar toolBar = name as ToolBar;
                    for (int i = toolBar.Items.Count - 1; i > 0; i--)
                    {
                        if ((toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnRefresh" && (toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnExport")
                        {
                            toolBar.Items.Remove((toolBar.Items[i] as SasControls.ToolBarButton));
                        }
                    }
                    SasControls.ToolBarButton toolBarButton3 = new SasControls.ToolBarButton();
                    toolBarButton3.Name = "btnMoi";
                    toolBarButton3.Text = "Thêm mới";
                    toolBarButton3.ToolTip = "F4";
                    toolBarButton3.ImagePath = "Images\\UpdateSearch.png";
                    toolBarButton3.Click += new RoutedEventHandler(ToolBarButtonF4_Click);
                    toolBar.Items.Insert(1, toolBarButton3);

                    //SasControls.ToolBarButton toolBarButton1 = new SasControls.ToolBarButton();
                    //toolBarButton1.Name = "btnXoa";
                    //toolBarButton1.Text = "Xóa";
                    //toolBarButton1.ToolTip = "F5";
                    //toolBarButton1.ImagePath = "Images\\AddNew.png";
                    //toolBarButton1.Click += new RoutedEventHandler(ToolBarButtonF5_Click);
                    //toolBar.Items.Insert(2, toolBarButton1);

                    SasControls.ToolBarButton toolBarButton2 = new SasControls.ToolBarButton();
                    toolBarButton2.Name = "btnUpdate";
                    toolBarButton2.Text = "Cập nhật";
                    toolBarButton2.ToolTip = "F6";
                    toolBarButton2.ImagePath = "Images\\Edit.png";
                    toolBarButton2.Click += new RoutedEventHandler(ToolBarButtonF6_Click);
                    toolBar.Items.Insert(3, toolBarButton2);
                }


                StartUp.obrowseBKCT.frmBrw.LanguageID = "PODMHDM_brwBKCT";
                StartUp.obrowseBKCT.ShowDialog();
            }
            catch (Exception e)
            {

            }

        }
        private static void ToolBarButtonF4_Click(object sender, RoutedEventArgs e)
        {
            string ma_vt = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["ma_vt"].ToString().Trim();

            string stt_rec = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec"].ToString().Trim();
            string stt_rec0 = StartUpTrans.DsTrans.Tables[1].DefaultView[0]["stt_rec0"].ToString().Trim();
            SqlCommand sqlcmd = new SqlCommand("select dvt,packing from dmvt where ma_vt = @ma_vt");
            sqlcmd.Parameters.Add("@ma_vt", SqlDbType.NVarChar).Value = (object)ma_vt;

            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];


            FrmSetValue frmSetValue = new FrmSetValue(table);
            frmSetValue.ShowDialog();

            if (frmSetValue.DialogResult == true)
            {
                string packing = frmSetValue.txtpacking.Text.Trim();
                string he_so = frmSetValue.txthe_so.Text.Trim();
                string dvt = frmSetValue.txtDvt1.Text.Trim();
                string so_luong = frmSetValue.txtso_luong.Text.Trim();
                string ngay_giao = Convert.ToDateTime(frmSetValue.txtNgay_bh.Value)
                    .ToString("yyyy-MM-dd");


                string packingSql = Convert.ToDecimal(packing).ToString();
                string heSoSql = Convert.ToDecimal(he_so).ToString();
                string soLuongSql = Convert.ToDecimal(so_luong).ToString();

                string updateCommand = string.Format(
                    @"INSERT INTO dmhdmctgh
        (stt_rec, stt_rec0, packing, he_so, dvt, so_luong, ngay_giao)
        VALUES
        ('{0}', '{1}', {2}, {3}, '{4}', {5}, '{6}')",
                    stt_rec, stt_rec0, packingSql, heSoSql, dvt, soLuongSql, ngay_giao);

                SqlCommand cmd = new SqlCommand(updateCommand);
                StartupBase.SasObj.ExcuteNonQuery(cmd);
                CallGridReportBKCT(false);

            }

        }


        //private static void ToolBarButtonF5_Click(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {
        //        DataView dataView = StartUp.obrowseBKCT.DataGrid.DataSource as DataView;
        //        if (StartUp.obrowseBKCT.ActiveRecord == null)
        //            return;

        //        if (StartUp.obrowseBKCT.DataGrid.ActiveCell != null && StartUp.obrowseBKCT.DataGrid.ActiveCell.IsInEditMode)
        //            StartUp.obrowseBKCT.DataGrid.ActiveCell.EndEditMode();
        //        StartUp.obrowseBKCT.ActiveRecord.Update();
        //        DataTable distinctValues = dataView.ToTable(true, "chon", "id");

        //        DataRowView[] array = (from DataRowView x in distinctValues.DefaultView
        //                               where (bool)x["chon"]
        //                               select x).ToArray();
        //        if (array.Length > 0)
        //        {
        //            string listSo_ct = string.Join(", ", array.Select((DataRowView x) => x["id"].ToString().Trim()).ToArray());
        //            foreach (DataRowView row in array)
        //            {
        //                int id = Convert.ToInt32(row["id"]);

        //                string sql = "DELETE FROM dmhdmctgh WHERE id = " + id;
        //                SqlCommand cmd = new SqlCommand(sql);

        //                StartUp.SasObj.ExcuteNonQuery(cmd);
        //            }
        //        }
        //    }
        //    catch (Exception ex) { }

        //}

        private static void ToolBarButtonF6_Click(object sender, RoutedEventArgs e)
        {
            DataView dataView = StartUp.obrowseBKCT.DataGrid.DataSource as DataView;
            if (StartUp.obrowseBKCT.ActiveRecord == null)
                return;

            if (StartUp.obrowseBKCT.DataGrid.ActiveCell != null && StartUp.obrowseBKCT.DataGrid.ActiveCell.IsInEditMode)
                StartUp.obrowseBKCT.DataGrid.ActiveCell.EndEditMode();
            StartUp.obrowseBKCT.ActiveRecord.Update();
            DataView dv = new DataView(dataView.Table);
            dv.RowFilter = "chon = true";

            DataTable distinctValues = dv.ToTable(true, "chon", "id", "dvt", "ngay_giao", "so_luong", "he_so", "packing");

            FrmSetValue frmSetValue = new FrmSetValue(distinctValues);
            frmSetValue.ShowDialog();

            if (frmSetValue.DialogResult == true)
            {
                string packing = frmSetValue.txtpacking.Text.Trim();
                string he_so = frmSetValue.txthe_so.Text.Trim();
                string dvt = frmSetValue.txtDvt1.Text.Trim();
                string so_luong = frmSetValue.txtso_luong.Text.Trim();
                string ngay_giao = Convert.ToDateTime(frmSetValue.txtNgay_bh.Value)
                    .ToString("yyyy-MM-dd");
                string id = distinctValues.Rows[0]["id"].ToString();

                string packingSql = Convert.ToDecimal(packing).ToString();
                string heSoSql = Convert.ToDecimal(he_so).ToString();
                string soLuongSql = Convert.ToDecimal(so_luong).ToString();

                string updateCommand = string.Format(
      @"UPDATE dmhdmctgh
        SET packing = {1},
            he_so = {2},
            dvt = '{3}',
            so_luong = {4},
            ngay_giao = '{5}'
        WHERE id = {0}",
      id,
      packingSql,
      heSoSql,
      dvt,
      soLuongSql,
      ngay_giao);

                SqlCommand cmd = new SqlCommand(updateCommand);
                StartupBase.SasObj.ExcuteNonQuery(cmd);
                CallGridReportBKCT(false);

            }

        }

        private static void FormBrowse_Esc(object sender, EventArgs e)
        {

        }
        public static void obrowseBKCT_CTRL_R(object sender, EventArgs e)
        {
            CallGridReportBKCT(false);
        }
    }
}
