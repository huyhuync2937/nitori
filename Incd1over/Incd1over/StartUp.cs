using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using Microsoft.SqlServer.Server;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Incd1over
{
    public class StartUp : StartupBase
    {
        public static DataSet dsReport = new DataSet();
        public static string hdTuNgay_ = "  -  -  ";
        public static string hdDenNgay_ = "  -  -  ";
        public static string ctTuNgay_ = "  -  -  ";
        public static string ctDenNgay_ = "  -  -  ";
        public static string tableList = "v_Incd1over";
        private static SqlCommand cmd = new SqlCommand();
        public static int kindStyleReport = -1;
        public static string g_strCondition = string.Empty;
        public static string g_strFilter = string.Empty;
        public static string g_ps_ck = "*";
        public static SasFormBrowes.FormBrowse oBrowse;
        public static SasFormBrowes.FormBrowse oBrowse_ct;
        public static FrmLoc _frmLoc;
        public static DataRow commandInfo;
        public static DateTime M_NGAY_KY1;
        public static DateTime M_NGAY_CT1;
        public static string M_MA_DVCS;
        public static string M_IP_TIEN;
        public static string M_IP_TIEN_NT;
        public static string M_IP_SL;
        public static string M_ma_nt0;
        public static object g_hdTuNg;
        public static object g_hdDenNg;
        public static object g_ctTuNg;
        public static object g_ctDenNg;
        public static object g_loaiNvl;
        public static string[] CanChangeValueFields;
        public static int m_user_id = 0;

        public override void Run()
        {
            StartupBase.Namespace = "Incd1over";
            try
            {
                StartUp.CanChangeValueFields = "ngay_bd,ma_may,ma_ca,ngay_kt,ma_ca_kt,sl_cai_lenh,sl_kg_lenh,id,ma_khuon,dvt".Split(',');

                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUp.M_IP_TIEN = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
                StartUp.M_IP_SL = StartupBase.SasObj.GetOption("M_IP_SL").ToString();
                StartUp.M_IP_TIEN_NT = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
                StartUp.M_NGAY_KY1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_NGAY_CT1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_CT1");
                StartUp.M_MA_DVCS = StartupBase.SasObj.DmdvcsInfo.Rows[0][0].ToString();
                StartUp.commandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                if (StartUp.commandInfo == null)
                    return;
                StartUp._frmLoc = new FrmLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void CallGridReport( bool isFirstLoad,object hdTuNg,object hdDenNg,object loaiNvl)
        {
            StartUp.g_hdTuNg = hdTuNg;
            StartUp.g_hdDenNg = hdDenNg;
            StartUp.g_loaiNvl = loaiNvl;

            if (isFirstLoad)
            {
                StartUp.cmd.CommandText = "Exec " + StartUp.commandInfo["store_proc"] + " @hdTuNg, @dhDenNg";
                StartUp.cmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdTuNg);
                StartUp.cmd.Parameters.Add("@dhDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdDenNg);

                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbtong";
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, StartUp.fieldShow(1, 0));
                //StartUp.oBrowse.F3 += new SasFormBrowes.FormBrowse.GridKeyUp_F3(StartUp.oBrowse_F3);
                StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
                StartUp.oBrowse.frmBrw.PreviewKeyDown += new KeyEventHandler(StartUp.FrmBrw_PreviewKeyDown);

                object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport");
                //if (name != null)
                //{
                //    ToolBar toolBar = name as ToolBar;
                //    //for (int i = toolBar.Items.Count - 1; i > 0; i--)
                //    //{
                //    //    if ((toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnRefresh" && (toolBar.Items[i] as SasControls.ToolBarButton).Name.ToString().Trim() != "btnExport")
                //    //    {
                //    //        toolBar.Items.Remove((toolBar.Items[i] as SasControls.ToolBarButton));
                //    //    }
                //    //}
                //    SasControls.ToolBarButton toolBarButton3 = new SasControls.ToolBarButton();
                //    toolBarButton3.BorderBrush = (Brush)Brushes.Transparent;
                //    toolBarButton3.Name = "btnXoa";
                //    toolBarButton3.Text = "Tạo lệnh sản xuất";
                //    toolBarButton3.ToolTip = "F2";
                //    toolBarButton3.ImagePath = "Images\\UpdateSearch.png";
                //    toolBarButton3.Click += new RoutedEventHandler(ToolBarButtonF2_Click);
                //    toolBar.Items.Insert(1, toolBarButton3);

                //}

                //StartUp.oBrowse.F5 += new SasFormBrowes.FormBrowse.GridKeyUp_F5(StartUp.oBrowse_F5);
                StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.oBrowse.frmBrw.PreviewKeyDown += new KeyEventHandler(FrmBrw_PreviewKeyDown);

                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
                //StartUp.oBrowse.SetRowColorByTag("sx2", "0", Colors.Black, true);
                //StartUp.oBrowse.SetRowColorByTag("sx3", "1", Colors.Red, false);
            }
            else
            {
                StartUp.dsReport.Tables.Clear();
                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
                StartUp.dsReport.Tables[0].TableName = "tbDetail";
                StartUp.dsReport.Tables[0].TableName = "tbtong";
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
                StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                StartUp.oBrowse.UpdateSumaryFields();
            }
            if (!isFirstLoad)
                return;
            StartUp.oBrowse.frmBrw.LanguageID = "Incd1over";
            StartUp.oBrowse.ShowDialog();
            StartUp._frmLoc.Close();
        }

        public static void FrmBrw_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.None)
            {
                switch (e.Key)
                {
                    case Key.Space:
                        SelectOne();
                        break;
                    //case Key.F2:
                    //    this.btnView_Click((object)null, (RoutedEventArgs)null);
                    //    break;
                    //case Key.F4:
                    //    this.btnNote_Click((object)null, (RoutedEventArgs)null);
                    //    break;
                    //case Key.F8:
                    //    this.btnXoa_Click((object)null, (RoutedEventArgs)null);
                    //    break;
                }
            }
            if (Keyboard.Modifiers != ModifierKeys.Control)
                return;
            switch (e.Key)
            {
                case Key.A:
                  SelectAll(true);
                    break;
                case Key.U:
                   SelectAll(false);
                    break;
            }
        }
        private static void SelectAll(bool tag)
        {
            if (StartUp.oBrowse.DataGrid.Records.Count == 0)
                return;
            foreach (DataRecord filteredInDataRecord in StartUp.oBrowse.DataGrid.RecordManager.GetFilteredInDataRecords())
                (filteredInDataRecord.DataItem as DataRowView)["chon"] = (object)tag;
        }

        private static void SelectOne()
        {
            if (StartUp.oBrowse.DataGrid.ActiveRecord == null)
                return;
            DataRecord activeRecord = StartUp.oBrowse.DataGrid.ActiveRecord as DataRecord;
            if (activeRecord == null || activeRecord.DataItem == null || activeRecord is FilterRecord)
                return;
            DataRowView dataItem = activeRecord.DataItem as DataRowView;
            bool flag = (bool)dataItem["chon"];
            dataItem["chon"] = (object)!flag;
        }

        private static void ToolBarButtonF2_Click(object sender, RoutedEventArgs e)
        {
            StartUp.m_user_id = Convert.ToInt32(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());

            if (StartUp.oBrowse.ActiveRecord == null)
                return;

            if (StartUp.oBrowse.DataGrid.ActiveCell != null && StartUp.oBrowse.DataGrid.ActiveCell.IsInEditMode)
                StartUp.oBrowse.DataGrid.ActiveCell.EndEditMode();
            StartUp.oBrowse.ActiveRecord.Update();

            DataView dataView = StartUp.oBrowse.DataGrid.DataSource as DataView;

            //DataTable distinctValues = dataView.ToTable(true, "chon", "id_hd", "so_ct0", "ngay_ct0", "ss");
            DataTable distinctValues = dataView.ToTable(true, "chon", "ma_nvl_gop", "vung_may", "nhom_may", "ngay_bd", "so_ca", "ma_may", "ss", "ma_ca");

            DataRowView[] array = (from DataRowView x in distinctValues.DefaultView
                                   where (bool)x["chon"]
                                   select x).ToArray();
            if (array.Length <= 0)
            {
                int num = (int)ExMessageBox.Show(100, StartupBase.SasObj, "Tick vào dòng cần tạo Lệnh sản xuất!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }

            DataRowView[] arrayNgaybd = (from DataRowView x in dataView
                                         where (bool)x["chon"] && x["ngay_bd"].ToString().Trim().Equals("")
                                         select x).ToArray();
            if (arrayNgaybd.Length > 0)
            {
                int num = (int)ExMessageBox.Show(531, StartupBase.SasObj, "Có vật tư chưa khai báo ngày sản xuất!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }

            DataRowView[] arrayMamay = (from DataRowView x in dataView
                                        where (bool)x["chon"] && x["ma_may"].ToString().Trim().Equals("")
                                        select x).ToArray();
            if (arrayMamay.Length > 0)
            {
                int num = (int)ExMessageBox.Show(530, StartupBase.SasObj, "Có hóa đơn chưa khai báo mã máy trên phần mềm kế toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            DataRowView[] arrayMaca = (from DataRowView x in dataView
                                       where (bool)x["chon"] && x["ma_ca"].ToString().Trim().Equals("")
                                       select x).ToArray();
            if (arrayMaca.Length > 0)
            {
                int num = (int)ExMessageBox.Show(532, StartupBase.SasObj, "Có vật tư chưa khai báo ngày sản xuất!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }


            //DataRowView[] arrayKH = (from DataRowView x in dataView
            //                         where (bool)x["chon"] && x["ma_kh"].ToString().Trim().Equals("")
            //                         orderby x["ngay_ct0"]
            //                         select x).ToArray();
            //if (arrayKH.Length > 0)
            //{
            //    int num = (int)ExMessageBox.Show(900, StartupBase.SasObj, "Có hóa đơn chưa khai báo mã khách hàng trên phần mềm kế toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    return;
            //}
            if (ExMessageBox.Show(8007, StartupBase.SasObj, "Bạn có muốn tạo lệnh sản xuất không?", "", MessageBoxButton.YesNo, MessageBoxImage.Asterisk, MessageBoxResult.No) == MessageBoxResult.No)
            {

                return;
            }
            if (array.Any(x => x["ss"]?.ToString() == "1"))
            {
                int num = (int)ExMessageBox.Show(322, StartupBase.SasObj, "Có hóa đơn đã được xử lý kế toán, không thể tạo lại lệnh sản xuất!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }

            if (array.Length > 0)
            {
                string ma_nvl_gop = "";
                string vung_may = "";
                string nhom_may = "";
                string so_ca = "";
                double so_ca_lam_tron = 0;
                double so_ngay = 0;
                string ma_ca = "";

                DateTime ngay_bd = new DateTime();

                foreach (DataRowView row in array)
                {
                    ma_nvl_gop = row["ma_nvl_gop"].ToString();
                    vung_may = row["vung_may"].ToString();
                    nhom_may = row["nhom_may"].ToString();
                    so_ca = row["so_ca"].ToString();
                    so_ca_lam_tron = Math.Ceiling(string.IsNullOrEmpty(so_ca) ? 0 : Double.Parse(so_ca));
                    ngay_bd = Convert.ToDateTime(row["ngay_bd"].ToString());
                    so_ngay = Math.Ceiling(so_ca_lam_tron / 3);


                    for (int i = 0; i < so_ngay; i++)
                    {

                        string ma_qs = "LSX";
                        string so_ct = GetNewSoct(StartupBase.SasObj, ma_qs, true);

                        SqlCommand sqlCommand2 = new SqlCommand();
                        sqlCommand2.CommandText = "dbo.ImexInv_LSX";
                        sqlCommand2.CommandType = CommandType.StoredProcedure;

                        sqlCommand2.Parameters.Add("@User_id", SqlDbType.Int)
                            .Value = StartUp.m_user_id;

                        sqlCommand2.Parameters.Add("@Ma_qs", SqlDbType.VarChar)
                            .Value = ma_qs;


                        sqlCommand2.Parameters.Add("@So_ct", SqlDbType.VarChar)
                        .Value = so_ct;

                        sqlCommand2.Parameters.Add("@Ma_dvcs", SqlDbType.VarChar)
                       .Value = StartupBase.SasObj.M_ma_dvcs.Trim();

                        sqlCommand2.Parameters.Add("@Xu_ly", SqlDbType.Char)
                            .Value = '0';
                        sqlCommand2.Parameters.Add("@Ma_nvl_gop", SqlDbType.VarChar)
                         .Value = ma_nvl_gop;
                        sqlCommand2.Parameters.Add("@Nhom_may", SqlDbType.VarChar)
                            .Value = nhom_may;
                        sqlCommand2.Parameters.Add("@Vung_may ", SqlDbType.VarChar)
                            .Value = vung_may;
                        sqlCommand2.Parameters.Add("@Ngay_kh", SqlDbType.DateTime)
                            .Value = ngay_bd.AddDays(i);


                        DataSet ds = StartupBase.SasObj.ExcuteReader(sqlCommand2);

                        DataTable dataTable2 = ds.Tables[ds.Tables.Count - 1];
                        string stt_rec = dataTable2.Rows[0]["stt_rec"].ToString().Trim();



                        if (CheckValidSoct(StartupBase.SasObj, ma_qs, so_ct, stt_rec))
                        {
                            so_ct = GetNewSoct(StartupBase.SasObj, ma_qs, true);
                        }
                        SqlCommand cmd = new SqlCommand("EXEC dbo.ImexInv_LSX;10 @Stt_rec");
                        cmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = stt_rec;
                        object obj = StartupBase.SasObj.ExcuteScalar(cmd);


                    }

                }
                int num4 = (int)MessageBox.Show("Đã thực hiện thành công (" + so_ngay.ToString() + "  hóa đơn)  ", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);




            }
        }

        public static bool CheckValidSoct(SasObject SasObj, string ma_qs, string so_ct, string stt_rec)
        {
            string format = "{0}";
            SqlCommand sqlCommand = new SqlCommand(string.Format(format, "CheckValidSoct"));
            sqlCommand.Parameters.Add(new SqlParameter("@ma_qs", ma_qs));
            sqlCommand.Parameters.Add(new SqlParameter("@so_ct", so_ct));
            sqlCommand.Parameters.Add(new SqlParameter("@stt_rec", stt_rec));
            sqlCommand.CommandType = CommandType.StoredProcedure;
            return (bool)SasObj.ExcuteScalar(sqlCommand);
        }

        public static string GetNewSoct(SasObject SasObj, string ma_qs, bool isUpdateNewSoCt)
        {
            try
            {
                string cmdText;
                //if (isUpdateNewSoCt)
                //{
                //    string format = "EXEC  {0} '" + ma_qs.Trim() + "'";
                //    cmdText = string.Format(format, "[GetNewSoct#2]");
                //}
                //else if (Convert.ToInt16(SasObj.GetOption("M_AUTO_SOCT").ToString()) == 1)
                //{
                cmdText = "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'";
                string update = "UPDATE dmqs SET so_ct = so_ct + 1 WHERE ma_qs = '" + ma_qs.Trim() + "'";

                //}
                //else
                //{
                //    string text;
                //    string format2 = (text = "EXEC  {0} '" + ma_qs.Trim() + "'");
                //    cmdText = string.Format(format2, "GetNewSoct");
                //}

                StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(update));
                DataTable dataTable = SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0];
                if (dataTable.Rows.Count > 0)
                {
                    DataRow dataRow = dataTable.Rows[0];
                    if (dataRow[1] != null && dataRow[1] != DBNull.Value)
                    {
                        string value = dataRow[1].ToString();
                        return string.Format(dataRow[0].ToString(), Convert.ToDouble(value));
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }

            return "";
        }
        //private static void oBrowse_F3(object sender, EventArgs e)
        //{
        //    DataView dataView = StartUp.oBrowse.DataGrid.DataSource as DataView;

        //    DataTable distinctValues = dataView.ToTable(true, "chon");

        //    DataRowView[] array = (from DataRowView x in distinctValues.DefaultView
        //                           where (bool)x["chon"]
        //                           select x).ToArray();
        //    if (array.Length <= 0)
        //    {
        //        int num = (int)ExMessageBox.Show(110, StartupBase.SasObj, "Phải đánh dấu trước khi sửa!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        //        return;
        //    }
        //    UpdateFieldHandler.Execute();
        //}
        public static void oBrowse_Esc(object sender, EventArgs e)
        {
        }

        public static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.commandInfo["rep_file"].ToString(), StartUp.kindStyleReport);
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.dsReport, "tbDetail");
            reportManager.Preview(StartUp.dsReport);
            SysFunc.ResetFilter(ref StartUp.dsReport, "tbDetail");
        }

        public static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            try
            {
                // Get all rows with id=1 and recalculate subsequent rows with the same ma_may
                RecalculateConsecutiveRows();
                
                bool kindReport = true;
            
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_loaiNvl);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static void RecalculateConsecutiveRows()
        {
            DataView dataView = StartUp.oBrowse.DataGrid.DataSource as DataView;
            if (dataView == null || dataView.Count == 0)
                return;

            const string DATE_FORMAT = "yyyyMMdd";
            
            // Group rows by ma_may and get those with id=1
            var id1Rows = (from DataRowView row in dataView
                          where SafeToInt32(row["id"]) == 1
                          select row).ToList();

            if (id1Rows.Count == 0)
                return;

            // Process each id=1 row
            foreach (DataRowView id1Row in id1Rows)
            {


                string ma_may = SafeToString(id1Row["ma_may"]).Trim();
                string ma_nvl_gop = SafeToString(id1Row["ma_nvl_gop"]).Trim();
                string vung_may = SafeToString(id1Row["vung_may"]).Trim();
                string nhom_may = SafeToString(id1Row["nhom_may"]).Trim();
                
                // Get all rows with the same ma_may, ma_nvl_gop, vung_may, nhom_may
                var groupedRows = (from DataRowView row in dataView
                                   where SafeToString(row["ma_may"]).Trim() == ma_may 
                                   orderby SafeToInt32(row["id"])
                                   select row).ToList();

                if (groupedRows.Count < 2)
                    continue;

                // Get id=1 row values
                string ngayBdStr = SafeToString(id1Row["ngay_bd"]).Trim();
                string maCaStr = SafeToString(id1Row["ma_ca"]).Trim();

                if (string.IsNullOrEmpty(ngayBdStr) || string.IsNullOrEmpty(maCaStr))
                    continue;

                DateTime currentNgayBd = DateTime.MinValue;
                if (!DateTime.TryParse(ngayBdStr, out currentNgayBd))
                    continue;

                string currentMaCa = maCaStr;

                // Update subsequent rows (id >= 2)
                for (int i = 0; i < groupedRows.Count; i++)
                {
                    DataRowView row = groupedRows[i];
                    
                    // Get so_ca for THIS specific row (not just from id=1)
                    string rowSoCaStr = SafeToString(row["so_ca"]).Trim();
                    string rowSoCalenh = SafeToString(row["so_ca_lenh"]).Trim();
                    
                    double rowNumSoCa = 0;
                    double tempVal = 0;
                    if (double.TryParse(rowSoCalenh, out tempVal) && tempVal != 0)
                    {
                        rowNumSoCa = tempVal;
                    }
                    else if (double.TryParse(rowSoCaStr, out tempVal))
                    {
                        rowNumSoCa = tempVal;
                    }
                    
                    double rowSoCaLamTron = Math.Ceiling(rowNumSoCa);
                    
                    // Calculate end shift for current row using THIS row's soCaLamTron
                    string maCaKt = GetCaKetThuc(rowSoCaLamTron, currentMaCa);
                    DateTime ngayKt = GetNgayKetThuc(rowSoCaLamTron, currentNgayBd, currentMaCa);

                    // Build WHERE clause for this specific row
                    int rowId = SafeToInt32(row["id"]);
                    string whereClause = string.Format(" WHERE ma_may = N'{0}' AND " +
                                       "id = {1}",
                                       ma_may, rowId);

                    string updateCommand = string.Format(
                        "UPDATE KHSX3_1 SET ngay_bd = '{0}', ma_ca = '{1}', ma_ca_kt = '{2}', ngay_kt = '{3}'{4}",
                        currentNgayBd.ToString(DATE_FORMAT),
                        currentMaCa,
                        maCaKt,
                        ngayKt.ToString(DATE_FORMAT),
                        whereClause);

                    SqlCommand cmd = new SqlCommand(updateCommand);
                    StartupBase.SasObj.ExcuteNonQuery(cmd);

                    // Prepare for next iteration
                    string nextMaCa = GetNextMaCa(maCaKt);
                    DateTime nextNgayBd = ngayKt;
                    
                    // If transitioning from CA03 to CA01, add one day
                    if (maCaKt == "CA03" && nextMaCa == "CA01")
                    {
                        nextNgayBd = ngayKt.AddDays(1);
                    }
                    
                    currentMaCa = nextMaCa;
                    currentNgayBd = nextNgayBd;
                }
            }
        }

        private static string SafeToString(object obj)
        {
            if (obj == null || obj is DBNull)
                return "";
            return obj.ToString();
        }

        private static int SafeToInt32(object obj)
        {
            if (obj == null || obj is DBNull)
                return 0;
            int result = 0;
            int.TryParse(obj.ToString(), out result);
            return result;
        }

        private static string GetCaKetThuc(double soCa, string ma_ca)
        {
            int ca = (int)Math.Ceiling(soCa);
            int startIndex;
            
            switch (ma_ca)
            {
                case "CA01":
                    startIndex = 1;
                    break;
                case "CA02":
                    startIndex = 2;
                    break;
                case "CA03":
                    startIndex = 3;
                    break;
                default:
                    startIndex = 1;
                    break;
            }

            int totalCa = startIndex - 1 + ca;
            int index = ((totalCa - 1) % 3) + 1;

            if (index == 1)
                return "CA01";
            else if (index == 2)
                return "CA02";
            else
                return "CA03";
        }

        private static DateTime GetNgayKetThuc(double soCa, DateTime ngay_bd, string ma_ca)
        {
            int startIndex;
            switch (ma_ca)
            {
                case "CA01":
                    startIndex = 1;
                    break;
                case "CA02":
                    startIndex = 2;
                    break;
                case "CA03":
                    startIndex = 3;
                    break;
                default:
                    startIndex = 1;
                    break;
            }

            int totalCa = startIndex - 1 + (int)Math.Ceiling(soCa);
            int soNgayTang = (totalCa - 1) / 3;

            return ngay_bd.AddDays(soNgayTang);
        }

        private static string GetNextMaCa(string currentMaCa)
        {
            switch (currentMaCa.Trim())
            {
                case "CA01":
                    return "CA02";
                case "CA02":
                    return "CA03";
                case "CA03":
                    return "CA01";
                default:
                    return "CA01";
            }
        }

        public static string fieldShow(int kindReport, int isDetail)
        {
            string str = string.Empty;
            string[] strArray1 = new string[2];
            switch (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper())
            {
                case "V":
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray2 = StartUp.commandInfo["Vbrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray2[1] : strArray2[0];
                            break;
                        case 2:
                            string[] strArray3 = StartUp.commandInfo["Vbrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray3[1] : strArray3[0];
                            break;
                    }
                    break;
                default:
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray4 = StartUp.commandInfo["Ebrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray4[1] : strArray4[0];
                            break;
                        case 2:
                            string[] strArray5 = StartUp.commandInfo["Ebrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray5[1] : strArray5[0];
                            break;
                    }
                    break;
            }
            return str;
        }

        public static DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "tableInfo";
            string str1 = "ĐƠN HÀNG TỪ NGÀY: " + StartUp.hdTuNgay_ + " ĐẾN NGÀY: " + StartUp.hdDenNgay_;
            string str2 = "PHÁT SINH TỪ NGÀY: " + StartUp.ctTuNgay_ + " ĐẾN NGÀY: " + StartUp.ctDenNgay_;
            string str3 = " ORDER DATE: " + StartUp.hdTuNgay_ + " TO DATE: " + StartUp.hdDenNgay_;
            string str4 = " ARISING FROM: " + StartUp.ctTuNgay_ + " TO DATE: " + StartUp.ctDenNgay_;
            dataTable.Columns.Add(new DataColumn("hdNgDenNg", typeof(string))
            {
                DefaultValue = (object)str1
            });
            dataTable.Columns.Add(new DataColumn("ctNgDenNg", typeof(string))
            {
                DefaultValue = (object)str2
            });
            dataTable.Columns.Add(new DataColumn("hdNgDenNg_Eng", typeof(string))
            {
                DefaultValue = (object)str3
            });
            dataTable.Columns.Add(new DataColumn("ctNgDenNg_Eng", typeof(string))
            {
                DefaultValue = (object)str4
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }
    }
}
