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

namespace COLGH
{
    public class StartUp : StartupBase
    {
        public static DataSet dsReport = new DataSet();
        public static string hdTuNgay_ = "  -  -  ";
        public static string hdDenNgay_ = "  -  -  ";
        public static string ctTuNgay_ = "  -  -  ";
        public static string ctDenNgay_ = "  -  -  ";
        public static string tableList = "v_COLGH";
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
        public static string g_maKh = string.Empty;
        public static string g_maVt = string.Empty;
        public static string[] CanChangeValueFields;
        public static int m_user_id = 0;
        private const string DATE_FORMAT = "yyyyMMdd";

        public override void Run()
        {
            StartupBase.Namespace = "COLGH";
            try
            {
                StartUp.CanChangeValueFields = "ETA_PIC,ETA_NCC,ETA_XN".Split(',');
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

        public static void CallGridReport( bool isFirstLoad,object hdTuNg,object hdDenNg, string filter, string maKh, string maVt)
        {
            StartUp.g_hdTuNg = hdTuNg;
            StartUp.g_hdDenNg = hdDenNg;

            StartUp.g_strFilter = filter;
            StartUp.g_maKh = maKh;
            StartUp.g_maVt = maVt;
            if (isFirstLoad)
            {
                StartUp.cmd.CommandText = "Exec " + StartUp.commandInfo["store_proc"] + " @hdTuNg, @dhDenNg,@ma_kh,@ma_vt";
                StartUp.cmd.Parameters.Add("@hdTuNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdTuNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdTuNg);
                StartUp.cmd.Parameters.Add("@dhDenNg", SqlDbType.VarChar).Value = string.IsNullOrEmpty(hdDenNg.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)hdDenNg);
                StartUp.cmd.Parameters.Add("@ma_kh", SqlDbType.VarChar).Value = string.IsNullOrEmpty(maKh) ? (object)"" : (object)maKh;
                StartUp.cmd.Parameters.Add("@ma_vt", SqlDbType.VarChar).Value = string.IsNullOrEmpty(maVt) ? (object)"" : (object)maVt;

                StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                DataTable dataTable = StartUp.dsReport.Tables[0].Copy();


                StartUp.dsReport.Tables[0].TableName = "tbDetail";
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataTable.DefaultView, StartUp.fieldShow(1, 0));
                //StartUp.oBrowse.F3 += new SasFormBrowes.FormBrowse.GridKeyUp_F3(StartUp.oBrowse_F3);
                StartUp.oBrowse.F7 += new SasFormBrowes.FormBrowse.GridKeyUp_F7(StartUp.oBrowse_F7);
                StartUp.oBrowse.frmBrw.PreviewKeyDown += new KeyEventHandler(StartUp.FrmBrw_PreviewKeyDown);

                object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport");
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
                    toolBarButton3.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton3.Name = "btnXoa";
                    toolBarButton3.Text = "Cập nhật dữ liệu";
                    toolBarButton3.ToolTip = "F2";
                    toolBarButton3.ImagePath = "Images\\UpdateSearch.png";
                    toolBarButton3.Click += new RoutedEventHandler(ToolBarButtonF2_Click);
                    toolBar.Items.Insert(1, toolBarButton3);

                    SasControls.ToolBarButton toolBarButton4 = new SasControls.ToolBarButton();
                    toolBarButton4.BorderBrush = (Brush)Brushes.Transparent;
                    toolBarButton4.Name = "btnXoa";
                    toolBarButton4.Text = "Xác nhận";
                    toolBarButton4.ToolTip = "F3";
                    toolBarButton4.ImagePath = "Images\\UpdateSearch.png";
                    toolBarButton4.Click += new RoutedEventHandler(ToolBarButtonF3_Click);
                    toolBar.Items.Insert(2, toolBarButton4);
                }

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
                dataTable.TableName = "tbDetail";
                StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
                StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
                StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                StartUp.oBrowse.UpdateSumaryFields();
            }
            if (!isFirstLoad)
                return;
            StartUp.oBrowse.frmBrw.LanguageID = "COLGH";
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
        public static DateTime? ParseYyyyMMdd(string input)
        {
            if (string.IsNullOrEmpty(input))
                return null;

            return DateTime.TryParseExact(
                input.Trim(),
                "yyyyMMdd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime result)
                ? result
                : (DateTime?)null;
        }
        private static void ToolBarButtonF2_Click(object sender, RoutedEventArgs e)
        {
            StartUp.m_user_id = Convert.ToInt32(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());

            if (StartUp.oBrowse.ActiveRecord == null)
                return;

            if (StartUp.oBrowse.DataGrid.ActiveCell != null && StartUp.oBrowse.DataGrid.ActiveCell.IsInEditMode)
                StartUp.oBrowse.DataGrid.ActiveCell.EndEditMode();
            StartUp.oBrowse.ActiveRecord.Update();

            string OldValue = "";
            FieldCollection fields = StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].Fields;
            Cell cell = StartUp.oBrowse.frmBrw.oBrowse.ActiveCell;
            if (cell.Field.Name.ToString() != "ETA_XN" && cell.Field.Name.ToString() != "ETA_NCC" && cell.Field.Name.ToString() != "ETA_PIC" )
            {
                return;
            }
            if (((IEnumerable<string>)StartUp.CanChangeValueFields).Any<string>((Func<string, bool>)(x => x == cell.Field.Name)))
                OldValue = cell.Value.ToString();

            FrmSetValue frmSetValue = new FrmSetValue(OldValue);
            frmSetValue.cbField.ItemsProvider = new ComboBoxItemsProvider();
            foreach (string updateValueField in StartUp.CanChangeValueFields)
            {
                string field = updateValueField;
                if (fields.Any<Field>((Func<Field, bool>)(x => x.Name == field)))
                    frmSetValue.cbField.ItemsProvider.Items.Add((object)new ComboBoxDataItem((object)field, fields[field].Label.ToString()));
                else
                    frmSetValue.cbField.ItemsProvider.Items.Add((object)new ComboBoxDataItem((object)field, field));
            }
            frmSetValue.cbField.ValuePath = "Value";

            if (((IEnumerable<string>)StartUp.CanChangeValueFields).Any<string>((Func<string, bool>)(x => x == cell.Field.Name)))
            {
                frmSetValue.cbField.Value = (object)cell.Field.Name;
            }
            else
                frmSetValue.cbField.SelectedIndex = 0;

            if (frmSetValue.cbField.Value.ToString().Trim() == "ETA_XN")
            {
                frmSetValue.Title = "Cap nhat thong tin ngày duyệt";
          
            }
          
            else if (frmSetValue.cbField.Value.ToString().Trim() == "ETA_PIC")
            {
                frmSetValue.Title = "Cap nhat thong tin ngày pic";
             
            }
            else if (frmSetValue.cbField.Value.ToString().Trim() == "ETA_NCC")
            {
                frmSetValue.Title = "Cap nhat thong tin ngày ncc";
               
            }
            bool? nullable = frmSetValue.ShowDialog();
            if ((!nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0 || (frmSetValue.cbField.Value == null || frmSetValue.cbField.Value.ToString().Trim() == ""))
            {
                return;
            }
            string col = frmSetValue.cbField.Value.ToString();
            string value = frmSetValue.txtNgay_bd.dValue.ToString(DATE_FORMAT);
            if (ExMessageBox.Show(1122, StartupBase.SasObj, "Có muốn cập nhật số liệu không?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) == MessageBoxResult.No)
                return;
            string field_update = "";
            Record[] array1 = (Record[])StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().ToArray<DataRecord>();
            int length = array1.Length;
            SasFormBrowes.FrmWaiting frmwait = new SasFormBrowes.FrmWaiting(length);
            try
            {
                frmwait.Show();
                for (int index1 = 0; index1 < length; ++index1)
                {
                    frmwait.Set(index1 + 1);
                    DataRecord dataRecord = array1[index1] as DataRecord;
                    DataRow rowFromBrowse = (dataRecord.DataItem as DataRowView).Row;
                    if (rowFromBrowse["chon"].ToString().Trim().ToUpper() == "TRUE" || rowFromBrowse["chon"].ToString().Trim().ToUpper() == "1")
                    {
                        if (col.Equals("ETA_XN"))
                        {
                            field_update = "ngay_giao_duyet";
                        }
                        else if (col.Equals("ETA_NCC"))
                        {
                            field_update = "ngay_giao_ncc";
                        }
                        else
                        {
                            field_update = "ngay_giao";

                        }
                        DateTime? etaPic = rowFromBrowse["ETA_PIC"] == DBNull.Value
       ? (DateTime?)null
       : Convert.ToDateTime(rowFromBrowse["ETA_PIC"]);

                        string pic = etaPic.HasValue
                            ? etaPic.Value.ToString("yyyyMMdd")
                            : string.Empty;
                        SqlCommand sqlCommand = new SqlCommand(string.Format("UPDATE [{0}] SET [{1}] = @Value WHERE LTRIM(RTRIM(so_ct)) = '{2}' and ngay_giao = '{3}' and ma_td = '{4}'", "dmhdmctgt", field_update, rowFromBrowse["so_ct"].ToString().Trim(), pic, rowFromBrowse["ma_vt"].ToString().Trim()));
                        sqlCommand.Parameters.Add("@Value", SqlDbType.NVarChar).Value = value.ToString().Trim();
                        StartupBase.SasObj.ExcuteNonQuery(sqlCommand);
                    }
                }
                frmwait.Set(length);
                frmwait.Hide();
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_strFilter, StartUp.g_maKh, StartUp.g_maVt);

            }
            catch (Exception ex)
            {
                frmwait.Hide();
                ErrorLog.CatchMessage(ex);
            }
            finally
            {
                frmwait.Hide();
                frmwait.Close();
                frmwait = null;
            }

        }
        private static void ToolBarButtonF3_Click(object sender, RoutedEventArgs e)
        {
            StartUp.m_user_id = Convert.ToInt32(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());

            if (StartUp.oBrowse.ActiveRecord == null)
                return;

            if (StartUp.oBrowse.DataGrid.ActiveCell != null && StartUp.oBrowse.DataGrid.ActiveCell.IsInEditMode)
                StartUp.oBrowse.DataGrid.ActiveCell.EndEditMode();
            StartUp.oBrowse.ActiveRecord.Update();

            string OldValue = "";
            FieldCollection fields = StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].Fields;
            Cell cell = StartUp.oBrowse.frmBrw.oBrowse.ActiveCell;
           
            
            string field_update = "";
            Record[] array1 = (Record[])StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().ToArray<DataRecord>();
            int length = array1.Length;
            SasFormBrowes.FrmWaiting frmwait = new SasFormBrowes.FrmWaiting(length);
            try
            {
                frmwait.Show();
                for (int index1 = 0; index1 < length; ++index1)
                {
                    frmwait.Set(index1 + 1);
                    DataRecord dataRecord = array1[index1] as DataRecord;
                    DataRow rowFromBrowse = (dataRecord.DataItem as DataRowView).Row;
                    if (rowFromBrowse["chon"].ToString().Trim().ToUpper() == "TRUE" || rowFromBrowse["chon"].ToString().Trim().ToUpper() == "1")
                    {
                       
                        DateTime? etaPic = rowFromBrowse["ETA_PIC"] == DBNull.Value
       ? (DateTime?)null
       : Convert.ToDateTime(rowFromBrowse["ETA_PIC"]);

                        string pic = etaPic.HasValue
                            ? etaPic.Value.ToString("yyyyMMdd")
                            : string.Empty;

                        DateTime? etaNcc = rowFromBrowse["ETA_NCC"] == DBNull.Value
     ? (DateTime?)null
     : Convert.ToDateTime(rowFromBrowse["ETA_PIC"]);

                        string picNcc = etaNcc.HasValue
                            ? etaNcc.Value.ToString("yyyyMMdd")
                            : string.Empty;

                        if (picNcc != string.Empty)
                            field_update = picNcc;
                        else
                            field_update = pic;

                            SqlCommand sqlCommand = new SqlCommand(string.Format("UPDATE [{0}] SET [{1}] = @Value WHERE LTRIM(RTRIM(so_ct)) = '{2}' and ngay_giao = '{3}'  and ma_td = '{4}'", "dmhdmctgt", "ngay_giao_duyet", rowFromBrowse["so_ct"].ToString().Trim(), pic, rowFromBrowse["ma_vt"].ToString().Trim()));
                        sqlCommand.Parameters.Add("@Value", SqlDbType.NVarChar).Value = field_update.ToString().Trim();

                        StartupBase.SasObj.ExcuteNonQuery(sqlCommand);
                    }
                }
                frmwait.Set(length);
                frmwait.Hide();
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_strFilter, StartUp.g_maKh, StartUp.g_maVt);

            }
            catch (Exception ex)
            {
                frmwait.Hide();
                ErrorLog.CatchMessage(ex);
            }
            finally
            {
                frmwait.Hide();
                frmwait.Close();
                frmwait = null;
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
                
                bool kindReport = true;
            
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_strFilter, StartUp.g_maKh, StartUp.g_maVt);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
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
