using Infragistics.Windows.DataPresenter;
using SasControls;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Linq;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using System.IO;
using Aspose.Cells;

namespace COSXLSX.CODMNVL
{
    public class StartUp : StartupBase
    {
        public static DataSet DataSourceReport = new DataSet();
        private static SqlCommand cmd = new SqlCommand();
        public static bool isNew = true;
        public static int M_ROUND = 0;
        public static ActionTask currActionTask = ActionTask.None;
        public static string M_FORMAT_STT_REC0 = "{0:0000}";
        private static DataRow CommandInfo;
        //private static SasFormBrowes.FormBrowse2 oBrowse;
        //private static FormBrowse1 oBrowse;
        private static SasFormBrowes.FormBrowse oBrowse;
        private static string SumFields = "sl_dm";

        public static DateTime M_ngay_ct0;
        public static codmnvlLoc _frmLoc;
        public static string sso_lsx_loc;
        public static string sMa_sp_loc;
        public static string sMa_bpht_loc;
        public static string sMa_hd_loc;
        public static DateTime? sNgay1_loc;
        public static DateTime? sNgay2_loc;
        public static string sMa_ky_loc;
        public static DataTable dataTable1;

        public static string Ten_file = "";
        public static string txtFileName = "";
        private bool bResult = false;

        public static string ma_ky_truoc = "";
        public static string so_lsx_truoc = "";
        public static string Ma_sp_truoc = "";
        public static string ma_bpht_truoc = "";
        public static string ma_hd_truoc = "";
        public override void Run()
        {
            StartupBase.Namespace = "COSXLSX.CODMNVL";
            try
            {
                StartUp.M_ngay_ct0 = Convert.ToDateTime(StartupBase.SasObj.GetSysvar("M_NGAY_KY1"));
                StartUp.M_FORMAT_STT_REC0 = StartupBase.SasObj.GetSysvar("M_FORMAT_STT_REC0").ToString();
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                if (StartUp.CommandInfo == null)
                {
                    int num = (int)ExMessageBox.Show(2075, StartupBase.SasObj, "Chưa khai báo command hoặc command ngầm định sai!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                        return;
                    Application.Current.Shutdown();
                }
                StartUp._frmLoc = new codmnvlLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void CallGridVouchers(
          bool isFirstLoad,
          string ma_ky,
          string so_lsx,
          string ma_sp,
          string ma_bpht,
          string ma_hd)
        {
            try
            {
                DataTable tb1;
                DataTable tb2;
                if (isFirstLoad)
                {
                    StartUp.CommandInfo["store_proc"].ToString().Split('|');
                    StartUp.cmd = new SqlCommand(StartUp.CommandInfo["store_proc"].ToString());
                    StartUp.cmd.CommandType = CommandType.StoredProcedure;
                    StartUp.cmd.Parameters.Add("@Ma_ky", SqlDbType.VarChar).Value = ma_ky;
                    DataSet ds = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                    //DataSourceReport = ds.Copy();
                    //StartUp.dataTable1 = dataSet.Tables[1].Copy();

                    tb1 = ds.Tables[0].Copy();
                    tb1.TableName = "tbMain";
                    tb2 = ds.Tables[1].Copy();
                    tb2.TableName = "tbDetail";
                    DataSourceReport.Tables.Add(tb1);
                    DataSourceReport.Tables.Add(tb2);
                    string[] strArray = StartUp.CommandInfo["VBrowse1"].ToString().Trim().Split('|');
                    if (StartupBase.M_LAN == "E")
                        strArray = StartUp.CommandInfo["EBrowse1"].ToString().Trim().Split('|');
                    string strBrowse = strArray[0];
                    string strBrowseCt = strArray[1];
                    StartUp.oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, tb1.DefaultView, strBrowse);
                    //StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)DataSourceReport.Tables[1].DefaultView;
                    StartUp.oBrowse.Esc += new SasFormBrowes.FormBrowse.GridKeyUp_Esc(StartUp.oBrowse_Esc);
                    StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                    StartUp.oBrowse.F4 += new SasFormBrowes.FormBrowse.GridKeyUp_F4(StartUp.oBrowse_F4);
                    StartUp.oBrowse.F8 += new SasFormBrowes.FormBrowse.GridKeyUp_F8(StartUp.oBrowse_F8);
                    StartUp.oBrowse.F6 += new SasFormBrowes.FormBrowse.GridKeyUp_F6(StartUp.oBrowse_F6);
                    StartUp.oBrowse.F5 += new SasFormBrowes.FormBrowse.GridKeyUp_F5(StartUp.oBrowse_F5);

                    StartUp.oBrowse.frmBrw.PreviewKeyDown += new KeyEventHandler(StartUp.frmBrw_PreviewKeyDown);
                    StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                    StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString()) + " - " + (StartupBase.M_LAN.Equals("V") ? " Ky: " : " Period: ") + StartUp.sMa_ky_loc;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnRefresh") as Button).Visibility = Visibility.Collapsed;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnEdit") as Button).Visibility = Visibility.Collapsed;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnDetail") as Button).Visibility = Visibility.Collapsed;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnPrint") as Button).Visibility = Visibility.Collapsed;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnOption") as Button).Visibility = Visibility.Collapsed;
                    (StartUp.oBrowse.frmBrw.ToolBar.FindName("btnCalculator") as Button).Visibility = Visibility.Collapsed;



                    ToolBarButton toolBarButton1 = new ToolBarButton();
                    toolBarButton1.Text = StartupBase.M_LAN.Equals("V") ? "Mới" : "New";
                    toolBarButton1.Text2 = StartupBase.M_LAN.Equals("V") ? "Mới" : "New";
                    toolBarButton1.Name = "btnMoi";
                    toolBarButton1.ToolTip = (object)"F4";
                    toolBarButton1.ImagePath = "Images\\AddNew.png";
                    toolBarButton1.BorderBrush = (Brush)null;
                    toolBarButton1.Click += new RoutedEventHandler(StartUp.btnMoi_Click);
                    ToolBarButton toolBarButton2 = new ToolBarButton();
                    toolBarButton2.Text = "Copy";
                    toolBarButton2.Text2 = "Copy";
                    toolBarButton2.Name = "btnCopy";
                    toolBarButton2.ToolTip = (object)"Ctrl+F4";
                    toolBarButton2.ImagePath = "Images\\Copy.png";
                    toolBarButton2.BorderBrush = (Brush)null;
                    toolBarButton2.Click += new RoutedEventHandler(StartUp.btnCopy_Click);
                    ToolBarButton toolBarButton3 = new ToolBarButton();
                    toolBarButton3.Text = StartupBase.M_LAN.Equals("V") ? "Sửa" : "Edit";
                    toolBarButton3.Text2 = StartupBase.M_LAN.Equals("V") ? "Sửa" : "Edit";
                    toolBarButton3.Name = "btnSua";
                    toolBarButton3.ToolTip = (object)"F3";
                    toolBarButton3.ImagePath = "Images\\Edit.png";
                    toolBarButton3.BorderBrush = (Brush)null;
                    toolBarButton3.Click += new RoutedEventHandler(StartUp.btnSua_Click);
                    ToolBarButton toolBarButton4 = new ToolBarButton();
                    toolBarButton4.Text = StartupBase.M_LAN.Equals("V") ? "Xem" : "View";
                    toolBarButton4.Text2 = StartupBase.M_LAN.Equals("V") ? "Xem" : "View";
                    toolBarButton4.Name = "btnView";
                    toolBarButton4.ToolTip = (object)"F2";
                    toolBarButton4.ImagePath = "Images\\View.png";
                    toolBarButton4.BorderBrush = (Brush)null;
                    toolBarButton4.Click += new RoutedEventHandler(StartUp.btnXem_Click);
                    ToolBarButton toolBarButton5 = new ToolBarButton();
                    toolBarButton5.Text = StartupBase.M_LAN.Equals("V") ? "Xóa" : "Delete";
                    toolBarButton5.Text2 = StartupBase.M_LAN.Equals("V") ? "Xóa" : "Delete";
                    toolBarButton5.Name = "btnXoa";
                    toolBarButton5.ToolTip = (object)"F8";
                    toolBarButton5.ImagePath = "Images\\Delete.png";
                    toolBarButton5.BorderBrush = (Brush)null;
                    toolBarButton5.Click += new RoutedEventHandler(StartUp.btnXoa_Click);
                    ToolBarButton toolBarButton6 = new ToolBarButton();
                    toolBarButton6.Text = StartupBase.M_LAN.Equals("V") ? "Sao chép kỳ trước" : "Copy of previous period";
                    toolBarButton6.Text2 = StartupBase.M_LAN.Equals("V") ? "Sao chép kỳ trước" : "Copy of previous period";
                    toolBarButton6.Name = "btnCopyky";
                    toolBarButton6.ToolTip = (object)"F5";
                    toolBarButton6.ImagePath = "Images\\Copy.png";
                    toolBarButton6.BorderBrush = (Brush)null;
                    toolBarButton6.Click += new RoutedEventHandler(StartUp.btnCopyky_Click);
                    ToolBarButton toolBarButton7 = new ToolBarButton();
                    toolBarButton7.Text = StartupBase.M_LAN.Equals("V") ? "Đổi mã vật tư" : "Change material code";
                    toolBarButton7.Text2 = StartupBase.M_LAN.Equals("V") ? "Đổi mã vật tư" : "Change material code";
                    toolBarButton7.Name = "btnDoima";
                    toolBarButton7.ToolTip = (object)"F6";
                    toolBarButton7.ImagePath = "Images\\Copy.png";
                    toolBarButton7.BorderBrush = (Brush)null;
                    toolBarButton7.Click += new RoutedEventHandler(StartUp.btnDoimaVT_Click);
                    ToolBarButton toolBarButton8 = new ToolBarButton();
                    toolBarButton8.Text = StartupBase.M_LAN.Equals("V") ? "Lấy mẫu Excel" : "Get Excel Templates";
                    toolBarButton8.Text2 = StartupBase.M_LAN.Equals("V") ? "Lấy mẫu Excel" : "Get Excel Templates";
                    toolBarButton8.Name = "btnMauExcel";
                    toolBarButton8.ToolTip = (object)"F7";
                    toolBarButton8.ImagePath = "Images\\View.png";
                    toolBarButton8.BorderBrush = (Brush)null;
                    toolBarButton8.Click += new RoutedEventHandler(StartUp.btnmauExcel_Click);
                    ToolBarButton toolBarButton9 = new ToolBarButton();
                    toolBarButton9.Text = StartupBase.M_LAN.Equals("V") ? "Import Excel" : "Import Excel";
                    toolBarButton9.Text2 = StartupBase.M_LAN.Equals("V") ? "Import Excel" : "Import Excel";
                    toolBarButton9.Name = "btnImport";
                    toolBarButton9.ToolTip = (object)"F9";
                    toolBarButton9.ImagePath = "Images\\Edit.png";
                    toolBarButton9.BorderBrush = (Brush)null;
                    toolBarButton9.Click += new RoutedEventHandler(StartUp.btnimportExcel_Click);
                    PrintToolBar PrintToolBar10 = new PrintToolBar();
                    PrintToolBar10.Content = StartupBase.M_LAN.Equals("V") ? "\n Ctrl + A: Chọn tất cả;Ctrl + U: Bỏ chọn tất cả" : "Ctrl + A: Chọn tất cả;Ctrl + U: Bỏ chọn tất cả";
                    PrintToolBar10.Name = "btntext";
                    if (StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport") is ToolBar name)
                    {
                        name.Items.Insert(0, (object)toolBarButton1);
                        name.Items.Insert(1, (object)toolBarButton2);
                        name.Items.Insert(2, (object)toolBarButton3);
                        name.Items.Insert(3, (object)toolBarButton4);
                        name.Items.Insert(4, (object)toolBarButton5);
                        name.Items.Insert(5, (object)toolBarButton6);
                        name.Items.Insert(6, (object)toolBarButton7);
                        name.Items.Insert(7, (object)toolBarButton8);
                        name.Items.Insert(8, (object)toolBarButton9);
                        name.Items.Insert(15, (object)PrintToolBar10);
                    }
                }
                //else
                //{
                //    StartUp.DataSourceReport.Tables.Remove("tbMain");
                //    StartUp.DataSourceReport.Tables.Remove("tbDetail");
                //    DataSourceReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                //    DataSourceReport.Tables[0].TableName = "tbMain";
                //    DataSourceReport.Tables[1].TableName = "tbDetail";
                //}
                if (!isFirstLoad)
                    return;
                StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => StartUp.oBrowse.frmBrw.oBrowse.Focus()));
                StartUp.oBrowse.frmBrw.LanguageID = "COSXLSX.CODMNVL_2";
                StartUp.oBrowse.ShowDialog();
                if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
        public static void ReloadData2(DataTable dt, string tag)
        {
            DataRow[] rows;
            rows = DataSourceReport.Tables[1].Select("tag = '"+tag+"'"); // UserName is Column Name
            foreach (DataRow r in rows)
                r.Delete();
            foreach (DataRow row in dt.Rows)
            {
                DataSourceReport.Tables[1].ImportRow(row);
            }
            DataSourceReport.Tables[1].AcceptChanges();
            DataSourceReport.Tables[1].AsEnumerable().OrderBy(r => r.Field<string>("ma_sp"));
            for (int index = 0; index < DataSourceReport.Tables[1].DefaultView.Count; index++)
            {
                DataSourceReport.Tables[1].Rows[index]["stt"] = index + 1;
            }
        }

        public static void ReloadData(string ma_ky)
        {
            StartUp.DataSourceReport.Tables.Remove("tbMain");
            StartUp.DataSourceReport.Tables.Remove("tbDetail");
            DataSet ds = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
            DataTable tb2 = ds.Tables[1].Copy();
            tb2.TableName = "tbDetail";
            DataTable tb1 = ds.Tables[0].Copy();
            tb1.TableName = "tbMain";
            StartUp.oBrowse.frmBrw.oBrowse.DataSource = tb1.DefaultView;

            DataSourceReport.Tables.Add(tb1);
            DataSourceReport.Tables.Add(tb2);
            StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
            StartUp.oBrowse.UpdateSumaryFields();
        }

        private static void frmBrw_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F2)
                StartUp.V_Xem();
            else if (e.Key == Key.F3)
                StartUp.V_Sua();
            else if (e.Key == Key.F4)
            {
                StartUp.V_Moi();
            }
            else if (e.Key == Key.F8)
            {
                StartUp.V_Xoa();
            }
            else
            {
                if (StartUp.oBrowse.DataGrid.ActiveRecord != null && StartUp.oBrowse.DataGrid.ActiveRecord is FilterRecord)
                    return;
                if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Space)
                {
                    e.Handled = true;
                    StartUp.SelectEntry();
                }
                if (Keyboard.Modifiers != ModifierKeys.Control)
                    return;
                switch (e.Key)
                {
                    case Key.A:
                        StartUp.SelectAll(true);
                        break;
                    case Key.U:
                        StartUp.SelectAll(false);
                        break;
                }
            }
            //if (StartUp.oBrowse.ActiveRecord != null && StartUp.oBrowse.ActiveRecord is FilterRecord)
            //    return;
            //else if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Space)
            //{
            //    e.Handled = true;
            //    SelectEntry();
            //    if (Keyboard.Modifiers != ModifierKeys.Control)
            //    {
            //        return;
            //    } else
            //    {
            //        switch (e.Key)
            //        {
            //            case Key.A:
            //                SelectAll(true);
            //                break;
            //            case Key.U:
            //                SelectAll(false);
            //                break;
            //        }
            //    }
            //}

        }


        private static void SelectEntry()
        {
            if (!(StartUp.oBrowse.ActiveRecord is DataRecord activeRecord) || activeRecord.RecordType != RecordType.DataRecord)
                return;

            if (StartUp.oBrowse.frmBrw.oBrowse.ActiveCell != null)
                StartUp.oBrowse.frmBrw.oBrowse.ActiveCell = (Infragistics.Windows.DataPresenter.Cell)null;
            Infragistics.Windows.DataPresenter.Cell cell = activeRecord.Cells["chon"];


            if (cell.Value == null || cell.Value is DBNull)
                cell.Value = false;
            else
                cell.Value = !(bool)cell.Value;
        }

        private static void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.V_Copy();
        }

        private static void SelectAll(bool tag)
        {
            if (StartUp.oBrowse.frmBrw.oBrowse.ActiveCell != null)
                StartUp.oBrowse.frmBrw.oBrowse.ActiveCell = (Infragistics.Windows.DataPresenter.Cell)null;
            DataSet y = StartUp.DataSourceReport.Copy();
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref y, "tbDetail");
            DataTable g = y.Tables[1].Copy();
            if (tag == false)
            {
                for (int index = 0; index < StartUp.oBrowse.frmBrw.oBrowse.Records.Count; ++index)
                {

                    (StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord).Cells["chon"].Value = tag;

                }
            }
            else
            {
                int a = g.Rows.Count;
                int b = StartUp.oBrowse.frmBrw.oBrowse.Records.Count;
                for (int index = 0; index < StartUp.oBrowse.frmBrw.oBrowse.Records.Count; ++index)
                {
                    for (int index0 = 0; index0 < g.Rows.Count; ++index0)
                    {
                        if ((Int64)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord).Cells["stt"].Value == (Int64)g.Rows[index0]["stt"])
                        {
                            (StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord).Cells["chon"].Value = tag;
                        }
                    }
                }
            }

        }

        private static void btnSua_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.V_Sua();
        }

        private static void btnXem_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.V_Xem();
        }

        private static void btnMoi_Click(object sender, RoutedEventArgs e)
        {
            StartUp.V_Moi();
        }

        private static void btnXoa_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.V_Xoa();
        }

        public static void V_Moi()
        {
            StartUp.isNew = true;
            StartUp.currActionTask = ActionTask.Add;
            DataRecord activeRecord1 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            DataRecord activeRecord2 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            string empty = string.Empty;
            string ssoLsxLoc1 = StartUp.sso_lsx_loc;
            string ssoLsxLoc2 = StartUp.sso_lsx_loc;
            string sMaSpLoc = StartUp.sMa_sp_loc;
            string sMahdLoc = StartUp.sMa_sp_loc;
            //int count = StartUp.DataSourceReport.Tables[0].DefaultView.Count;
            int index = 0;
            FrmCapNhat frmCapNhat = new FrmCapNhat(empty, ssoLsxLoc1, ssoLsxLoc2, sMaSpLoc, sMahdLoc, "", "");
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Them dinh muc nguyen vat lieu" : "Add bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk || StartUp.oBrowse == null)
                return;
            index = 0;
            DataRowView data = (DataRowView)null;
            ReloadData(StartUp.sMa_ky_loc);

        }

        public static void V_Copy()
        {
            StartUp.isNew = true;
            StartUp.currActionTask = ActionTask.Copy;
            DataRecord activeRecord1 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            DataRecord activeRecord2 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            string empty1 = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            string empty4 = string.Empty;
            if (activeRecord1 == null || activeRecord1.RecordType != RecordType.DataRecord)
                return;
            int index = 0;
            int count = StartUp.DataSourceReport.Tables[1].DefaultView.Count;
            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_ky"].Value.ToString() + '_' + activeRecord1.Cells["ma_px"].Value.ToString(), "", "", activeRecord1.Cells["ma_sp"].Value.ToString(), "", activeRecord1.Cells["ma_px"].Value.ToString(), activeRecord1.Cells["ma_vt"].Value.ToString());
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Them dinh muc nguyen vat lieu" : "Add bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk || StartUp.oBrowse == null)
                return;
            DataRowView data = (DataRowView)null;
            ReloadData(StartUp.sMa_ky_loc);

            //if (StartUp.DataSourceReport.Tables[0].DefaultView.Count > 0)
            //{
            //    for (int index1 = 0; index1 < StartUp.DataSourceReport.Tables[0].DefaultView.Count; ++index1)
            //    {
            //        if (frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_sp"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_ky"].ToString().Trim()))
            //        {
            //            index = index1;
            //            data = StartUp.DataSourceReport.Tables[0].DefaultView[index1];
            //            break;
            //        }
            //    }
            //    StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
            //   {
            //       if (data != null)
            //           StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord);
            //       else
            //           StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = StartUp.oBrowse.frmBrw.oBrowse.Records.FirstOrDefault<Record>((Func<Record, bool>)(x => (x as DataRecord).DataItem == data));
            //   }));
            //}
        }

        public static void V_Sua()
        {
            StartUp.isNew = false;
            StartUp.currActionTask = ActionTask.Edit;
            DataRecord activeRecord1 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            DataRecord activeRecord2 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            string empty1 = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            string empty4 = string.Empty;
            if (activeRecord1 == null || activeRecord1.RecordType != RecordType.DataRecord)
                return;
            int index = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.Index;

            string ma_sp = activeRecord1.Cells["ma_sp"]?.Value?.ToString() ?? "";
            string ma_ky = activeRecord1.Cells["ma_ky"]?.Value?.ToString() ?? "";
            string ma_px = activeRecord1.Cells["ma_px"]?.Value?.ToString() ?? "";
            string ma_vt = activeRecord1.Cells["ma_vt"]?.Value?.ToString() ?? "";

            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_ky"].Value.ToString() + '_' + activeRecord1.Cells["ma_px"].Value.ToString(), "", "", activeRecord1.Cells["ma_sp"].Value.ToString(), "", activeRecord1.Cells["ma_px"].Value.ToString(), activeRecord1.Cells["ma_vt"].Value.ToString());
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Sua dinh muc nguyen vat lieu" : "Edit bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk)
                return;
            index = 0;
            ReloadData(StartUp.sMa_ky_loc);
            /*if (StartUp.DataSourceReport.Tables[1].DefaultView.Count > 0)
            {
                for (int index1 = 0; index1 < StartUp.DataSourceReport.Tables[1].DefaultView.Count; ++index1)
                {
                    //if (frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_sp"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[1].DefaultView[index1]["ma_ky"].ToString().Trim()))
                    //{
                    //    index = index1;
                    //    break;
                    //}
                }
                StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord)));
            }*/
        }

        public static void V_Xem()
        {
            StartUp.isNew = false;
            StartUp.currActionTask = ActionTask.View;
            DataRecord activeRecord1 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            DataRecord activeRecord2 = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord as DataRecord;
            string empty1 = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            string empty4 = string.Empty;
            if (activeRecord1 == null || activeRecord1.RecordType != RecordType.DataRecord)
                return;
            int index = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.Index;
            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_ky"].Value.ToString() + '_' + activeRecord1.Cells["ma_px"].Value.ToString(), "", "", activeRecord1.Cells["ma_sp"].Value.ToString(), "", activeRecord1.Cells["ma_px"].Value.ToString(), activeRecord1.Cells["ma_vt"].Value.ToString());
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Xem dinh muc nguyen vat lieu" : "View bill of material entry";
            frmCapNhat.EnableEditMode(false);
            frmCapNhat.ShowDialog();
        }
        public static void V_Xoa()
        {
            string empty1 = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            if (!(StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord is DataRecord activeRecord) || activeRecord.RecordType != RecordType.DataRecord || ExMessageBox.Show(1535, StartupBase.SasObj, "Có chắc chắn xóa không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
                return;
            string str1 = activeRecord.Cells["so_lsx"].Value.ToString();
            string str2 = activeRecord.Cells["ma_bpht"].Value.ToString();
            string str3 = activeRecord.Cells["ma_sp"].Value.ToString();
            string str4 = activeRecord.Cells["ma_hd"].Value.ToString();
            SqlCommand sqlcmd1 = new SqlCommand();
            sqlcmd1.CommandText = "Delete [cosxlsx-Dmdmvtct] Where so_lsx = @so_lsx and ma_bpht = @Mabpht and ma_sp = @Masp and ma_hd = @Mahd and ma_ky = @Maky";
            sqlcmd1.Parameters.Add("@so_lsx", SqlDbType.VarChar).Value = str1;
            sqlcmd1.Parameters.Add("@Mabpht", SqlDbType.VarChar).Value = str2;
            sqlcmd1.Parameters.Add("@Masp", SqlDbType.VarChar).Value = str3;
            sqlcmd1.Parameters.Add("@Mahd", SqlDbType.VarChar).Value = str4;
            sqlcmd1.Parameters.Add("@Maky", SqlDbType.VarChar).Value = StartUp.sMa_ky_loc;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd1);
            SqlCommand sqlcmd2 = new SqlCommand();
            sqlcmd2.CommandText = "Delete [cosxlsx-Dmdmvt] Where so_lsx = @so_lsx and ma_bpht = @Mabpht and ma_sp = @Masp and ma_hd = @Mahd and ma_ky = @Maky";
            sqlcmd2.Parameters.Add("@so_lsx", SqlDbType.VarChar).Value = str1;
            sqlcmd2.Parameters.Add("@Mabpht", SqlDbType.VarChar).Value = str2;
            sqlcmd2.Parameters.Add("@Masp", SqlDbType.VarChar).Value = str3;
            sqlcmd2.Parameters.Add("@Mahd", SqlDbType.VarChar).Value = str4;
            sqlcmd2.Parameters.Add("@Maky", SqlDbType.VarChar).Value = StartUp.sMa_ky_loc;
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd2);
            int visibleIndex = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.VisibleIndex;
            StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
            StartUp.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
            {
                if (StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().Count<DataRecord>() > visibleIndex)
                {
                    StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().ElementAt<DataRecord>(visibleIndex - 1);
                }
                else
                {
                    StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
                    StartUp.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                    {
                        if (StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord == null || StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.VisibleIndex != -1)
                            return;
                        StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
                    }));
                }
            }));
            ReloadData(StartUp.sMa_ky_loc);

        }
        //public static void V_Xoa()
        //{
        //    string empty1 = string.Empty;
        //    string empty2 = string.Empty;
        //    string empty3 = string.Empty;
        //    if (!(StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord is DataRecord activeRecord) || activeRecord.RecordType != RecordType.DataRecord || ExMessageBox.Show(1535, StartupBase.SasObj, "Có chắc chắn xóa không?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes) != MessageBoxResult.Yes)
        //        return;


        //    StartUp.oBrowse.DataGrid.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
        //    DataTable dataSource1 = StartUp.DataSourceReport.Tables[0];
        //    dataSource1.AcceptChanges();
        //    for (int index = 0; index < dataSource1.Rows.Count; ++index)
        //    {
        //        if ((bool)dataSource1.Rows[index]["chon"])
        //        {
        //            string cmdText = string.Format("delete [cosxlsx-dmdmvtct]  WHERE ma_ky = '" + dataSource1.Rows[index]["ma_ky"] + "' and ma_sp = '" + dataSource1.Rows[index]["ma_sp"] + "' and ma_px ='" + dataSource1.Rows[index]["ma_px"] + "' and ma_vt ='" + dataSource1.Rows[index]["ma_vt"] + "'");
        //            StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(cmdText));
        //        }
        //    }

        //    /*int visibleIndex = StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.VisibleIndex;
        //    StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
        //    StartUp.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
        //   {
        //       if (StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().Count<DataRecord>() > visibleIndex)
        //       {
        //           StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().ElementAt<DataRecord>(visibleIndex - 1);
        //       }
        //       else
        //       {
        //           StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
        //           StartUp.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
        //     {
        //         if (StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord == null || StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord.VisibleIndex != -1)
        //             return;
        //         StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)StartUp.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
        //     }));
        //       }
        //   }));*/
        //    ReloadData(StartUp.sMa_ky_loc);

        //}

        private static void oBrowseCt_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F2)
                StartUp.V_Xem();
            else if (e.Key == Key.F3)
                StartUp.V_Sua();
            else if (e.Key == Key.F4)
            {
                StartUp.V_Moi();
            }
            else
            {
                if (e.Key != Key.F8)
                    return;
                StartUp.V_Xoa();
            }
        }

        private static void oBrowse_F8(object sender, EventArgs e)
        {
            StartUp.V_Xoa();
        }

        private static void oBrowse_F6(object sender, EventArgs e)
        {
            StartUp.V_DoimaVT();
        }

        private static void oBrowse_F3(object sender, EventArgs e)
        {
            StartUp.V_Sua();
        }

        private static void oBrowse_F4(object sender, EventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                StartUp.V_Copy();
            }
            else
            {
                if (Keyboard.Modifiers != ModifierKeys.None)
                    return;
                StartUp.V_Moi();
            }
        }

        public static void oBrowse_Esc(object sender, EventArgs e)
        {
        }

        private static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
        }

        private static void btnCopyky_Click(object sender, RoutedEventArgs e)
        {
            StartUp.V_Copyky();
        }

        private static void oBrowse_F5(object sender, EventArgs e)
        {
            StartUp.V_Copyky();
        }
        public static void V_Copyky()
        {
            codmnvlLoc2 frm = new codmnvlLoc2();
            frm.ShowDialog();
            //Thực hiện lấy dữ liệu
            if (frm.isClose)
            {
                try
                {
                    if (StartUp.ma_ky_truoc.Trim() != StartUp.sMa_ky_loc.Trim())
                    {
                        FrmWaiting waiting = new FrmWaiting(600);
                        waiting.Show();
                        waiting.Set(100);
                        using (SqlCommand sqlcmd = new SqlCommand("EXEC [dbo].[COSXLSX-CODMNVL_F5] @ma_bpht0, @So_lsx0, @Ma_sp0, @Ma_hd0,@Ma_ky0, @Ma_ky"))
                        {
                            sqlcmd.Parameters.Add("@ma_bpht0", SqlDbType.VarChar).Value = StartUp.ma_bpht_truoc;
                            sqlcmd.Parameters.Add("@So_lsx0", SqlDbType.VarChar).Value = StartUp.so_lsx_truoc;
                            sqlcmd.Parameters.Add("@Ma_sp0", SqlDbType.VarChar).Value = StartUp.Ma_sp_truoc;
                            sqlcmd.Parameters.Add("@Ma_hd0", SqlDbType.VarChar).Value = StartUp.ma_hd_truoc;
                            sqlcmd.Parameters.Add("@Ma_ky0", SqlDbType.VarChar).Value = StartUp.ma_ky_truoc;
                            sqlcmd.Parameters.Add("@Ma_ky", SqlDbType.VarChar).Value = StartUp.sMa_ky_loc;
                            StartUp.SasObj.ExcuteNonQuery(sqlcmd);
                        }
                        waiting.Set(300);
                        //Hiện thị lại dữ liệu
                        if (frm.isClose)
                        {
                            StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
                            ReloadData(StartUp.sMa_ky_loc);
                            string message = "Sao chép định mức thành công !";
                            int num1 = (int)ExMessageBox.Show(StartupBase.SasObj, message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);

                        }
                        waiting.Set(600);
                        waiting.Close();
                    }
                }
                catch (Exception ex)
                {
                    SasErrorLib.ErrorLog.CatchMessage(ex);
                }
            }
        }


        public static void btnDoimaVT_Click(object sender, RoutedEventArgs e)
        {
            StartUp.V_DoimaVT();

        }

        public static void V_DoimaVT()
        {
            FrmTaoHDA frmTaoHDA = new FrmTaoHDA();
            frmTaoHDA.ShowDialog();
            if (!frmTaoHDA.isOk)
                return;
            string text = frmTaoHDA.txtMa_qs_pt.Text;
            StartUp.oBrowse.DataGrid.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            DataTable dataSource1 = StartUp.DataSourceReport.Tables[1];
            //DataView dataSource2 = StartUp.DataSourceReport.Tables[1].DefaultView;
            dataSource1.AcceptChanges();
            string str = "";
            for (int index = 0; index < dataSource1.Rows.Count; ++index)
            {
                if ((bool)dataSource1.Rows[index]["chon"])
                {
                    string user_id = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                    string time = DateTime.Now.ToString("HH:mm:ss").ToString();
                    object date = DateTime.Now.ToString("yyyy-MM-dd");
                    
                    string cmdText = string.Format("UPDATE [cosxlsx-dmdmvtct] SET ma_vt = '" + text + "' " +
                        ", date2 = '"+date+"'" +
                        ", time2 = '"+time+"'" +
                        ", user_id2 = '"+user_id+"'"+
                        " WHERE ma_ky = '" + dataSource1.Rows[index]["ma_ky"] + "' and ma_sp = '" + dataSource1.Rows[index]["ma_sp"] + "' and ma_px ='" + dataSource1.Rows[index]["ma_px"] + "' and ma_vt ='" + dataSource1.Rows[index]["ma_vt"] + "'");
                    StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(cmdText));
                    //str = str + "," + dataSource1.Rows[index]["tag"]+dataSource1.Rows[index]["ma_vt"];
                }
            }
            //str = str.Remove(0, 1);

            //string cmdText = string.Format("UPDATE [cosxlsx-dmdmvtct] SET ma_vt = '{0}' WHERE dbo.InList(ma_sp+ma_ky+ma_px+ma_vt, '{1}', ',') = 1", text, (object)str, "");
            //StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(cmdText));

            int num = (int)ExMessageBox.Show(1123, StartupBase.SasObj, "Chương trình đã thực hiện xong!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
            ReloadData(StartUp.sMa_ky_loc);

        }

        public static void btnmauExcel_Click(object sender, RoutedEventArgs e)
        {
            StartUp.V_mauExcel();

        }


        public static void V_mauExcel()
        {
            StartupBase.SasObj.SynchroFile(".\\Excel-Mau", "BOMIMEX".ToString().Trim() + ".xls");
            string sourceFileName = StartupBase.SasObj.M_StartUp_Path + "Excel-Mau\\" + "BOMIMEX".ToString().Trim() + ".xls";
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 2003 (.xls)|*.xls";
            string text = "";
            if (saveFileDialog.ShowDialog() == true)
            {
                text = saveFileDialog.FileName;
                try
                {
                    File.Copy(sourceFileName, text, overwrite: true);
                }
                catch (Exception ex)
                {
                    ExMessageBox.Show(610, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return;
                }
            }
            if (File.Exists(text))
            {
                Process.Start(text);
            }

        }

        public static void btnimportExcel_Click(object sender, RoutedEventArgs e)
        {
            StartUp.V_importExcel();

        }

        public ImportInfo Info
        {
            get;
            private set;
        }

        public static void V_importExcel()
        {
            //string currentDirectory = Environment.CurrentDirectory;
            //string text = txtFileName;
            //string fileName = "";
            //if (text.Trim() != "")
            //{
            //    text = Path.GetDirectoryName(txtFileName);
            //    if (text.Trim() != "" && Directory.Exists(text))
            //    {
            //        Environment.CurrentDirectory = text;
            //    }
            //    fileName = Path.GetFileName(txtFileName);
            //}
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.FileName = fileName;
            //openFileDialog.Filter = "Microsoft excel|*.xls;*.xlsx";
            //if (openFileDialog.ShowDialog() == true)
            //{
            //    txtFileName = openFileDialog.FileName;
            //    Ten_file = openFileDialog.SafeFileName;


            //    Directory.CreateDirectory(Environment.GetEnvironmentVariable("Temp") + "\\ExcelImport\\");
            //    string text1 = Environment.GetEnvironmentVariable("Temp") + "\\ExcelImport\\" + Ten_file.Trim();
            //    string text2 = txtFileName;
            //    if (File.Exists(text1))
            //    {
            //        FileInfo fileInfo = new FileInfo(text1);
            //        if (fileInfo.IsReadOnly)
            //        {
            //            fileInfo.IsReadOnly = false;
            //        }
            //    }
            //    File.Copy(text2, text1, overwrite: true);
            //    ImportInfo info = default(ImportInfo);
            //    info.Name = "Danh mục định mức NVL".ToString();
            //    info.FileName = text1;
            //    info.TableTemplate = "DMDMNVLIMEX".ToString();
            //    info.ExcelTemplate = "DMDMNVLIMEX".ToString();
            //    info.Ma_Imex = "IN07".ToString();
            //    info.PostProc = "smimexdmdmnvl".ToString();

            //}
            //Environment.CurrentDirectory = currentDirectory;
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "xlsx (*.xlsx)|*.xlsx|xls (*.xls)|*.xls";
                string filepath = "";
                if (openFileDialog.ShowDialog() == true)
                {
                    filepath = openFileDialog.FileName;
                    if (File.Exists(filepath))
                    {
                        DataSet b_ds = new DataSet();
                        Workbook workbook = new Workbook(filepath);
                        Worksheet worksheet = workbook.Worksheets[0]; ;

                        b_ds.Tables.Add(worksheet.Cells.ExportDataTableAsString(1, 0, worksheet.Cells.MaxRow + 1, worksheet.Cells.MaxColumn + 1, true));
                        DataTable dtDataSave = b_ds.Tables[0].Clone();
                        foreach (DataRow row in b_ds.Tables[0].Rows)
                        {
                            var rows = row;
                            var isRow = TrimRow(ref rows);
                            if (!isRow)
                            {
                                continue;
                            }
                            dtDataSave.ImportRow(row);
                        }

                        ////foreach (DataRow r in dtDataSave.Rows)
                        ////{
                        ////    SqlCommand sqlcmd = new SqlCommand();
                        ////    sqlcmd.CommandText = "Select ma_ky,ma_sp From [cosxlsx-Dmdmvt] Where ma_sp = @Masp and ma_ky = @Maky";
                        ////    sqlcmd.Parameters.Add("@Masp", SqlDbType.Char).Value = (object)r["ma_sp"];
                        ////    sqlcmd.Parameters.Add("@Maky", SqlDbType.Char).Value = (object)r["ma_ky"];
                        ////    if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                        ////    {
                        ////        dtDataSave.Rows[index].Delete();
                        ////    }
                        ////}
                        string y = "";

                        //for (int index = 0; index < dtDataSave.Rows.Count; ++index)
                        //{
                        //    SqlCommand sqlcmd = new SqlCommand();
                        //    sqlcmd.CommandText = "Select ma_ky,ma_sp From [cosxlsx-Dmdmvt] Where ma_sp = @Masp and ma_ky = @Maky and ma_px = @Mapx";
                        //    sqlcmd.Parameters.Add("@Masp", SqlDbType.Char).Value = dtDataSave.Rows[index]["ma_sp"].ToString().Trim();
                        //    sqlcmd.Parameters.Add("@Maky", SqlDbType.Char).Value = dtDataSave.Rows[index]["ma_ky"].ToString().Trim();
                        //    sqlcmd.Parameters.Add("@Mapx", SqlDbType.Char).Value = dtDataSave.Rows[index]["ma_px"].ToString().Trim();
                        //    DataTable tblData = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
                        //    if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count > 0)
                        //    {
                        //        y = y + "\nkỳ:" + dtDataSave.Rows[index]["ma_ky"].ToString().Trim() + "  -   Sản phẩm:" + dtDataSave.Rows[index]["ma_sp"].ToString().Trim() + "  -   Chuyền:" + dtDataSave.Rows[index]["ma_px"].ToString().Trim() + ";";
                        //        dtDataSave.Rows[index]["ma_ky"] = "";
                        //        //int num = (int)ExMessageBox.Show(2630, StartupBase.SasObj, "dữ liệu bị hủy do có bản ghi đã tồn tại trong chương trình!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        //        //return;
                        //    }
                        //}
                        //for (int index0 = 0; index0 < dtDataSave.Rows.Count; ++index0)
                        //{
                        //    string str2 = dtDataSave.Rows[index0]["ma_ky"].ToString().Trim();
                        //    if (str2.Equals(""))
                        //    {
                        //        dtDataSave.Rows[index0].Delete();
                        //    }
                        //}
                        //dtDataSave.AcceptChanges();




                        //foreach (DataRow r in dtDataSave.Rows)
                        //{
                        //    object x = (object)r["so_luong"].ToString().Trim() != "" ? (object)r["so_luong"].ToString().Trim().Replace(".", "") : "0";
                        //    SqlCommand sqlcmd = new SqlCommand();
                        //    sqlcmd.CommandText = "insert into [impor_DINHMUC]" +
                        //           "(ma_ky, ma_sp, ma_vt, ma_px, sl_dm, ghi_chu,ma_hd,ma_bpht,so_lsx)" +
                        //    " values(@ma_ky, @ma_sp, @ma_vt, @ma_px, @sl_dm, @ghi_chu,@ma_hd,@ma_bpht,@so_lsx)";
                        //    sqlcmd.Parameters.Add("@ma_ky", SqlDbType.Char).Value = (object)r["ma_ky"];
                        //    sqlcmd.Parameters.Add("@ma_sp", SqlDbType.Char).Value = (object)r["ma_sp"];
                        //    sqlcmd.Parameters.Add("@ma_vt", SqlDbType.Char).Value = (object)r["ma_vt"];
                        //    sqlcmd.Parameters.Add("@ma_px", SqlDbType.Char).Value = (object)r["ma_px"];
                        //    sqlcmd.Parameters.Add("@ma_hd", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@ma_bpht", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@so_lsx", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@sl_dm", SqlDbType.Decimal).Value = (object)x.ToString().Replace(",", ".");
                        //    sqlcmd.Parameters.Add("@ghi_chu", SqlDbType.NVarChar).Value = (object)r["ghi_chu"];
                        //    StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                        //}


                        SqlParameter param = new SqlParameter();
                        param.ParameterName = "@table_dinhmuc";
                        param.SqlDbType = SqlDbType.Structured;
                        param.Value = dtDataSave;
                        param.Direction = ParameterDirection.Input;

                        SqlCommand sqlCmd = new SqlCommand("usp_Insert_dinhmuc");
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.Add(param);
                        //StartupBase.SasObj.ExcuteNonQuery(sqlCmd);
                        DataTable tblData10 = StartupBase.SasObj.ExcuteReader(sqlCmd).Tables[0].Copy();


                        //int z = dtDataSave.Rows.Count;
                        //for (int index = 0; index < dtDataSave.Rows.Count-1; ++index)
                        //{
                        //    string str1 = dtDataSave.Rows[index]["ma_ky"].ToString().Trim();
                        //    string str1_ = dtDataSave.Rows[index]["ma_sp"].ToString().Trim();
                        //    string str1__ = dtDataSave.Rows[index]["ma_px"].ToString().Trim();

                        //    for (int index0 = index + 1; index0 < dtDataSave.Rows.Count; ++index0)
                        //    {
                        //        string str2 = dtDataSave.Rows[index0]["ma_ky"].ToString().Trim();
                        //        string str2_ = dtDataSave.Rows[index0]["ma_sp"].ToString().Trim();
                        //        string str2__ = dtDataSave.Rows[index0]["ma_px"].ToString().Trim();
                        //        if (str1.Equals(str2) && str1_.Equals(str2_) && str1__.Equals(str2__))
                        //        {
                        //            //dtDataSave.Rows[index0].Delete();
                        //            dtDataSave.Rows[index0]["ma_ky"] = "";
                        //        }
                        //    }
                        //    //int index0 = index + 1;
                        //    //while (index0 < dtDataSave.Rows.Count)
                        //    //{
                        //    //    string str2 = dtDataSave.Rows[index0]["ma_ky"].ToString().Trim();
                        //    //    string str2_ = dtDataSave.Rows[index0]["ma_sp"].ToString().Trim();
                        //    //    string str2__ = dtDataSave.Rows[index0]["ma_px"].ToString().Trim();
                        //    //    if (str1.Equals(str2) && str1_.Equals(str2_) && str1__.Equals(str2__))
                        //    //    {
                        //    //        dtDataSave.Rows[index0].Delete();
                        //    //        dtDataSave.AcceptChanges();
                        //    //    }
                        //    //    else
                        //    //    {
                        //    //        index0++;
                        //    //    }
                        //    //}

                        //}
                        //for (int index0 = 0; index0 < dtDataSave.Rows.Count; ++index0)
                        //{
                        //    string str2 = dtDataSave.Rows[index0]["ma_ky"].ToString().Trim();
                        //    if (str2.Equals(""))
                        //    {
                        //        dtDataSave.Rows[index0].Delete();
                        //    }
                        //}
                        //    dtDataSave.AcceptChanges();

                        //foreach (DataRow r in dtDataSave.Rows)
                        //{
                        //    SqlCommand sqlcmd = new SqlCommand();
                        //    sqlcmd.CommandText = "insert into [cosxlsx-Dmdmvt]" +
                        //        "(ma_ky, ma_sp, ma_px, date0, time0, user_id0, date2, time2, user_id2, loai_cp,ma_hd,ma_bpht,so_lsx)" +
                        // " values(@ma_ky, @ma_sp, @ma_px, GETDATE(), CONVERT(VARCHAR(8), GETDATE(), 114), @user_id, GETDATE(), CONVERT(VARCHAR(8), GETDATE(), 114), @user_id, 0,@ma_hd,@ma_bpht,@so_lsx)";
                        //    sqlcmd.Parameters.Add("@ma_ky", SqlDbType.Char).Value = (object)r["ma_ky"];
                        //    sqlcmd.Parameters.Add("@ma_sp", SqlDbType.Char).Value = (object)r["ma_sp"];
                        //    sqlcmd.Parameters.Add("@ma_px", SqlDbType.Char).Value = (object)r["ma_px"];
                        //    sqlcmd.Parameters.Add("@ma_hd", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@ma_bpht", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@so_lsx", SqlDbType.Char).Value = "";
                        //    sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                        //    StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                        //}


                        SqlCommand sqlcmd1 = new SqlCommand();
                        sqlcmd1.CommandText = "exec [impor_DINHMUC_PR] @user_id";
                        sqlcmd1.Parameters.Add("@user_id", SqlDbType.Int).Value = StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                        //StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                        DataTable tblData = StartupBase.SasObj.ExcuteReader(sqlcmd1).Tables[0].Copy();

                        if (tblData.Rows.Count > 0)
                        {
                            foreach (DataRow r in tblData.Rows)
                            {
                                y = y + "\nKỳ: " + (object)r["ma_ky"].ToString().Trim() + "  - Sản phẩm: " + (object)r["ma_sp"].ToString().Trim() + "  - Chuyền: " + (object)r["ma_px"].ToString().Trim() + "- VT: " + (object)r["ma_vt"].ToString().Trim() + ";";
                            }

                            string message = "Lỗi: định mức trùng mã vật tư: " + y;
                            int num1 = (int)ExMessageBox.Show(StartupBase.SasObj, message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                        else
                        {
                            string message = "Import định mức thành công !";
                            int num1 = (int)ExMessageBox.Show(StartupBase.SasObj, message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        }
                    }
                    else
                    {
                        string message = "Không tìm thấy file tại đường dẫn: " + filepath;
                        int num = (int)ExMessageBox.Show(1912, StartupBase.SasObj, message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    }
                    ReloadData(StartUp.sMa_ky_loc);
                    string message1 = "Import định mức thành công !";
                    int num10 = (int)ExMessageBox.Show(StartupBase.SasObj, message1, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        public static bool TrimRow(ref DataRow row)
        {
            string value = null;
            string colName = null;
            bool isRow = false;
            try
            {
                DataTable dtData = row.Table;
                foreach (DataColumn col in dtData.Columns)
                {
                    colName = col.ColumnName;
                    value = row[colName].ToString().Trim();
                    row[colName] = value;
                    if (value != "#N/A" & !string.IsNullOrEmpty(value))
                    {
                        isRow = true;
                    }
                }
                return isRow;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
