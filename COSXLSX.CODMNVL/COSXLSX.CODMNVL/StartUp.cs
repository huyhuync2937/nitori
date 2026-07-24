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
        private static SasFormBrowes.FormBrowse2 oBrowse;
        public static DateTime M_ngay_ct0;
        public static codmnvlLoc _frmLoc;
        public static string sso_lsx_loc;
        public static string sMa_sp_loc;
        public static string sMa_bpht_loc;
        public static string sMa_hd_loc;
        public static DateTime? sNgay1_loc;
        public static DateTime? sNgay2_loc;
        public static string sMa_ky_loc;

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
                if (isFirstLoad)
                {
                    StartUp.CommandInfo["store_proc"].ToString().Split('|');
                    StartUp.cmd = new SqlCommand(StartUp.CommandInfo["store_proc"].ToString());
                    StartUp.cmd.CommandType = CommandType.StoredProcedure;
                    StartUp.cmd.Parameters.Add("@Ma_ky", SqlDbType.VarChar).Value = ma_ky;
                    StartUp.cmd.Parameters.Add("@so_lsx", SqlDbType.VarChar).Value = so_lsx;
                    StartUp.cmd.Parameters.Add("@Ma_sp", SqlDbType.VarChar).Value = ma_sp;
                    StartUp.cmd.Parameters.Add("@Ma_bpht", SqlDbType.VarChar).Value = ma_bpht;
                    StartUp.cmd.Parameters.Add("@Ma_hd", SqlDbType.VarChar).Value = ma_hd;
                    DataSet dataSet = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                    DataTable table1 = dataSet.Tables[0].Copy();
                    DataTable table2 = dataSet.Tables[1].Copy();
                    table1.TableName = "tbMain";
                    table2.TableName = "tbDetail";
                    StartUp.DataSourceReport.Tables.Add(table1);
                    StartUp.DataSourceReport.Tables.Add(table2);
                    DataRelation relation = new DataRelation("tag_Relation", table1.Columns["tag"], table2.Columns["tag"], false);
                    StartUp.DataSourceReport.Relations.Add(relation);
                    string[] strArray = StartUp.CommandInfo["VBrowse1"].ToString().Trim().Split('|');
                    if (StartupBase.M_LAN == "E")
                        strArray = StartUp.CommandInfo["EBrowse1"].ToString().Trim().Split('|');
                    string strBrowse = strArray[0];
                    string strBrowseCt = strArray[1];
                    StartUp.oBrowse = new SasFormBrowes.FormBrowse2(StartupBase.SasObj, table1.DefaultView, table2.DefaultView, strBrowse, strBrowseCt, "tag");
                    StartUp.oBrowse.Esc += new SasFormBrowes.FormBrowse2.GridKeyUp_Esc(StartUp.oBrowse_Esc);
                    StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse2.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                    StartUp.oBrowse.F4 += new SasFormBrowes.FormBrowse2.GridKeyUp_F4(StartUp.oBrowse_F4);
                    StartUp.oBrowse.F8 += new SasFormBrowes.FormBrowse2.GridKeyUp_F8(StartUp.oBrowse_F8);
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
                    if (StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport") is ToolBar name)
                    {
                        name.Items.Insert(0, (object)toolBarButton1);
                        name.Items.Insert(1, (object)toolBarButton2);
                        name.Items.Insert(2, (object)toolBarButton3);
                        name.Items.Insert(3, (object)toolBarButton4);
                        name.Items.Insert(4, (object)toolBarButton5);
                        name.Items.Insert(5, (object)toolBarButton6);
                    }
                }
                else
                {
                    StartUp.DataSourceReport.Relations.Remove("tag_Relation");
                    StartUp.DataSourceReport.Tables.Remove("tbMain");
                    StartUp.DataSourceReport.Tables.Remove("tbDetail");
                    DataTable table1 = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[0].Copy();
                    table1.TableName = "tbMain";
                    StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)table1.DefaultView;
                    StartUp.oBrowse.ObrowseView = table1.DefaultView;
                    StartUp.DataSourceReport.Tables.Add(table1);
                    DataTable table2 = StartupBase.SasObj.ExcuteReader(StartUp.cmd).Tables[1].Copy();
                    table2.TableName = "tbDetail";
                    StartUp.oBrowse.frmBrw.oBrowseCt.DataSource = (IEnumerable)table2.DefaultView;
                    StartUp.oBrowse.ObrowseViewCt = table2.DefaultView;
                    StartUp.DataSourceReport.Tables.Add(table2);
                    DataRelation relation = new DataRelation("tag_Relation", table1.Columns["tag"], table2.Columns["tag"], false);
                    StartUp.DataSourceReport.Relations.Add(relation);
                }
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
            else
            {
                if (e.Key != Key.F8)
                    return;
                StartUp.V_Xoa();
            }
        }

        private static void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (StartUp.oBrowse.ActiveRecord == null)
                return;
            StartUp.V_Copy();
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
            int count = StartUp.DataSourceReport.Tables[0].DefaultView.Count;
            int index = 0;
            FrmCapNhat frmCapNhat = new FrmCapNhat(empty, ssoLsxLoc1, ssoLsxLoc2, sMaSpLoc, sMahdLoc);
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Them dinh muc nguyen vat lieu" : "Add bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk || (count == StartUp.DataSourceReport.Tables[0].DefaultView.Count || StartUp.oBrowse == null))
                return;
            index = 0;
            DataRowView data = (DataRowView)null;
            if (StartUp.DataSourceReport.Tables[0].DefaultView.Count > 0)
            {
                for (int index1 = 0; index1 < StartUp.DataSourceReport.Tables[0].DefaultView.Count; ++index1)
                {
                    if (frmCapNhat.txtSo_lsx.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["so_lsx"].ToString().Trim()) && frmCapNhat.txtMa_bpht.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_bpht"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_sp"].ToString().Trim()))
                    {
                        index = index1;
                        data = StartUp.DataSourceReport.Tables[0].DefaultView[index1];
                        break;
                    }
                }
                StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (data != null)
                       StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord);
                   else
                       StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = StartUp.oBrowse.frmBrw.oBrowse.Records.FirstOrDefault<Record>((Func<Record, bool>)(x => (x as DataRecord).DataItem == data));
               }));
            }
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
            int count = StartUp.DataSourceReport.Tables[0].DefaultView.Count;
            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["so_lsx"].Value.ToString() + '_' + activeRecord1.Cells["ma_bpht"].Value.ToString() + '_' + activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_hd"].Value.ToString(), activeRecord1.Cells["so_lsx"].Value.ToString(), activeRecord1.Cells["ma_bpht"].Value.ToString(), activeRecord1.Cells["ma_sp"].Value.ToString(), activeRecord1.Cells["ma_hd"].Value.ToString());
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Them dinh muc nguyen vat lieu" : "Add bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk || (count == StartUp.DataSourceReport.Tables[0].DefaultView.Count || StartUp.oBrowse == null))
                return;
            DataRowView data = (DataRowView)null;
            if (StartUp.DataSourceReport.Tables[0].DefaultView.Count > 0)
            {
                for (int index1 = 0; index1 < StartUp.DataSourceReport.Tables[0].DefaultView.Count; ++index1)
                {
                    if (frmCapNhat.txtSo_lsx.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["so_lsx"].ToString().Trim()) && frmCapNhat.txtMa_bpht.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_bpht"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_sp"].ToString().Trim()))
                    {
                        index = index1;
                        data = StartUp.DataSourceReport.Tables[0].DefaultView[index1];
                        break;
                    }
                }
                StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (data != null)
                       StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord);
                   else
                       StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = StartUp.oBrowse.frmBrw.oBrowse.Records.FirstOrDefault<Record>((Func<Record, bool>)(x => (x as DataRecord).DataItem == data));
               }));
            }
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
            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["so_lsx"].Value.ToString() + '_' + activeRecord1.Cells["ma_bpht"].Value.ToString() + '_' + activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_hd"].Value.ToString(), activeRecord1.Cells["so_lsx"].Value.ToString(), activeRecord1.Cells["ma_bpht"].Value.ToString(), activeRecord1.Cells["ma_sp"].Value.ToString(), activeRecord1.Cells["ma_hd"].Value.ToString());
            frmCapNhat.Title = StartupBase.M_LAN.Equals("V") ? "Sua dinh muc nguyen vat lieu" : "Edit bill of material entry";
            frmCapNhat.EnableEditMode(true);
            frmCapNhat.ShowDialog();
            if (!frmCapNhat.IsOk)
                return;
            index = 0;
            if (StartUp.DataSourceReport.Tables[0].DefaultView.Count > 0)
            {
                for (int index1 = 0; index1 < StartUp.DataSourceReport.Tables[0].DefaultView.Count; ++index1)
                {
                    if (frmCapNhat.txtSo_lsx.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["so_lsx"].ToString().Trim()) && frmCapNhat.txtMa_bpht.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_bpht"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_sp"].ToString().Trim()) && frmCapNhat.txtMa_sp.Text.Trim().ToString().Equals(StartUp.DataSourceReport.Tables[0].DefaultView[index1]["ma_hd"].ToString().Trim()))
                    {
                        index = index1;
                        break;
                    }
                }
                StartUp.oBrowse.frmBrw.oBrowse.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => StartUp.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)(StartUp.oBrowse.frmBrw.oBrowse.Records[index] as DataRecord)));
            }
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
            FrmCapNhat frmCapNhat = new FrmCapNhat(activeRecord1.Cells["so_lsx"].Value.ToString() + '_' + activeRecord1.Cells["ma_bpht"].Value.ToString() + '_' + activeRecord1.Cells["ma_sp"].Value.ToString() + '_' + activeRecord1.Cells["ma_hd"].Value.ToString(), activeRecord1.Cells["so_lsx"].Value.ToString(), activeRecord1.Cells["ma_bpht"].Value.ToString(), activeRecord1.Cells["ma_sp"].Value.ToString(), activeRecord1.Cells["ma_hd"].Value.ToString());
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
        }

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
        public static void V_Copyky()
        {
            codmnvlLoc2 frm = new codmnvlLoc2();
            frm.ShowDialog();
            //Thực hiện lấy dữ liệu
            if (!frm.isClose)
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
                        if (!frm.isClose)
                        {
                            StartUp.CallGridVouchers(false, StartUp.sMa_ky_loc, StartUp.sso_lsx_loc, StartUp.sMa_sp_loc, StartUp.sMa_bpht_loc, StartUp.sMa_hd_loc);
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
    }
}
