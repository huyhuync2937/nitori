using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace POCNCC1
{
    public class StartUp : StartupBase
    {
        public static string TableName = "ph101";
        public static string TableViewName = "v_ph101";
        public static DataSet DataSource = new DataSet();
        public static string SqlTableKey = "stt_rec";
        public static string SqlTableObjectName = "";
        private static SqlCommand cmd = new SqlCommand();

        private static SasFormBrowes.FormBrowse2 oBrowse;
        private static DataRow CommandInfo;
        public static DateTime M_ngay_ct0;
        public static string M_ma_nt0;
        private static FormLoc _frmLoc;
        public static string Trangthai = "";
        public static string Soyeucau = "";
        public static bool IsOk = false;
        public override void Run()
        {
            StartupBase.Namespace = "POCNCC1";
            try
            {
                DateTime now1 = DateTime.Now;
                StartUp.CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                StartUp.M_ngay_ct0 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
                if (StartUp.CommandInfo == null)
                    return;
                StartUp.TableName = StartUp.CommandInfo["parameter"].ToString().Split('|')[0];
                StartUp.TableViewName = StartUp.CommandInfo["parameter"].ToString().Split('|')[1];
                StartUp._frmLoc = new FormLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
                DateTime now2 = DateTime.Now;
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        public static void CallGridVouchers(
          object StartDate,
          object EndDate,
          int KindReport,
          string filterpbg)
        {
            try
            {
                string strArray1 = StartUp.CommandInfo["store_proc"].ToString();
                StartUp.cmd = new SqlCommand();
                StartUp.cmd.CommandText = "Exec " + strArray1 + " @StartDate , @EndDate , @Condition";
                StartUp.cmd.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = string.IsNullOrEmpty(StartDate.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)StartDate);
                StartUp.cmd.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = string.IsNullOrEmpty(EndDate.ToString()) ? (object)"" : (object)string.Format("{0:yyyyMMdd}", (object)(DateTime)EndDate);
                StartUp.cmd.Parameters.Add("@Condition", SqlDbType.NVarChar).Value = (object)filterpbg;
                DataSource = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                string[] strArray2;
                string[] strArray3;
                if (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
                {
                    strArray2 = StartUp.CommandInfo["VBrowse1"].ToString().Trim().Split('|');
                    strArray3 = StartUp.CommandInfo["VBrowse2"].ToString().Trim().Split('|');
                }
                else
                {
                    strArray2 = StartUp.CommandInfo["EBrowse1"].ToString().Trim().Split('|');
                    strArray3 = StartUp.CommandInfo["EBrowse2"].ToString().Trim().Split('|');
                }
                string strBrowse1 = strArray2[0];
                string strBrowseCt1 = strArray2[1];
                string strBrowse2 = strArray3[0];
                string strBrowseCt2 = strArray3[1];
                StartUp.oBrowse = KindReport != 0 ? new SasFormBrowes.FormBrowse2(StartupBase.SasObj, StartUp.DataSource.Tables[0].DefaultView, StartUp.DataSource.Tables[1].DefaultView, strBrowse2, strBrowseCt2, "stt_rec") : new SasFormBrowes.FormBrowse2(StartupBase.SasObj, StartUp.DataSource.Tables[0].DefaultView, StartUp.DataSource.Tables[1].DefaultView, strBrowse1, strBrowseCt1, "stt_rec");
                StartUp.oBrowse.CTRL_R += new SasFormBrowes.FormBrowse2.GridKeyUp_CTRL_R(StartUp.oBrowse_CTRL_R);
                StartUp.oBrowse.frmBrw.oBrowse.FieldSettings.AllowEdit = new bool?(false);
                StartUp.oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.CommandInfo["bar"].ToString() : StartUp.CommandInfo["bar2"].ToString());
                StartUp.oBrowse.Esc += new SasFormBrowes.FormBrowse2.GridKeyUp_Esc(StartUp.oBrowse_Esc);
                StartUp.oBrowse.frmBrw.PreviewKeyDown += new KeyEventHandler(StartUp.frmBrw_PreviewKeyDown);
                object name = StartUp.oBrowse.frmBrw.ToolBar.FindName("tbReport");
                if (StartUp.Trangthai.Equals("1"))
                {
                    StartUp.oBrowse.F3 += new SasFormBrowes.FormBrowse2.GridKeyUp_F3(oBrowse_F3);
                    if (name != null)
                    {
                        ToolBar toolBar = name as ToolBar;
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        ToolBarButton toolBarButton1 = new ToolBarButton();
                        toolBarButton1.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton1.Name = "btnF2";
                        toolBarButton1.Text = (StartupBase.M_LAN.Equals("V") ? "Chọn yêu cầu mua hàng" : "Requset purchase");
                        toolBarButton1.ToolTip = (object)"F2";
                        toolBarButton1.ImagePath = "Images\\Preview.png";
                        toolBarButton1.Click += new RoutedEventHandler(oBrowse_F2);
                        toolBar.Items.Insert(1, toolBarButton1);

                        ToolBarButton toolBarButton = new ToolBarButton();
                        toolBarButton.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton.Name = "btnF3";
                        toolBarButton.Text = (StartupBase.M_LAN.Equals("V") ? "Duyệt" : "Aproval");
                        toolBarButton.ToolTip = (object)"F3";
                        toolBarButton.ImagePath = "Images\\Preview.png";
                        toolBarButton.Click += new RoutedEventHandler(oBrowse_F3);
                        toolBar.Items.Insert(2, toolBarButton);
                    }
                }
                else if (StartUp.Trangthai.Equals("2"))
                {
                    StartUp.oBrowse.F4 += new SasFormBrowes.FormBrowse2.GridKeyUp_F4(oBrowse_F4);
                    if (name != null)
                    {
                        ToolBar toolBar = name as ToolBar;
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        ToolBarButton toolBarButton1 = new ToolBarButton();
                        toolBarButton1.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton1.Name = "btnF4";
                        toolBarButton1.Text = (StartupBase.M_LAN.Equals("V") ? "Hủy" : "Cancel");
                        toolBarButton1.ToolTip = (object)"F4";
                        toolBarButton1.ImagePath = "Images\\Preview.png";
                        toolBarButton1.Click += new RoutedEventHandler(oBrowse_F4);
                        toolBar.Items.Insert(1, toolBarButton1);
                    }
                }
                else
                {
                    StartUp.oBrowse.F3 += new SasFormBrowes.FormBrowse2.GridKeyUp_F3(oBrowse_F3);
                    StartUp.oBrowse.F4 += new SasFormBrowes.FormBrowse2.GridKeyUp_F4(oBrowse_F4);
                    if (name != null)
                    {
                        ToolBar toolBar = name as ToolBar;
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        toolBar.Items.RemoveAt(1);
                        ToolBarButton toolBarButton = new ToolBarButton();
                        toolBarButton.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton.Name = "btnF3";
                        toolBarButton.Text = (StartupBase.M_LAN.Equals("V") ? "Duyệt" : "Aproval");
                        toolBarButton.ToolTip = (object)"F3";
                        toolBarButton.ImagePath = "Images\\Preview.png";
                        toolBarButton.Click += new RoutedEventHandler(oBrowse_F3);
                        toolBar.Items.Insert(1, toolBarButton);
                        ToolBarButton toolBarButton2 = new ToolBarButton();
                        toolBarButton2.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton2.Name = "btnF4";
                        toolBarButton2.Text = (StartupBase.M_LAN.Equals("V") ? "Hủy" : "Cancel");
                        toolBarButton2.ToolTip = (object)"F4";
                        toolBarButton2.ImagePath = "Images\\Preview.png";
                        toolBarButton2.Click += new RoutedEventHandler(oBrowse_F4);
                        toolBar.Items.Insert(2, toolBarButton2);

                        ToolBarButton toolBarButton1 = new ToolBarButton();
                        toolBarButton1.BorderBrush = (Brush)Brushes.Transparent;
                        toolBarButton1.Name = "btnF2";
                        toolBarButton1.Text = (StartupBase.M_LAN.Equals("V") ? "Chọn yêu cầu mua hàng" : "Requset purchase");
                        toolBarButton1.ToolTip = (object)"F2";
                        toolBarButton1.ImagePath = "Images\\Preview.png";
                        toolBarButton1.Click += new RoutedEventHandler(oBrowse_F2);
                        toolBar.Items.Insert(3, toolBarButton1);
                    }
                }
                StartUp.oBrowse.frmBrw.LanguageID = "POCNCC1_1";
                StartUp.oBrowse.ShowDialog();
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.InnerException.Message);
            }
        }

        private static void Updatetrangthai(string trang_thai)
        {
            string dongMessage = "";
            StartUp.oBrowse.DataGrid.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            DataView dataSource = StartUp.oBrowse.DataGrid.DataSource as DataView;
            dataSource.Table.AcceptChanges();

            // DataView dataSource = StartUp.DataSource.Tables[0].DefaultView;         
            DataTable datact = StartUp.DataSource.Tables[1];
            string str = string.Join(",", dataSource.Cast<DataRowView>().Where<DataRowView>((Func<DataRowView, bool>)(x =>
            {
                if (!(bool)x["tag"] || x["status2"].ToString().Equals("2"))
                {
                    if ((bool)x["tag"] && x["status2"].ToString().Equals("2"))
                    {
                        dongMessage += x["so_ct"].ToString().Trim() + ", ";
                    }
                    return false;
                }
                return true;
            })).Select<DataRowView, string>((Func<DataRowView, string>)(a => a["so_ct"].ToString().Trim())).ToArray<string>());


            string stt_rec = string.Join(",", dataSource.Cast<DataRowView>().Where<DataRowView>((Func<DataRowView, bool>)(x =>
            {
                if (!(bool)x["tag"] || x["status2"].ToString().Equals("2"))
                {
                    if ((bool)x["tag"] && x["status2"].ToString().Equals("2"))
                    {
                        dongMessage += x["stt_rec"].ToString().Trim() + ", ";
                    }
                    return false;
                }
                return true;
            })).Select<DataRowView, string>((Func<DataRowView, string>)(a => a["stt_rec"].ToString().Trim())).ToArray<string>());

            string[] arrStt_rec = stt_rec.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < arrStt_rec.Length; i++)
            {
                StartUp.Getdmgia_ncc(arrStt_rec[i].Trim());
            }

            if (!string.IsNullOrEmpty(str))
            {
                int user_id = (int)Convert.ToInt16(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                string cmdText = string.Format("UPDATE " + StartUp.TableName + " SET status={0}, nv_duyet = '{1}' WHERE dbo.InList(ma_hd, '{2}', ',') = 1", (object)trang_thai, (object)(trang_thai.Equals("2") ? user_id.ToString() : ""), (object)str);
                StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(cmdText));
                foreach (DataRowView view in dataSource)
                {
                    if ((bool)view["tag"] || !view["status2"].ToString().Equals("2"))
                    {
                        if(!string.IsNullOrEmpty(view["so_yeucau"].ToString()))
                        {
                            string sql = string.Format("Update " + StartUp.TableName + " Set so_yeucau = '{0}' where ma_hd ='{1}';", view["so_yeucau"].ToString().Trim(), view["ma_hd"].ToString().Trim());
                            StartupBase.SasObj.ExcuteNonQuery(new SqlCommand(sql));
                        }
                    }
                }
            }

            string filterpbg = StartUp._frmLoc.GetFilterPHieubaogia();
            StartUp.QueryData(false, StartUp._frmLoc.TxtStartDateTime.Value, StartUp._frmLoc.TxtEndDateTime.Value, int.Parse(StartUp._frmLoc.cbMauBaoCao2.Value.ToString()), filterpbg);
            int num = (int)ExMessageBox.Show(1123, StartupBase.SasObj, "Chương trình đã thực hiện xong!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            if (!string.IsNullOrEmpty(dongMessage))
            {
                MessageBox.Show("Không thể thay đổi trạng thái đơn hàng đã đóng: " + dongMessage, "Thông báo");
            }
            DataSource = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
            StartUp.DataSource.Tables[1].DefaultView.RowFilter = "stt_rec  ='" + StartUp.DataSource.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)StartUp.DataSource.Tables[0].DefaultView;
            StartUp.oBrowse.frmBrw.oBrowseCt.DataSource = (IEnumerable)StartUp.DataSource.Tables[1].DefaultView;
            StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
            StartUp.oBrowse.UpdateSumaryFields();
        }

        public static void Getdmgia_ncc (string stt_rec)
        {
            string format = "EXEC dbo.GetDmgia_ncc @stt_rec";
            SqlCommand sqlcmd = new SqlCommand(format);
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 16).Value = (object)stt_rec;
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
        }
        private static void frmBrw_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Space:
                    if (Keyboard.Modifiers != ModifierKeys.None || StartUp.oBrowse.ActiveRecord == null)
                        break;
                    e.Handled = true;
                    StartUp.Select();
                    break;
                case Key.A:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    StartUp.SelectAll(true);
                    break;
                case Key.U:
                    if (Keyboard.Modifiers != ModifierKeys.Control)
                        break;
                    StartUp.SelectAll(false);
                    break;
            }
        }
        private static void oBrowse_F2(object sender, EventArgs e)
        {
            if (StartUp.oBrowse.DataGrid.ActiveRecord == null || (StartUp.oBrowse.DataGrid.ActiveRecord is FilterRecord || StartUp.oBrowse.DataGrid.ActiveRecord is SummaryRecord))
                return;

            Purchase frmpur = new Purchase();
            string ma_hd = StartUp.oBrowse.ActiveRecord.Cells["so_yeucau"].Value.ToString();
            if (!string.IsNullOrEmpty(ma_hd))
                frmpur.txtSoyeucau.Filter = "ma_hd = '" + ma_hd.Trim() + "'";
            frmpur.ShowDialog();
            if (StartUp.IsOk)
            {
                if (!string.IsNullOrEmpty(StartUp.Soyeucau))
                {
                    StartUp.oBrowse.ActiveRecord.Cells["so_yeucau"].Value = StartUp.Soyeucau;
                }
            }
        }
        private static void oBrowse_F3(object sender, EventArgs e)
        {
            Updatetrangthai("2");

        }
        private static void oBrowse_F4(object sender, EventArgs e)
        {
            if (StartUp.TableName == "[DMHD]" || StartUp.TableName == "[dmhd]")
                Updatetrangthai("0");
            else
                Updatetrangthai("1");
        }
        private static void Select()
        {
            if (StartUp.oBrowse.DataGrid.ActiveRecord == null || (StartUp.oBrowse.DataGrid.ActiveRecord is FilterRecord || StartUp.oBrowse.DataGrid.ActiveRecord is SummaryRecord))
                return;
            Cell cell = StartUp.oBrowse.ActiveRecord.Cells["tag"];
            cell.Value = (object)!(bool)cell.Value;
        }

        private static void SelectAll(bool tag)
        {
            StartUp.oBrowse.DataGrid.ExecuteCommand(DataPresenterCommands.EndEditModeAndAcceptChanges);
            DataView dataSource = StartUp.oBrowse.DataGrid.DataSource as DataView;
            foreach (DataRowView dataRowView in dataSource)
                dataRowView[nameof(tag)] = (object)tag;
            dataSource.Table.AcceptChanges();
        }

        public static void QueryData(
          bool isFirstLoad,
          object StartDate,
          object EndDate,
          int KindReport,
          string filterpbg)
        {
            try
            {
                if (isFirstLoad)
                {
                    StartUp.CallGridVouchers(StartDate, EndDate, KindReport, filterpbg);
                }
                else
                {
                    DataSet dataSet = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
                    DataTable dataTable1 = dataSet.Tables[0].Copy();
                    dataTable1.TableName = "tbMain";
                    DataTable dataTable2 = dataSet.Tables[1].Copy();
                    dataTable2.TableName = "tbDetail";
                    dataTable2.DefaultView.RowFilter = "stt_rec  ='" + dataTable1.DefaultView[0]["stt_rec"].ToString() + "'";
                    StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable1.DefaultView;
                    StartUp.oBrowse.frmBrw.oBrowseCt.DataSource = (IEnumerable)dataTable2.DefaultView;
                    StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
                    StartUp.oBrowse.UpdateSumaryFields();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            string filterpbg = StartUp._frmLoc.GetFilterPHieubaogia();
            StartUp.QueryData(false, StartUp._frmLoc.TxtStartDateTime.Value, StartUp._frmLoc.TxtEndDateTime.Value, int.Parse(StartUp._frmLoc.cbMauBaoCao2.Value.ToString()), filterpbg);
        }

        private static void oBrowse_Esc(object sender, EventArgs e)
        {
            StartUp.oBrowse.frmBrw.Close();
            if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                return;
            Application.Current.Shutdown();
        }

        public static string GetTableShow(bool KindReport)
        {
            string empty = string.Empty;
            string upper = StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper();
            string str1 = "ngay_ct;Ma_ct0;so_ct;ma_kh;";
            string str2 = !upper.Equals("V") ? str1 + "ten_kh2;dien_giai;tk;tk_du;" : str1 + "ten_kh;dien_giai;tk;tk_du;";
            string str3 = !KindReport ? str2 + "ps_no_nt;ps_co_nt;ma_vv;ma_phi;" : str2 + "ps_no;ps_co;ma_vv;ma_phi;";
            return !upper.Equals("V") ? str3 + "ten_tk2;ten_tk2_du;ma_ct;ma_dvcs" : str3 + "ten_tk;ten_tk_du;ma_ct;ma_dvcs";
        }
    }
}
