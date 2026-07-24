using SasControls;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Threading;

namespace FxTinhKH
{
    public class StartUp : StartupBase
    {
        public static string Ma_nt0 = "VND";
        public static string Language = "V";
        public static string KeyFilter = "";
        private int nYear = DateTime.Now.Year - 1;
        private const int RF_PROCESSMESSAGE = 41251;
        private const int RF_PROCESSWAITINGSHOW = 41254;
        private const int RF_PROCESSWAITING = 41253;
        private DataRow drCommandInfo;
        private DateTime _ngay_ct1;
        private DateTime _ngay_ct2;
        private DateTime _ngay_gia_px;
        public static DateTime M_ngay_ct0;
        public static DateTime M_ngay_ks;

        public override void Run()
        {
            StartupBase.Namespace = "FxTinhKH";
            this.Show(StartupBase.Menu_Id);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int SendMessage(IntPtr hwnd, [MarshalAs(UnmanagedType.U4)] int Msg, IntPtr wParam, IntPtr lParam);

        private void Show(string id)
        {
            this.drCommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, id);
            StartUp.Ma_nt0 = (string)StartupBase.SasObj.GetOption("M_MA_NT0");
            StartUp.Language = StartupBase.SasObj.GetOption("M_LAN").ToString();
            if (this.drCommandInfo == null || this.drCommandInfo.ItemArray.Length == 0)
                return;
            StartUp.M_ngay_ct0 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
            StartUp.M_ngay_ks = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KS");
            DateTime.TryParse(StartupBase.SasObj.GetSysvar("M_NGAY_CT1").ToString(), out this._ngay_ct1);
            DateTime.TryParse(StartupBase.SasObj.GetSysvar("M_NGAY_CT2").ToString(), out this._ngay_ct2);
            Debug.WriteLine(!DateTime.TryParse(StartupBase.SasObj.GetOption("M_NGAY_GIA_PX").ToString(), out this._ngay_gia_px) ? (object)null : (object)this._ngay_gia_px, "Ngay");
            Debug.WriteLine((object)this._ngay_gia_px, "ngay gia");
            DataTable dataTable = new DataTable("Filter");
            dataTable.Columns.Add("ky1", typeof(int));
            dataTable.Columns.Add("ky2", typeof(int));
            dataTable.Columns.Add("nam1", typeof(int));
            dataTable.Columns.Add("ma_cc");
            DateTime today = DateTime.Today;
            int num = DateTime.Today.Year;
            int num2 = this._ngay_ct1.Month;
            int num3 = this._ngay_ct2.Month;
            DateTime dateTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime dateTime2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            try
            {
                object sysVar = StartupBase.SasObj.GetSysvar("M_THANG1");
                num2 = ((sysVar.ToString() == "") ? DateTime.Now.Month : ((int)Convert.ToInt16(sysVar)));
                object sysVar2 = StartupBase.SasObj.GetSysvar("M_THANG2");
                num3 = ((sysVar2.ToString() == "") ? DateTime.Now.Month : ((int)Convert.ToInt16(sysVar2)));
                dateTime = new DateTime(DateTime.Now.Year, num2, 1);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            dataTable.Rows.Add(new object[] { num2, num3, num, "" });
            Ingia_tbLoc ingiaTbLoc = new Ingia_tbLoc();
            ingiaTbLoc.ShowInTaskbar = true;
            ingiaTbLoc.DisplayLanguage = StartUp.Language;
            ingiaTbLoc.DataContext = (object)dataTable;
            ingiaTbLoc.BindingSasObj = StartupBase.SasObj;
            ingiaTbLoc.Procedure = this.drCommandInfo["store_proc"].ToString();
            ingiaTbLoc.Title = SysFunc.Cat_Dau(this.drCommandInfo[StartupBase.M_LAN.Equals("V") ? "bar" : "bar2"].ToString());
            SysFunc.LoadIcon((Window)ingiaTbLoc);
            ingiaTbLoc.ShowDialog();
            if (ingiaTbLoc.bResult)
            {
                int num6 = 1;
                int num4 = 1;
                int num5 = 2017;
                try
                {
                    num6 = (int)dataTable.Rows[0]["ky1"];
                    num4 = (int)dataTable.Rows[0]["ky2"];
                    num5 = (int)dataTable.Rows[0]["nam1"];
                }
                catch (Exception ex)
                {
                }
                SqlCommand sqlcmd = new SqlCommand(this.drCommandInfo["store_proc"].ToString());
                sqlcmd.CommandType = CommandType.StoredProcedure;
                sqlcmd.Parameters.Add("@nam", SqlDbType.Int).Value = (object)num5;
                sqlcmd.Parameters.Add("@ky1", SqlDbType.Int).Value = (object)num6;
                sqlcmd.Parameters.Add("@ky2", SqlDbType.Int).Value = (object)num4;
                sqlcmd.Parameters.Add("@key", SqlDbType.VarChar).Value = (object)StartUp.KeyFilter;
                sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)(int)StartupBase.SasObj.UserInfo.Rows[0]["user_id"];
                sqlcmd.Parameters.Add("@Ma_dvcs", SqlDbType.VarChar).Value = (object)ingiaTbLoc.M_MA_DVCS.ToString().Trim();
                DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
                string strBrowse = !(StartupBase.M_LAN == "V") ? this.drCommandInfo["Ebrowse1"].ToString() : this.drCommandInfo["Vbrowse1"].ToString();
                SasFormBrowes.FormBrowse oBrowse = new SasFormBrowes.FormBrowse(StartupBase.SasObj, dataSet.Tables[dataSet.Tables.Count - 1].DefaultView, strBrowse);
                int num7;
                oBrowse.frmBrw.Loaded += (RoutedEventHandler)((s, e) => oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => num7 = (int)ExMessageBox.Show(1540, StartupBase.SasObj, "Chương trình đã thực hiện xong!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk))));
                oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? this.drCommandInfo["bar"].ToString() : this.drCommandInfo["bar2"].ToString());
                oBrowse.frmBrw.LanguageID = "FXTinhKHBrowse";
                oBrowse.ShowDialog();
                ingiaTbLoc.Close();
                if (Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    return;
                Application.Current.Shutdown();
            }
            else
            {
                ingiaTbLoc.Close();
                if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
                    Application.Current.Shutdown();
            }
        }
    }
}
