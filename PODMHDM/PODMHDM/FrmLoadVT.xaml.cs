using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using System.Data.SqlClient;
using SasVoucherLib;

namespace PODMHDM
{
    public partial class FrmLoadVT : Form
    {
        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }
        public string MaVT { get; set; }
        public string Status { get; set; }

        public FrmLoadVT()
        {
            this.LanguageID = "PODMHDM_VT";
            this.InitializeComponent();
        }

        public FrmLoadVT(DateTime fromDate, DateTime toDate, string maVT, string status) : this()
        {
            this.FromDate = fromDate;
            this.ToDate = toDate;
            this.MaVT = maVT;
            this.Status = status;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            StartUp.isOk = false;
            SysFunc.LoadIcon((Window)this);
            string prog = "";

            string filter = " 1=1 ";
            if (!string.IsNullOrEmpty(this.MaVT))
                filter += " and ma_vt = ''" + this.MaVT.Trim().Replace("'", "''") + "''";
            if (!string.IsNullOrEmpty(this.Status))
            {
                if (this.Status.Equals("1"))
                {
                    prog = "COSLDHT";
                }
                else
                {
                    prog = "COSLDHN";
                }
            }
                //filter += " and status = ''" + this.Status.Trim() + "''";

            string sql = "Exec " + prog + " '" + this.FromDate.ToString("yyyy-MM-dd") + "', '" + this.ToDate.ToString("yyyy-MM-dd")+ "'";
            StartUp.HDBData = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql));
            if (StartUp.HDBData.Tables[0].DefaultView.Count < 1)
            {
                int num2 = (int)ExMessageBox.Show(3280, StartupBase.SasObj, "Không tìm thấy vật tư nào!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.Close();
                return;
            }
            this.GrdBrowse.DataSource = (IEnumerable)StartUp.HDBData.Tables[0].DefaultView;
            //this.GrdBrowseCt.DataSource = (IEnumerable)StartUp.HDBData.Tables[1].DefaultView;
            string[] strArray = StartUpTrans.CommandInfo["Vbrowse1"].ToString().Split('|');
            if (strArray.Length > 0)
            {
                this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, "ma_vt:80:h=Mã vật tư", StartUp.HDBData.Tables[0]));
                SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, "ma_vt:80:h=Mã vật tư");
                //this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strArray[1], StartUp.HDBData.Tables[1]));
                //SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, StartUp.stringBrowse2);
                //this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                //{
                //    if (this.GrdBrowse.Records.Count <= 0)
                //        return;
                //    this.GrdBrowse.ActiveRecord = this.GrdBrowse.Records[0];
                //}));
            }
        }

        private void GrdBrowse_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            try
            {
                BasicGridView basicGridView = sender as BasicGridView;
                if (basicGridView.ActiveRecord == null || (basicGridView.ActiveRecord.Index < 0 || basicGridView.ActiveRecord.RecordType != RecordType.DataRecord))
                    return;
                StartUp.HDBData.Tables[1].DefaultView.RowFilter = "";
                string str = "1 =1 ";
                StartUp.HDBData.Tables[1].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object)"" : (object)" and ", (object)"stt_rec", (object)((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
                this.GrdBrowseCt.DataSource = (IEnumerable)StartUp.HDBData.Tables[1].DefaultView;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.GrdBrowse.ActiveRecord == null || this.GrdBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
                return;
            string str = "1 =1 ";
            StartUp.HDBData.Tables[0].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object)"" : (object)" and ", (object)"stt_rec", (object)((this.GrdBrowse.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
            StartUp.isOk = true;
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            StartUp.isOk = false;
            this.Close();
        }
    }
}
