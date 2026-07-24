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
    public partial class FrmLoaddonhang : Form
    {
        public FrmLoaddonhang(DateTime tungay, DateTime denngay, string filter)
        {
            this.LanguageID = "PODMHDM_DH";
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            string sql = "Exec " + StartUp.storeproc + " '" + tungay.ToString("yyyy-MM-dd") + "', '" + denngay.ToString("yyyy-MM-dd") + "', '" + filter + "'";
            StartUp.HDBData = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql));
            if (StartUp.HDBData.Tables[0].DefaultView.Count < 1)
            {
                int num2 = (int)ExMessageBox.Show(3280, StartupBase.SasObj, "Không tìm thấy đơn hàng nào!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.Close();
            }
            this.GrdBrowse.DataSource = (IEnumerable)StartUp.HDBData.Tables[0].DefaultView;
            this.GrdBrowseCt.DataSource = (IEnumerable)StartUp.HDBData.Tables[1].DefaultView;
            string[] strArray = StartUpTrans.CommandInfo["Vbrowse1"].ToString().Split('|');
            if(strArray.Length>0)
            {
                this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, strArray[0], StartUp.HDBData.Tables[0]));
                SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, StartUp.stringBrowse1);
                this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strArray[1], StartUp.HDBData.Tables[1]));
                SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, StartUp.stringBrowse2);
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                {
                    if (this.GrdBrowse.Records.Count <= 0)
                        return;
                    this.GrdBrowse.ActiveRecord = this.GrdBrowse.Records[0];
                }));
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            StartUp.isOk = false;
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
