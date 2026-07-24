using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.Data;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace Poctpnc
{
    public partial class FrmView : Form
    {
        public bool isOk;
        public DataSet dsPn;
        private string stt_rec;

        public FrmView(string filter)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.dsPn = StartUp.GetPN(filter);
            DataColumn chooseColumn = new DataColumn("choose",typeof(bool));
            chooseColumn.DefaultValue = (object)false;
            this.dsPn.Tables[0].Columns.Add(chooseColumn);
            this.GrdBrowse.DataSource = (IEnumerable)this.dsPn.Tables[0].DefaultView;
            this.GrdBrowseCt.DataSource = (IEnumerable)this.dsPn.Tables[1].DefaultView;
            string strBrowse1 = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[2];
            this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, strBrowse1, this.dsPn.Tables[0]));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, strBrowse1);
            string strBrowse2 = StartUpTrans.CommandInfo[StartUpTrans.M_LAN.Equals("V") ? "Vbrowse2" : "Ebrowse2"].ToString().Split('|')[3];
            this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2, this.dsPn.Tables[1]));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2);
            this.grdMain.pnlButton.Height = 25.0;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.GrdBrowse.Focus();
            this.isOk = false;
            if (this.GrdBrowse.Records.Count <= 0)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GrdBrowse.ActiveRecord = (Record)(this.GrdBrowse.Records[0] as DataRecord)));
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
            this.dsPn.Tables[0].DefaultView.RowFilter += "" + string.Format("{0} = '{1}'", (object)" stt_rec ", (object)this.stt_rec);
            this.isOk = true;
            this.Close();
        }

        private void GrdBrowse_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            try
            {
                BasicGridView basicGridView = sender as BasicGridView;
                if (basicGridView.ActiveRecord == null || (basicGridView.ActiveRecord.Index < 0 || basicGridView.ActiveRecord.RecordType != RecordType.DataRecord))
                    return;
                this.dsPn.Tables[1].DefaultView.RowFilter = "";
                string str = "";
                this.stt_rec = ((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString();
                this.dsPn.Tables[1].DefaultView.RowFilter += str + string.Format("{0} = '{1}'", (object)" stt_rec ", (object)this.stt_rec);
                this.GrdBrowseCt.DataSource = (IEnumerable)this.dsPn.Tables[1].DefaultView;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

    }
}
