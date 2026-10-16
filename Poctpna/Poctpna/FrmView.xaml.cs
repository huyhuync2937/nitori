using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;

namespace Poctpna
{
    public partial class FrmView : Form
    {
        public bool isOk;
        public DataSet dsHdm;

        public FrmView(string filter)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.dsHdm = StartUp.GetHdm(filter);
            this.GrdBrowse.DataSource = (IEnumerable)this.dsHdm.Tables[0].DefaultView;
            this.GrdBrowseCt.DataSource = (IEnumerable)this.dsHdm.Tables[2].DefaultView;
            string strBrowse1;
            if (StartUpTrans.M_LAN.Equals("V"))
                strBrowse1 = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|')[2];
            else
                strBrowse1 = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|')[2];
            this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, strBrowse1, this.dsHdm.Tables[0]));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, strBrowse1);
            string strBrowse2;
            if (StartUpTrans.M_LAN.Equals("V"))
                strBrowse2 = StartUpTrans.CommandInfo["Vbrowse2"].ToString().Split('|')[3];
            else
                strBrowse2 = StartUpTrans.CommandInfo["Ebrowse2"].ToString().Split('|')[3];
            this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2, this.dsHdm.Tables[2]));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            if (this.GrdBrowse.Records.Count <= 0)
                return;
            this.GrdBrowse.ActiveRecord = (Record)(this.GrdBrowse.Records[0] as DataRecord);
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            // Ctrl + A: chọn tất cả vật tư theo đơn hàng hiện tại trên GrdBrowseCt
            if (e.Key.Equals((object)Key.A) && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                this.ChonTatCaVatTu(true);
                e.Handled = true;
                return;
            }
            // Ctrl + U: hủy chọn tất cả vật tư theo đơn hàng hiện tại trên GrdBrowseCt
            if (e.Key.Equals((object)Key.U) && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                this.ChonTatCaVatTu(false);
                e.Handled = true;
                return;
            }
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }

        private void ChonTatCaVatTu(bool chon)
        {
            try
            {
                DataView detailView = this.dsHdm.Tables[2].DefaultView;
                if (detailView.Count <= 0)
                    return;
                object chonValue = detailView.Table.Columns["chon"].DataType == typeof(bool) ? (object)chon : (object)(chon ? 1 : 0);
                foreach (DataRowView drv in detailView)
                    drv["chon"] = chonValue;
                // Đánh dấu đơn hàng hiện tại để các vật tư được lấy khi nhấn OK
                if (this.GrdBrowse.ActiveRecord != null && this.GrdBrowse.ActiveRecord.RecordType == RecordType.DataRecord)
                {
                    DataRowView header = (this.GrdBrowse.ActiveRecord as DataRecord).DataItem as DataRowView;
                    if (header != null)
                        header["chon"] = this.dsHdm.Tables[0].Columns["chon"].DataType == typeof(bool) ? (object)chon : (object)(chon ? 1 : 0);
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.GrdBrowse.ActiveRecord == null || this.GrdBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
                return;
            string str = "1 = 1 ";
            //this.dsHdm.Tables[0].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object)"" : (object)" and ", (object)"stt_rec", (object)((this.GrdBrowse.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
            this.dsHdm.Tables[0].DefaultView.RowFilter =
    (string.IsNullOrEmpty(str) ? "" : str + " AND ") + "chon = 1";
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
                this.dsHdm.Tables[2].DefaultView.RowFilter = "";
                string str = "1 = 1 ";
                this.dsHdm.Tables[2].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object)"" : (object)" and ", (object)" stt_rec ", (object)((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
                this.GrdBrowseCt.DataSource = (IEnumerable)this.dsHdm.Tables[2].DefaultView;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
    }
}
