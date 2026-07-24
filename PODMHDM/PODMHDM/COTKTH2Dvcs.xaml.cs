using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using SasVoucherLib;

namespace PODMHDM
{
    public partial class COTKTH2Dvcs : Form
    {
        public COTKTH2Dvcs()
        {


            this.LanguageID = "PODMHDM_DH";
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            SysFunc.LoadIcon((Window)this);
            if (!(StartupBase.M_LAN != "V"))
                return;
            this.Title = "FX list";
            //string sql = "Exec " + StartUp.storeproc + " '" + tungay.ToString("yyyy-MM-dd") + "', '" + denngay.ToString("yyyy-MM-dd") + "', '" + filter + "'";
            //StartUp.HDBData = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql));
            //if (StartUp.HDBData.Tables[0].DefaultView.Count < 1)
            //{
            //    int num2 = (int)ExMessageBox.Show(3280, StartupBase.SasObj, "Không tìm thấy đơn hàng nào!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.Close();
            //}
            //this.GrdBrowse.DataSource = (IEnumerable)StartUp.HDBData.Tables[0].DefaultView;
            //this.GrdBrowseCt.DataSource = (IEnumerable)StartUp.HDBData.Tables[1].DefaultView;
            //string[] strArray = StartUpTrans.CommandInfo["Vbrowse1"].ToString().Split('|');
            //if (strArray.Length > 0)
            //{
            //    this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, strArray[0], StartUp.HDBData.Tables[0]));
            //    SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, StartUp.stringBrowse1);
            //    this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strArray[1], StartUp.HDBData.Tables[1]));
            //    SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, StartUp.stringBrowse2);
            //    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
            //    {
            //        if (this.GrdBrowse.Records.Count <= 0)
            //            return;
            //        this.GrdBrowse.ActiveRecord = this.GrdBrowse.Records[0];
            //    }));
            //}
        }

        private void allBox_Loaded(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               if (this.GrdCt.Records.Count > 0)
                   this.GrdCt.ActiveRecord = this.GrdCt.Records[0];
               this.GrdCt.Focus();
           }), DispatcherPriority.Background);
        }

        private void GrdCt_KeyUp(object sender, KeyEventArgs e)
        {
            XamDataGrid grid = sender as XamDataGrid;
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (e.Key == Key.A)
                    this.SelectAll(grid, true);
                if (e.Key == Key.U)
                    this.SelectAll(grid, false);
            }
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Space)
                return;
            this.SelectEntry(grid);
        }

        private void SelectEntry(XamDataGrid grid)
        {
            if (!(grid.ActiveRecord is DataRecord activeRecord) || activeRecord.RecordType != RecordType.DataRecord)
                return;
            DataRowView dataItem = activeRecord.DataItem as DataRowView;
            bool flag = !(bool)dataItem["tag"];
            dataItem["tag"] = (object)flag;
        }

        private void SelectAll(XamDataGrid grid, bool tag)
        {
            for (int index = 0; index < grid.Records.Count; ++index)
                ((grid.Records[index] as DataRecord).DataItem as DataRowView)[nameof(tag)] = (object)tag;
        }

        private void chkSelect_Click(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
            DataView dataSource = this.GrdCt.DataSource as DataView;
            bool? isChecked = checkBox.IsChecked;
            bool flag = isChecked.GetValueOrDefault() && isChecked.HasValue;
            foreach (DataRowView dataRowView in dataSource)
                dataRowView["tag"] = (object)flag;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            StartUp.isOk = false;
        }
        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
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
