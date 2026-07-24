using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace CACTPC1
{
    public partial class COTKTH2Dvcs : Form
    {
        public COTKTH2Dvcs()
        {
            this.InitializeComponent();
    
           // this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           //{
           //    if (StartupBase.M_LAN != "V")
           //    {
           //        Field field = this.GrdCt.FieldLayouts[0].Fields["ten_vt"];
           //        field.Visibility = Visibility.Hidden;
           //        field.Width = new FieldLength?(new FieldLength(0.0));
           //    }
           //    else
           //    {
           //        Field field = this.GrdCt.FieldLayouts[0].Fields["ten_vt2"];
           //        field.Visibility = Visibility.Hidden;
           //        field.Width = new FieldLength?(new FieldLength(0.0));
           //    }
           //}), DispatcherPriority.Background);
            SysFunc.LoadIcon((Window)this);
            if (!(StartupBase.M_LAN != "V"))
                return;
            this.Title = "FX list";
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

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.DialogResult = new bool?(true);
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
