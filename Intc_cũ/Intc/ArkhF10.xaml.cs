using SasControls;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Intc
{
    public partial class ArkhF10 : FormFilter
    {
        private bool isOK = false;

        public ArkhF10()
        {
            this.InitializeComponent();
            this.txtSapXep.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.isOK = true;
            this.Close();
        }

        public bool ShowDialog()
        {
            base.ShowDialog();
            return this.isOK;
        }

        private void txt_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }

        private void FormFilter_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtSapXep.Value = (object)"1";
            this.txtSapXep.Focus();
        }

    }
}
