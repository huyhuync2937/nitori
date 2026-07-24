using Infragistics.Windows.Editors;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace COTD
{
    public partial class CotdLoc : FormFilter
    {
        public bool isClose = true;

        public CotdLoc()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmCotdLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtTd.Value = (object)StartUp._parameter;
            this.txtTd.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.isClose = false;
            this.Hide();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isClose = true;
            this.Close();
        }

    }
}
