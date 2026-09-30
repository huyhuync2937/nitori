using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
namespace POCNCC1
{
    /// <summary>
    /// Interaction logic for Arttkb1F3.xaml
    /// </summary>
    public partial class Purchase : FormFilter
    {

        public Purchase()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            txtSoyeucau.SearchInit();
            this.txtSoyeucau.IsFocus = true;
            StartUp.IsOk = false;
        }
        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            StartUp.IsOk = true;
            this.Close();
        }
        private void btnHuy_Click(object sender, RoutedEventArgs e)
        {
            StartUp.IsOk = false;
            this.Close();
        }
        private void txtSoyeucau_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtSoyeucau.Text))
            {
                if (this.txtSoyeucau == null)
                    return;
                StartUp.Soyeucau =this.txtSoyeucau.RowResult["ma_hd"].ToString();
            }
            else
                StartUp.Soyeucau = string.Empty;
        }
    }
}
