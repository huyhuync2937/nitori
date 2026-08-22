using SasControls;
using SasFormBrowes;
using System;
using System.Windows;
using static Infragistics.Windows.Helpers.ActionHistory;

namespace COSLDHN
{
    public partial class FrmSetValue : Form
    {
        public string title_ = "Cập nhật giá trị";

        public FrmSetValue(string OldValue,string field)
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
           
             this.txtsl.Text = OldValue;
                SysFunc.LoadIcon((Window)this);
                this.txtsl.IsFocus = true;
           
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            //this.txtNewValue.SearchInit();
            //if( this.txtNgay_bd.Value == null)
            //{
            //    int num = (int)ExMessageBox.Show(9050, StartupBase.SasObj, "Chưa nhập giá trị ngày.!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    return;
            //} 
            this.DialogResult = new bool?(true);
            this.Close();
        }
        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.DialogResult = new bool?(false);
            this.Close();
        }
       
    }
}
