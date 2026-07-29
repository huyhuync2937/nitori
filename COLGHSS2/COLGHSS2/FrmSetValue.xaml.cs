using SasControls;
using SasFormBrowes;
using System;
using System.Windows;
using static Infragistics.Windows.Helpers.ActionHistory;

namespace COLGHSS2
{
    public partial class FrmSetValue : Form
    {
        public string title_ = "Cập nhật giá trị";

        public FrmSetValue(string OldValue,string field)
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            if(field.Trim() == "ma_may")
            {
                this.txtNewValue.Text = OldValue;
                SysFunc.LoadIcon((Window)this);
                this.txtNewValue.IsFocus = true;
            }
            else if (field.Trim() == "ma_ca" || field.Trim() == "ma_ca_kt")
            {
                this.txtNewValue.Text = OldValue;
                SysFunc.LoadIcon((Window)this);
                this.txtNewValue.IsFocus = true;
            }
            else if (field.Trim() == "sl_kg_lenh" || field.Trim() == "sl_cai_lenh")
            {
                this.txtsl.Text = OldValue;
                SysFunc.LoadIcon((Window)this);
                this.txtsl.IsFocus = true;
            }
            else
            {
                DateTime dt;
                string result = "";
                if (DateTime.TryParse(OldValue, out dt))
                {
                     result = dt.ToString("dd-MM-yyyy");
                }
                this.txtNgay_bd.Value = result;
                SysFunc.LoadIcon((Window)this);
                this.txtNgay_bd.Focus() ;
            }
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.txtNewValue.SearchInit();
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
        private void txtNgay_bd_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_bd.Value == DBNull.Value)
                this.txtNgay_bd.Value = (object)DateTime.Now;
            if (this.txtNgay_bd.IsFocusWithin && !( !(this.txtNgay_bd.dValue != new DateTime())))
                return;
        }

        private void cbField_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (this.cbField.Value.ToString().Trim() == "ma_ts")
                this.txtNewValue.ListID = "dmts";
            if (this.cbField.Value.ToString().Trim() == "ma_ca")
                this.txtNewValue.ListID = "dmca";

            this.txtNewValue.SearchInit();
            this.txtNewValue_PreviewLostFocus(null, null);
        }

        private void txtNewValue_PreviewLostFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            if(this.txtNewValue.RowResult != null)
            {
                if(this.txtNewValue.ListID == "dmts")
                {
                    this.txtTen_bp.Text = this.txtNewValue.RowResult["ten_ts"].ToString().Trim() ;
                        
                }
                if (this.txtNewValue.ListID == "dmca")
                {
                    this.txtTen_bp.Text = this.txtNewValue.RowResult["ten_ca"].ToString().Trim();

                }

            }
            else
            {
                this.txtTen_bp.Text = "";
            }
        }
    }
}
