using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Windows;
using System.Windows.Input;

namespace COSXKSX.KSXW
{
    public partial class FrmPoctpnaCopy : Form
    {
        public static bool isCopy = false;
        public static DateTime ngay_ct;

        public FrmPoctpnaCopy()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FrmCopy_Loaded);
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmCopy_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct_old.Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ngay_ksx"];
            this.txtNgay_ct_new.Value = (object)DateTime.Now.Date;
            this.txtNgay_ct_new.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct_new.dValue == new DateTime() || string.IsNullOrEmpty(this.txtNgay_ct_new.Text.Trim().ToString()))
            {
                int num = (int)ExMessageBox.Show(2315, StartupBase.SasObj, "Ngày lệnh sản xuất mới không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            else
            {
                FrmPoctpnaCopy.ngay_ct = DateTime.Parse(this.txtNgay_ct_new.Value.ToString());
                if (!this.txtNgay_ct_new.IsValueValid)
                {
                    int num = (int)ExMessageBox.Show(2330, StartupBase.SasObj, "Ngày lệnh sản xuất không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else
                {
                    FrmPoctpnaCopy.isCopy = true;
                    this.Close();
                }
            }
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            FrmPoctpnaCopy.isCopy = false;
            this.Close();
        }

        private void txtNgay_ct_new_LostFocus(object sender, RoutedEventArgs e)
        {
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            FrmPoctpnaCopy.isCopy = false;
            this.Close();
        }
    }
}
