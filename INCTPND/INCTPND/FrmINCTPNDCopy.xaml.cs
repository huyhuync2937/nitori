using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace INCTPND
{
    public partial class FrmINCTPNDCopy : Form
    {
        public static bool isCopy = false;
        public static DateTime ngay_ct;

        public FrmINCTPNDCopy()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FrmCopy_Loaded);
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmCopy_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct_old.Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmINCTPND.iRow]["ngay_ct"];
            this.txtNgay_ct_new.Value = (object)DateTime.Now.Date;
            this.txtNgay_ct_new.Focus();
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            FrmINCTPNDCopy.isCopy = false;
            this.Close();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct_new.dValue == new DateTime() || string.IsNullOrEmpty(this.txtNgay_ct_new.Text.Trim().ToString()))
            {
                int num = (int)ExMessageBox.Show(555, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            else
            {
                FrmINCTPNDCopy.ngay_ct = DateTime.Parse(this.txtNgay_ct_new.Value.ToString());
                if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(FrmINCTPNDCopy.ngay_ct.Date)))
                {
                    int num = (int)ExMessageBox.Show(565, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày khóa sổ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else if (!this.txtNgay_ct_new.IsValueValid)
                {
                    int num = (int)ExMessageBox.Show(570, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else if (Convert.ToDateTime(this.txtNgay_ct_new.Value) < NgayTC.GetStartDate(StartUp.M_ngay_ct0))
                {
                    int num = (int)ExMessageBox.Show(575, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày mở sổ!", "SASERP 20 .NET", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else
                {
                    FrmINCTPNDCopy.isCopy = true;
                    this.Close();
                }
            }
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            FrmINCTPNDCopy.isCopy = false;
            this.Close();
        }
    }
}
