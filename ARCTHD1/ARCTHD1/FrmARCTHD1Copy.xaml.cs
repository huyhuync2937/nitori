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

namespace ARCTHD1
{
    public partial class FrmARCTHD1Copy : Form
    {
        public bool isCopy = false;
        public static DateTime ngay_ct;

        public FrmARCTHD1Copy()
        {
            this.InitializeComponent();
            this.Loaded += new RoutedEventHandler(this.FrmCopy_Loaded);
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmCopy_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct_old.Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow]["ngay_ct"];
            this.txtNgay_ct_new.Value = (object)DateTime.Now.Date;
            this.txtNgay_ct_new.Focus();
        }

        private void btnNhan_Click(object sender, RoutedEventArgs e)
        {
            if (!this.txtNgay_ct_new.IsValueValid || this.txtNgay_ct_new.Value == DBNull.Value)
            {
                int num = (int)ExMessageBox.Show(450, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
                this.txtNgay_ct_new.SelectAll();
            }
            else if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct_new.dValue)))
            {
                int num = (int)ExMessageBox.Show(455, StartupBase.SasObj, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            else if (!SysFunc.CheckValidNgayMs(StartupBase.SasObj, new DateTime?(this.txtNgay_ct_new.dValue)))
            {
                int num = (int)ExMessageBox.Show(460, StartupBase.SasObj, "Ngày hạch toán phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            else
            {
                FrmARCTHD1Copy.ngay_ct = DateTime.Parse(this.txtNgay_ct_new.Value.ToString());
                this.isCopy = true;
                this.Close();
            }
        }

        private void btnHuy_Click(object sender, RoutedEventArgs e)
        {
            this.isCopy = false;
            this.Close();
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.isCopy = false;
            this.Close();
        }

    }
}
