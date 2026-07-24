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

namespace Poctpna
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
            this.txtNgay_ct_old.Value = StartUpTrans.DsTrans.Tables[0].Rows[FrmPoctpna.iRow]["ngay_ct"];
            this.txtNgay_ct_new.Value = (object)DateTime.Now.Date;
            this.txtNgay_ct_new.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.txtNgay_ct_new.dValue == new DateTime() || string.IsNullOrEmpty(this.txtNgay_ct_new.Text.Trim().ToString()))
            {
                int num = (int)ExMessageBox.Show(515, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            else
            {
                FrmPoctpnaCopy.ngay_ct = DateTime.Parse(this.txtNgay_ct_new.Value.ToString());
                if (!SysFunc.CheckValidNgayKs(StartupBase.SasObj, new DateTime?(FrmPoctpnaCopy.ngay_ct.Date)))
                {
                    int num = (int)ExMessageBox.Show(520, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày khóa sổ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else if (!this.txtNgay_ct_new.IsValueValid)
                {
                    int num = (int)ExMessageBox.Show(525, StartupBase.SasObj, "Ngày chứng từ mới không hợp lệ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtNgay_ct_new.Focus();
                }
                else if (Convert.ToDateTime(this.txtNgay_ct_new.Value) < NgayTC.GetStartDate(StartUp.M_ngay_ct0))
                {
                    int num = (int)ExMessageBox.Show(530, StartupBase.SasObj, "Ngày chứng từ mới phải sau ngày mở sổ!", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.OK, MessageBoxImage.Asterisk);
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
