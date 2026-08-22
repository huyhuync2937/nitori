using SasControls;
using SasFormReport;
using SasLib;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace SasVoucherLib
{
    /// <summary>Interaction logic for Glgsdkct.xaml</summary>
    /// <summary>VoucherContextMenuDel</summary>
    public partial class VoucherContextMenuDel : FormFilter
    {
        public bool bIsOK;

        public VoucherContextMenuDel(SasObject _SasObj)
        {
            this.InitializeComponent();
            this.BindingSasObj = _SasObj;
        }

        private void this_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.BindingSasObj != null)
            {
                this.txtMaDVCS.SearchInit();
                if (this.BindingSasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
                {
                    if (this.txtMaDVCS.RowResult != null)
                        this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
                }
                else if (this.txtMaDVCS.RowResult != null)
                    this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
            }
            this.txtSo_ct1.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.CheckValid())
                return;
            this.bIsOK = true;
            this.Close();
        }

        private bool CheckValid()
        {
            bool flag = true;
            Convert.ToDateTime(this.BindingSasObj.GetSysvar("M_NGAY_KY1"));
            if (flag && (this.txtNgay_ct1.Value == null || this.txtNgay_ct1.Value.ToString() == ""))
            {
                int num = (int)ExMessageBox.Show(-1145, this.BindingSasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag = false;
                this.txtNgay_ct1.Focus();
            }
            if (flag && !this.txtNgay_ct1.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(-1150, this.BindingSasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag = false;
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
            }
            DateTime dateTime1 = Convert.ToDateTime(this.txtNgay_ct1.Value);
            if (flag && (this.txtNgay_ct2.Value == null || this.txtNgay_ct2.Value.ToString() == ""))
            {
                int num = (int)ExMessageBox.Show(-1160, this.BindingSasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag = false;
                this.txtNgay_ct2.Focus();
            }
            if (flag && !this.txtNgay_ct2.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(-1165, this.BindingSasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag = false;
                this.txtNgay_ct2.Focus();
                this.txtNgay_ct2.SelectAll();
            }
            DateTime dateTime2 = Convert.ToDateTime(this.txtNgay_ct2.Value);
            if (flag && dateTime1 > dateTime2)
            {
                int num = (int)ExMessageBox.Show(-1175, this.BindingSasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                flag = false;
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
            }
            return flag;
        }

        private void txtMaDVCS_LostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.lblTenDVCS.Text = "";
            if (this.BindingSasObj.GetOption("M_LAN").ToString().ToUpper().Equals("V"))
            {
                if (this.txtMaDVCS.RowResult == null)
                    return;
                this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs"].ToString();
            }
            else
            {
                if (this.txtMaDVCS.RowResult == null)
                    return;
                this.lblTenDVCS.Text = this.txtMaDVCS.RowResult["ten_dvcs2"].ToString();
            }
        }
    }
}
