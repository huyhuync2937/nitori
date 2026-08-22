using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace Poctpna
{
    public partial class FrmPoctpnaGetHDM : Form
    {
        public bool isOk = false;
        public FrmView frm;
        public DataSet dsHdm;

        public FrmPoctpnaGetHDM()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.isOk = false;
            this.Close();
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct_old.Focus();
            this.txtMa_kh.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.CheckValid())
                return;
            this.Hide();
            this.frm = new FrmView(this.GetFilter());
            if (StartUpTrans.M_LAN != "V")
                this.frm.Title = "Contract list";
            this.frm.ShowDialog();
            this.isOk = this.frm.isOk;
            this.dsHdm = this.frm.dsHdm;
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            this.Close();
        }

        public bool CheckValid()
        {
            bool flag = true;
            if (!this.txtNgay_ct_old.IsValueValid)
            {
                flag = false;
                int num = (int)ExMessageBox.Show(570, StartupBase.SasObj, "Ngày bắt đầu không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_old.Focus();
            }
            if (!this.txtNgay_ct_new.IsValueValid)
            {
                flag = false;
                int num = (int)ExMessageBox.Show(575, StartupBase.SasObj, "Ngày kết thúc không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct_new.Focus();
            }
            return flag;
        }

        private void txtMa_kh_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.tblTen_kh.Text = "";
            if (this.txtMa_kh.RowResult == null)
                return;
            this.tblTen_kh.Text = !StartUpTrans.M_LAN.ToUpper().Equals("V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
        }

        private string GetFilter()
        {
            string str = "1=1";
            if (this.txtNgay_ct_old.dValue != new DateTime())
                str = str + " AND ngay_ct >= '" + string.Format("{0:yyyyMMdd}", (object)this.txtNgay_ct_old.dValue) + "'";
            if (this.txtNgay_ct_new.dValue != new DateTime())
                str = str + " AND ngay_ct <= '" + string.Format("{0:yyyyMMdd}", (object)this.txtNgay_ct_new.dValue) + "'";
            if (!string.IsNullOrEmpty(this.txtMa_kh.Text.Trim()))
                str = str + " AND ma_kh LIKE '" + this.txtMa_kh.Text.Trim() + "%'";
            if (!string.IsNullOrEmpty(this.txtma_hdm.Text.Trim()))
                str = str + " AND ma_hdm LIKE '%" + this.txtma_hdm.Text.Trim() + "%'";          
            return str + " AND status != '0' and status != '5' and status != '6' and status != '7' and status != '8' And isnull(status2,'') <> '2'";
        }
    }
}
