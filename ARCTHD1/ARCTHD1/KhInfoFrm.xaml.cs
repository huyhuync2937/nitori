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

namespace ARCTHD1
{
    public partial class KhInfoFrm : Form
    {
        public bool IsAllowSave = true;

        public KhInfoFrm()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            DataRow row = StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow];
            if (row["ten_kh_thue"].ToString().Trim() == "")
                row["ten_kh_thue"] = (object)row["ten_kh"].ToString();
            this.txtTenkh.Text = row["ten_kh_thue"].ToString();
            this.txtDiaChi.Text = row["dia_chi"].ToString();
            this.txtMaSoThue.Text = row["ma_so_thue"].ToString();
            this.txtTenkh.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.IsAllowSave = true;
            if (!SysFunc.CheckSumMaSoThue(this.txtMaSoThue.Text.Trim()) && !string.IsNullOrEmpty(this.txtMaSoThue.Text.Trim()))
            {
                switch (StartUpTrans.M_MST_CHECK.Trim())
                {
                    case "2":
                        int num = (int)ExMessageBox.Show(510, StartupBase.SasObj, "Mã số thuế không hợp lệ, không lưu được!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        this.IsAllowSave = false;
                        break;
                }
            }
            if (!this.IsAllowSave)
                return;
            DataRow row = StartUpTrans.DsTrans.Tables[0].Rows[FrmArcthd1.iRow];
            row["ten_kh_thue"] = (object)this.txtTenkh.Text;
            row["dia_chi"] = (object)this.txtDiaChi.Text;
            row["ma_so_thue"] = (object)this.txtMaSoThue.Text;
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }
    }
}
