using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;

namespace COSXLSX.CODMNVL
{
    public partial class FrmTaoHDA : Form
    {
        public bool isOk = false;
        public int kind = 1;
        public string filterma_qs = "";
        public string So_ct_format { get; set; }
        public double So_ct_num { get; set; }
        public int So_ct_length { get; set; }
        public DataRowView[] Rows { get; set; }
        public FrmTaoHDA()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.So_ct_length = StartupBase.SasObj.GetDatabaseFieldLength("so_ct");
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMa_qs_pt.IsFocus = true;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtMa_qs_pt.Text.Trim()) || !this.txtMa_qs_pt.CheckLostFocus())
            {
                int num = (int)ExMessageBox.Show(697, StartupBase.SasObj, "Chưa vào mã vật tư", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtMa_qs_pt.IsFocus = true;
            }
            //else if (string.IsNullOrEmpty(this.txttk_co.Text.Trim()) || !this.txttk_co.CheckLostFocus())
            //{
            //    int num = (int)ExMessageBox.Show(698, StartupBase.SasObj, "Chưa vào tài khoản có", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txttk_co.IsFocus = true;
            //}
            //else if (string.IsNullOrEmpty(this.txttk_no.Text.Trim()) || !this.txttk_no.CheckLostFocus())
            //{
            //    int num = (int)ExMessageBox.Show(699, StartupBase.SasObj, "Chưa vào tài khoản nợ", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    this.txttk_no.IsFocus = true;
            //}
            else
            {
                this.isOk = true;
                this.Close();
            }
        }
        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            this.Close();
        }
        private void txtMa_qs_pt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtMa_qs_pt != null)
            {
                if (this.txtMa_qs_pt.RowResult == null)
                    return;
                this.lbltendvcss.Text = this.txtMa_qs_pt.RowResult[StartupBase.M_LAN.Equals("V") ? "ten_vt" : "ten_vt2"].ToString();
            }
            else
                this.lbltendvcss.Text = "";
            if (e.NewFocus == this.GrdOkCancel.pnlButton.btnOk)
                return;
        }

        public string GetNewSoct(string ma_qs)
        {
            string str1;
            DataTable table = this.BindingSasObj.ExcuteReader(new SqlCommand(Convert.ToInt16(this.BindingSasObj.GetOption("M_AUTO_SOCT").ToString()) == (short)1 ? "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'" : (str1 = "EXEC  [GetNewSoct] '" + ma_qs.Trim() + "'"))).Tables[0];
            if (table.Rows.Count > 0)
            {
                DataRow row = table.Rows[0];
                if (row[1] != null && row[1] != DBNull.Value)
                {
                    string str2 = row[1].ToString();
                    return string.Format(row[0].ToString(), (object)Convert.ToDouble(str2));
                }
            }
            return "";
        }


        private void txtso_ct1_LostFocus(object sender, RoutedEventArgs e)
        {
            
        }

        private void txttk_no_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }

        private void txttk_co_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }
    }
}
