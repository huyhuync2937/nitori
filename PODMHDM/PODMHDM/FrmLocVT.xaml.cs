using SasControls;
using SasDataLib;
using SasDefine;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace PODMHDM
{
    public partial class FrmLocVT : Form
    {
        public FrmLocVT()
        {
            this.InitializeComponent();
        }

        private void FrmLocVT_Loaded(object sender, RoutedEventArgs e)
        {
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtTungay.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            this.txtDenngay.Value = DateTime.Now.Date;
            this.txtStatus.Text = "1";
            this.txtTungay.Focus();
        }

        private void _confirmGridview_OnOk(object sender, RoutedEventArgs e)
        {
            if (Checkvalid())
                return;
            DateTime fromDate = Convert.ToDateTime(this.txtTungay.Value);
            DateTime toDate = Convert.ToDateTime(this.txtDenngay.Value);
            string maVT = this.txtMaVT.Text.Trim();
            string maNcc = this.txtMaNcc.Text.Trim();
            string maPic = this.txtMaPic.Text.Trim();
            string status = this.txtStatus.Text.Trim();
            string prog = "";
            if (!string.IsNullOrEmpty(status))
            {
                if (status.Equals("1"))
                {
                    prog = "COSLDHT";
                }
                else
                {
                    prog = "COSLDHN";
                }
            }

            SqlCommand cmd = new SqlCommand(prog);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@hdTuNg", fromDate.ToString("yyyyMMdd"));
            cmd.Parameters.AddWithValue("@hdDenNg", toDate.ToString("yyyyMMdd"));
            cmd.Parameters.AddWithValue("@maVt", string.IsNullOrEmpty(maVT) ? (object)DBNull.Value : maVT);
            cmd.Parameters.AddWithValue("@ma_kh", string.IsNullOrEmpty(maNcc) ? (object)DBNull.Value : maNcc);
            cmd.Parameters.AddWithValue("@pic", string.IsNullOrEmpty(maPic) ? (object)DBNull.Value : maPic);
            cmd.Parameters.AddWithValue("@mode", 2);
            DataTable dataTable = StartupBase.SasObj.ExcuteReader(cmd).Tables[0].Copy();
            COTKTH2Dvcs cotktH2Dvcs = new COTKTH2Dvcs();
            //dataTable.DefaultView.RowFilter = !string.IsNullOrEmpty(this.txtMa_dvcs.Text.Trim()) ? "ma_dvcs LIKE '" + this.txtMa_dvcs.Text + "'" : "1=1";
            cotktH2Dvcs.GrdCt.DataSource = dataTable.DefaultView;
            cotktH2Dvcs.ShowDialog();
            //this.Close();
            //if (StartupBase.M_LAN != "V")
            //    cotktH2Dvcs.Title = "Supplier list";
            //bool? nullable = cotktH2Dvcs.ShowDialog();
            //if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
            //    return;
            DataRow[] dataRowArray = dataTable.Select("tag = 1");
            StartUp.dataRowArray = dataRowArray;
            //this.Close();

            //FrmLoaddonhang Frmlocdon = new FrmLoaddonhang((DateTime)this.txtTungay.Value, (DateTime)this.txtDenngay.Value, filter);
            //Frmlocdon.ShowDialog();
            this.Close();

        }

        bool Checkvalid()
        {
            bool flag = false;
            if (Convert.ToDateTime(txtTungay.Value) > Convert.ToDateTime(txtDenngay.Value))
            {
                flag = true;
                MessageBox.Show("Từ ngày phải <= đến ngày.", "Thông báo");
                this.txtTungay.Focus();
            }
            else if (txtStatus.Text.Trim() != "1" && txtStatus.Text.Trim() != "2")
            {
                flag = true;
                MessageBox.Show("Trạng thái chỉ được nhập 0 hoặc 1.", "Thông báo");
                this.txtStatus.Focus();
            }
            return flag;
        }

        private void txtStatus_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = e.Text != "1" && e.Text != "2";
        }

        private void _confirmGridview_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void txtMaVT_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }

        private void txtDenngay_LostFocus(object sender, RoutedEventArgs e)
        {

        }

        private void txtTungay_LostFocus(object sender, RoutedEventArgs e)
        {

        }
    }
}
