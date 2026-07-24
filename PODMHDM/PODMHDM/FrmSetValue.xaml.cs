using SasControls;
using SasFormBrowes;
using System;
using System.Data;
using System.Windows;
using System.Windows.Documents;

namespace PODMHDM
{
    public partial class FrmSetValue : Form
    {
        public string title_ = "Cập nhật giá trị";
        public bool isOk = false;
        public decimal packing = 0;
        public decimal he_so = 0;

        public FrmSetValue(DataTable table)
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;

            // Kiểm tra table và số dòng trước khi truy cập
            if (table != null && table.Rows.Count > 0)
            {
                DataRow row = table.Rows[0];

                // Xử lý dvt an toàn: kiểm tra cột tồn tại + không phải DBNull
                if (table.Columns.Contains("dvt") && row["dvt"] != DBNull.Value)
                {
                    string dvtValue = row["dvt"].ToString();
                    if (!string.IsNullOrEmpty(dvtValue))
                    {
                        this.txtDvt1.Text = dvtValue;
                    }
                }

                // Xử lý packing an toàn: kiểm tra cột tồn tại + không phải DBNull
                if (table.Columns.Contains("packing") && row["packing"] != DBNull.Value)
                {
                    decimal.TryParse(row["packing"].ToString(), out packing);
                }
                else
                {
                    packing = 0;
                }

                DataRowView drv = table.DefaultView[0];
                this.DataContext = drv;
            }
            else
            {
                packing = 0;
            }

            txtNgay_bh.Focus();
            SysFunc.LoadIcon((Window)this);
        }


        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }
        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            //this.txtNewValue.SearchInit();
            //if(this.txtNewValue.RowResult == null)
            //{
            //    int num = (int)ExMessageBox.Show(9050, StartupBase.SasObj, "Chưa chọn giá trị hoặc giá trị không có trong danh mục.!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            //    return;
            //}    

            this.DialogResult = new bool?(true);
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.DialogResult = new bool?(false);
            this.Close();
        }
        private void txtso_luong_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtpacking.Text) ||
                string.IsNullOrEmpty(this.txthe_so.Text) ||
                Convert.ToInt32(this.txtpacking.Value ?? 0) == 0 ||
                Convert.ToInt32(this.txthe_so.Value ?? 0) == 0)
                return;
            this.txtso_luong.Value = Convert.ToInt32(this.txtpacking.Value) * Convert.ToInt32(this.txthe_so.Value);
        }
        private void txtpacking_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txtpacking.Text.Trim() == ""))
                return;
            this.txtpacking.Text = packing.ToString();

        }
        private void txthe_so_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(this.txthe_so.Text.Trim() == ""))
                return;
            this.txthe_so.Text = he_so.ToString();
        }
    }
}


