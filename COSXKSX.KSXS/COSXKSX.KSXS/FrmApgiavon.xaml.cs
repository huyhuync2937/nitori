using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace COSXKSX.KSXS
{
    public partial class FrmApgiavon : Form
    {
        private DataSet dsSource = new DataSet();

        public FrmApgiavon()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.ShowInTaskbar = false;
            this.txtNam.Value = DateTime.Now.Year;
            this.txtThang.Value = DateTime.Now.Month;
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Title = SysFunc.Cat_Dau(this.Title);
                this.txtNam.Focus();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void BtnIn_Click(object sender, RoutedEventArgs e)
        {
        }

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
         
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void txtGia_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {

        }

        private void txtKho_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (this.txtKho.RowResult == null)
                return;
            this.txtTen_kho.Text = StartUpTrans.M_LAN.Equals("V") ? this.txtKho.RowResult["ten_kho"].ToString() : this.txtKho.RowResult["ten_kho2"].ToString();
        }

        private void BtnNhan_Click(object sender, RoutedEventArgs e)
        {
            if(StartUpTrans.DsTrans.Tables[2].Rows.Count <0)
            {
                int num = (int)ExMessageBox.Show(4221, StartupBase.SasObj, "Chưa có NVL!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }    
            string mthang = "";
            string colthang = "";
            string colthang_nt = "";
            string filter = " 1=1 ";
            if (!string.IsNullOrEmpty(this.txtTen_kho.Text))
                filter += " and ma_kho like '" +this.txtKho.Text.Trim()+"%'";
            mthang = this.txtThang.Value.ToString();
            if (mthang.Length < 2)
            {
                colthang_nt = "gia_nt0" + mthang;
                colthang = "gia0" + mthang;             
            }
            else
            {
                colthang_nt = "gia_nt" + mthang;
                colthang = "gia" + mthang;
            }  
            string sql = "Select ma_vt, " + colthang + ", "+ colthang_nt + " from dmgiavv where nam = @nam and " + filter;
            SqlCommand cmd = new SqlCommand(sql);
            cmd.Parameters.Add("@nam",SqlDbType.Char).Value = (object)this.txtNam.Value.ToString();
         
            DataTable TblGia = StartupBase.SasObj.ExcuteReader(cmd).Tables[0];
            if(TblGia.Rows.Count!=0)
            {
                foreach(DataRowView view in StartUpTrans.DsTrans.Tables[2].DefaultView)
                {
                    DataRow[] rowGia = TblGia.Select(" ma_vt = '" +view["ma_vt"].ToString().Trim()+ "'");
                    if(rowGia.Length>0)
                    {
                        view["gia"] = rowGia[0][colthang];
                        view["gia_nt"] = rowGia[0][colthang_nt];
                        decimal num1 = 0;
                        decimal num2 = 0;
                        decimal num3 = 0;                      
                        decimal.TryParse(rowGia[0][colthang].ToString(),out num1);
                        decimal.TryParse(rowGia[0][colthang_nt].ToString(), out num2);
                        decimal.TryParse(view["so_luong"].ToString(), out num3);
                        view["tien"] = num1 * num3;
                        view["tien_nt"] = num2 * num3;                       
                    }
                }
            }
            ExMessageBox.Show(1691, StartupBase.SasObj, "Đã hoàn thành", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.Close();
        }
    }
}
