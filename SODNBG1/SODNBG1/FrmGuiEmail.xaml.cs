using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace SODNBG1
{
    public partial class FrmGuiEmail : Form
    {
        public bool isError = true;

        public FrmGuiEmail()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartUpTrans.M_LAN;
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
        }

     private void Form_Loaded(object sender, RoutedEventArgs e)
        {

            string companyName = StartUp.dtRegInfo.Rows[0]["content"].ToString();

            SqlCommand sql_link = new SqlCommand("SELECT top 1  ten_td FROM dmtd5 ");
            string link = "";
            DataTable table_link = StartupBase.SasObj.ExcuteReader(sql_link).Tables[0];
            foreach (DataRow row in table_link.Rows)
            {
                link = row["ten_td"].ToString().Trim();
            }

            string strnoidung = "Kính gửi quý khách hàng \n Cảm ơn quý khách hàng đã mua hàng tại " + StartUp.dtRegInfo.Rows[0]["content"].ToString();
            strnoidung += "\n Quý khách ấn vào link " + link + " được đính kèm để kiểm tra thông tin.";
            strnoidung += "\n Mọi thông tin sai sót xin vui lòng phải hồi lại Công ty để nhận được hỗ trợ. \n Trân trọng cảm ơn.!";
            this.txtTieude.Text = "Hóa đơn bán hàng";
            this.txtNoidung.Text = strnoidung;

            string strnoidung1 = $@"
        <html>
        <body>
            <p>Kính gửi quý khách hàng,</p>

            <p>
                Cảm ơn quý khách hàng đã mua hàng tại <b>{companyName}</b>
            </p>

            <p>
                Quý khách vui lòng nhấn vào link bên dưới để kiểm tra thông tin:
            </p>

            <p>
                <a href='{link}' target='_blank'>
                    Xem thông tin hóa đơn
                </a>
            </p>

            <p>
                Mọi thông tin sai sót xin vui lòng phản hồi lại Công ty để nhận được hỗ trợ.
            </p>

            <p>Trân trọng cảm ơn.!</p>
        </body>
        </html>";

            this.txtTieude.Text = "Hóa đơn bán hàng";
            this.txtNoidung.Text = strnoidung;
            this.txtNoidung1.Text = strnoidung1;

            this.txtTieude.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.isError = false;
            this.Close();
        }


        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isError = true;
            this.Close();
        }
    }
}
