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

            SqlCommand sql_link = new SqlCommand("SELECT top 1  ten_td FROM dmtd5 ");
            string link = "";
            DataTable table_link = StartupBase.SasObj.ExcuteReader(sql_link).Tables[0];
            foreach (DataRow row in table_link.Rows)
            {
                link = row["ten_td"].ToString().Trim();
            }

            string strnoidung = "Kính gửi Quý Nhà cung cấp,\n"
                + "Trước hết, Nitori Furniture Việt Nam (NFV) xin chân thành cảm ơn Quý Nhà cung cấp đã đồng hành\n"
                + "và hợp tác trong việc cung cấp nguyên vật liệu cho công ty chúng tôi.\n"
                + "NFV kính đề nghị Quý Nhà cung cấp gửi báo giá cho một số nguyên vật liệu theo nhu cầu mua hàng\n"
                + "của chúng tôi.\n"
                + "Quý Nhà cung cấp vui lòng truy cập đường dẫn dưới đây để xem thông tin chi tiết và thực hiện báo\n"
                + "giá:\n"
                + link + "\n"
                + "Chúng tôi rất mong tiếp tục nhận được sự hỗ trợ và hợp tác từ Quý Nhà cung cấp.\n"
                + "Xin chân thành cảm ơn và kính chúc Quý Công ty ngày càng phát triển, thịnh vượng.\n"
                + "Trân trọng,\n"
                + "Nitori Furniture Việt Nam (NFV)\n"
                + "****************\n"
                + "Subject: Request for Quotation (RFQ) – Nitori\n"
                + "Dear Valued Supplier,\n"
                + "First of all, Nitori Furniture Vietnam (NFV) would like to express our sincere appreciation for your\n"
                + "continued support and cooperation in supplying raw materials to our company.\n"
                + "We would like to invite your company to submit quotations for several raw materials required by NFV.\n"
                + "Please access the link below to review the detailed information and submit your quotation:\n"
                + link + "\n"
                + "We look forward to your continued support and cooperation.\n"
                + "Thank you very much, and we wish your company continued success and prosperity.\n"
                + "Best regards,\n"
                + "Nitori Furniture Vietnam (NFV)";

            string strnoidung1 = $@"
        <html>
        <body>
            <p>Kính gửi Quý Nhà cung cấp,</p>

            <p>
                Trước hết, Nitori Furniture Việt Nam (NFV) xin chân thành cảm ơn Quý Nhà cung cấp đã đồng hành
                và hợp tác trong việc cung cấp nguyên vật liệu cho công ty chúng tôi.
            </p>

            <p>
                NFV kính đề nghị Quý Nhà cung cấp gửi báo giá cho một số nguyên vật liệu theo nhu cầu mua hàng của chúng tôi.
            </p>

            <p>
                Quý Nhà cung cấp vui lòng truy cập đường dẫn dưới đây để xem thông tin chi tiết và thực hiện báo giá:
            </p>

            <p>
                <a href='{link}' target='_blank'>
                    {link}
                </a>
            </p>

            <p>
                Chúng tôi rất mong tiếp tục nhận được sự hỗ trợ và hợp tác từ Quý Nhà cung cấp.
            </p>

            <p>Xin chân thành cảm ơn và kính chúc Quý Công ty ngày càng phát triển, thịnh vượng.</p>

            <p>
                Trân trọng,<br/>
                Nitori Furniture Việt Nam (NFV)
            </p>

            <hr/>

            <p>Dear Valued Supplier,</p>

            <p>
                First of all, Nitori Furniture Vietnam (NFV) would like to express our sincere appreciation for your
                continued support and cooperation in supplying raw materials to our company.
            </p>

            <p>
                We would like to invite your company to submit quotations for several raw materials required by NFV.
            </p>

            <p>
                Please access the link below to review the detailed information and submit your quotation:
            </p>

            <p>
                <a href='{link}' target='_blank'>
                    {link}
                </a>
            </p>

            <p>
                We look forward to your continued support and cooperation.
            </p>

            <p>Thank you very much, and we wish your company continued success and prosperity.</p>

            <p>
                Best regards,<br/>
                Nitori Furniture Vietnam (NFV)
            </p>
        </body>
        </html>";

            this.txtTieude.Text = "Đề nghị báo giá - Nitori";
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
