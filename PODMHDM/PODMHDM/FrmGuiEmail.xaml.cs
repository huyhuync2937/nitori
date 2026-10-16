using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace PODMHDM
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

            string strnoidung = "Kính gửi Quý Nhà cung cấp,\n";
            strnoidung += "Nitori Furniture Việt Nam (NFV) xin chân thành cảm ơn Quý Nhà cung cấp đã luôn đồng hành và hợp tác cùng chúng tôi trong việc cung cấp nguyên vật liệu.\n";
            strnoidung += "Chúng tôi xin gửi đến Quý Công ty Đơn đặt hàng (Purchase Order - PO) cho các nguyên vật liệu theo nhu cầu mua hàng của NFV.\n";
            strnoidung += "Quý Công ty vui lòng truy cập đường dẫn dưới đây để xem chi tiết Đơn đặt hàng và Lịch yêu cầu giao hàng:\n";
            strnoidung += link + "\n";
            strnoidung += "Kính đề nghị Quý Công ty xác nhận đơn hàng và phối hợp thực hiện giao hàng theo lịch trình đã được thống nhất.\n";
            strnoidung += "Chúng tôi rất mong tiếp tục nhận được sự hỗ trợ và hợp tác từ Quý Công ty trong thời gian tới.\n";
            strnoidung += "Xin chân thành cảm ơn và kính chúc Quý Công ty ngày càng phát triển, thịnh vượng.\n";
            strnoidung += "Trân trọng,\n";
            strnoidung += "Nitori Furniture Việt Nam (NFV)\n";
            strnoidung += "************************\n";
            strnoidung += "Dear Valued Supplier,\n";
            strnoidung += "Nitori Furniture Vietnam (NFV) would like to express our sincere appreciation for your continued support and cooperation in supplying raw materials to our company.\n";
            strnoidung += "We are pleased to send you our Purchase Order (PO) for the required raw materials.\n";
            strnoidung += "Please access the link below to review the detailed Purchase Order and Delivery Schedule:\n";
            strnoidung += link + "\n";
            strnoidung += "We kindly request your confirmation of the order and your cooperation in arranging delivery according to the agreed schedule.\n";
            strnoidung += "We look forward to your continued support and partnership.\n";
            strnoidung += "Thank you for your cooperation.\n";
            strnoidung += "Best regards,\n";
            strnoidung += "Nitori Furniture Vietnam (NFV)";

            this.txtTieude.Text = "Đơn đặt hàng (PO) - Nitori";
            this.txtNoidung.Text = strnoidung;

            string strnoidung1 = $@"
        <html>
        <body>
            <p>Kính gửi Quý Nhà cung cấp,</p>

            <p>
                Nitori Furniture Việt Nam (NFV) xin chân thành cảm ơn Quý Nhà cung cấp đã luôn đồng hành và hợp tác cùng chúng tôi trong việc cung cấp nguyên vật liệu.
            </p>

            <p>
                Chúng tôi xin gửi đến Quý Công ty Đơn đặt hàng (Purchase Order - PO) cho các nguyên vật liệu theo nhu cầu mua hàng của NFV.
            </p>

            <p>
                Quý Công ty vui lòng truy cập đường dẫn dưới đây để xem chi tiết Đơn đặt hàng và Lịch yêu cầu giao hàng:
            </p>

            <p>
                <a href='{link}' target='_blank'>
                    [Link đơn hàng]
                </a>
            </p>

            <p>
                Kính đề nghị Quý Công ty xác nhận đơn hàng và phối hợp thực hiện giao hàng theo lịch trình đã được thống nhất.
            </p>

            <p>
                Chúng tôi rất mong tiếp tục nhận được sự hỗ trợ và hợp tác từ Quý Công ty trong thời gian tới.
            </p>

            <p>
                Xin chân thành cảm ơn và kính chúc Quý Công ty ngày càng phát triển, thịnh vượng.
            </p>

            <p>
                Trân trọng,<br/>
                Nitori Furniture Việt Nam (NFV)
            </p>

            <hr/>

            <p>Dear Valued Supplier,</p>

            <p>
                Nitori Furniture Vietnam (NFV) would like to express our sincere appreciation for your continued support and cooperation in supplying raw materials to our company.
            </p>

            <p>
                We are pleased to send you our Purchase Order (PO) for the required raw materials.
            </p>

            <p>
                Please access the link below to review the detailed Purchase Order and Delivery Schedule:
            </p>

            <p>
                <a href='{link}' target='_blank'>
                    [PO Link]
                </a>
            </p>

            <p>
                We kindly request your confirmation of the order and your cooperation in arranging delivery according to the agreed schedule.
            </p>

            <p>
                We look forward to your continued support and partnership.
            </p>

            <p>Thank you for your cooperation.</p>

            <p>
                Best regards,<br/>
                Nitori Furniture Vietnam (NFV)
            </p>
        </body>
        </html>";

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
