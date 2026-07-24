using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System.Windows;

namespace ARCTHD1
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
            string strnoidung = "Kính gửi quý khách hàng \n Cảm ơn quý khách hàng đã mua hàng tại " + StartUp.dtRegInfo.Rows[0]["content"].ToString();
            strnoidung += "\n Quý khách tải xuống file PDF hóa đơn bán hàng được đính kèm để kiểm tra thông tin.";
            strnoidung += "\n Mọi thông tin sai sót xin vui lòng phải hồi lại Công ty để nhận được hỗ trợ. \n Trân trọng cảm ơn.!";
            this.txtTieude.Text = "Hóa đơn bán hàng";
            this.txtNoidung.Text = strnoidung;


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
