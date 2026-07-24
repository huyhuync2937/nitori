using Microsoft.Win32;
using SasControls;
using SasDataLib;
using SasFormBrowes;
using SasUtilities;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;

namespace Poctpna
{
    /// <summary>
    /// Interaction logic for FrmImportInvoice.xaml
    /// </summary>
    public partial class FrmImport : Form
    {
        public bool isOk = false;
        public DataSet dsInvoice;

        public FrmImport()
        {
            InitializeComponent();
        }
        private void FrmImport_Loaded(object sender, RoutedEventArgs e)
        {
            this.Title = string.Format(StartUp.M_LAN == "V" ? "Nhap khau hoa don dien tu" : "Import of electronic invoices");
            this.txtFile.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\invoice.xml";
            DataTable LocalTable1 = StartUpTrans.DsTrans.Tables[0].Clone();
            DataTable LocalTable2 = StartUpTrans.DsTrans.Tables[1].Clone();
            DataTable LocalTable3 = StartUpTrans.DsTrans.Tables[2].Clone();
            dsInvoice = new DataSet();
            dsInvoice.Tables.Add(LocalTable1);
            dsInvoice.Tables.Add(LocalTable2);
            dsInvoice.Tables.Add(LocalTable3);
        }
        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            this.Close();
        }
        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtFile.Text))
            {
                DataSet ds = Converts.XmlToDataSet(this.txtFile.Text.Trim());
                string HDDT = string.Empty;
                bool flag = Converts.GetNccHDDT((DataSet)ds, ref HDDT);
                if (flag)
                {
                    int num = (int)ExMessageBox.Show(10, StartupBase.SasObj, "Hóa đơn điện tử không đúng. Vui lòng kiểm tra lại!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return;
                }
                string result = Converts.GetDsInvoice(this.BindingSasObj, StartUp.Ma_ct, ds, ref dsInvoice, HDDT);
                this.isOk = true;
                this.Close();
            }
            else
            {
                int num = (int)ExMessageBox.Show(30, StartupBase.SasObj, "File không tồn tại, hãy thử lại lần sau!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.isOk = false;
                this.Close();
            }
        }
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "xml";
            openFileDialog.FilterIndex = 0;
            openFileDialog.Filter = "XML Files (*.xml)|*.xml";
            openFileDialog.ShowDialog();
            if (string.IsNullOrEmpty(openFileDialog.FileName))
                return;
            this.txtFile.Text = openFileDialog.FileName;
        }
    }
}
