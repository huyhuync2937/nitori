using SasControls;
using System;
using System.Data;
using System.Windows;
using System.Windows.Media.Imaging;
using iTextSharp.text.pdf;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing;
using iTextSharp.text;
using System.Text;
using System.Runtime.CompilerServices;

namespace Invt
{
    /// <summary>
    /// Interaction logic for FrmMavach.xaml
    /// </summary>
    public partial class FrmMavach : FormList
    {
        private DataTable newDataTable = null;
        private DataTable dataTableSP = null;
        public FrmMavach()
        {
            InitializeComponent();
        }
        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            NumericTextBox numericTextBox = sender as NumericTextBox;
            if (!numericTextBox.IsInEditMode)
                return;
            numericTextBox.SelectAll();
        }

        private void txtLostFocus(object sender, RoutedEventArgs e)
        {
            NumericTextBox numericTextBox = sender as NumericTextBox;
            if (numericTextBox.Value != null && !(numericTextBox.Value is DBNull))
                return;
            numericTextBox.Value = (object)0.0;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.DialogResult = new bool?(true);
        }

        private void Mavach_ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            DataRow row = this.dataTableSP.Rows[0];
            for (int i = 0; i < (int)row["sl_tem"]-1; i++)
            {
                
                DataRow row1 = this.dataTableSP.NewRow();
                row1["ma_vt"] = row["ma_vt"].ToString();
                row1["ten_vt"] = row["ten_vt"].ToString();
                row1["ten_vt2"] = row["ten_vt2"].ToString();
                row1["dvt"] = row["dvt"].ToString();
                row1["dvt1"] = row["dvt1"].ToString();
                row1["part_no"] = row["part_no"].ToString();
                row1["sl_sp"] = row["sl_sp"].ToString();
                row1["gia_sp"] = row["gia_sp"].ToString();
                row1["sl_tem"] = row["sl_tem"].ToString();
                this.dataTableSP.Rows.Add(row1);
            }    
            DataSet datasetreport = new DataSet();
            datasetreport.Tables.Add(this.dataTableSP);
            this.Close();
            SasFormReport.ReportManager reportManager = new SasFormReport.ReportManager(StartupBase.SasObj, "INVTMAVACH");
            reportManager.Preview(datasetreport);
        }
        private void Mavach_ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void FrmMavach_Loaded(object sender, RoutedEventArgs e)
        {
            this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
            this.dataTableSP = new DataTable();
            this.dataTableSP.Columns.Add("ma_vt", typeof(string));
            this.dataTableSP.Columns.Add("ten_vt", typeof(string));
            this.dataTableSP.Columns.Add("ten_vt2", typeof(string));
            this.dataTableSP.Columns.Add("dvt", typeof(string));
            this.dataTableSP.Columns.Add("dvt1", typeof(string));
            this.dataTableSP.Columns.Add("part_no", typeof(string));
            this.dataTableSP.Columns.Add("sl_sp", typeof(decimal));
            this.dataTableSP.Columns.Add("gia_sp", typeof(decimal));
            this.dataTableSP.Columns.Add("sl_tem", typeof(Int32));
            foreach (DataRow row in this.newDataTable.Rows)
            {
                DataRow row1 = this.dataTableSP.NewRow();
                row1["ma_vt"] = row["ma_vt"].ToString().Trim();
                row1["ten_vt"] = row["ten_vt"].ToString().Trim();
                row1["ten_vt2"] = row["ten_vt2"].ToString().Trim();
                row1["dvt"] = row["dvt"].ToString().Trim();
                row1["dvt1"] = row["dvt1"].ToString().Trim();
                row1["part_no"] = row["part_no"].ToString().Trim();
                row1["sl_sp"] = 1.0;
                row1["gia_sp"] = 1000.0;
                row1["sl_tem"] = 1;
                this.dataTableSP.Rows.Add(row1);
            }    
            this.gridMain.DataContext = this.dataTableSP;
            Barcode128 barcode128 = new Barcode128();
            barcode128.ChecksumText = true;
            barcode128.TextAlignment = Element.ALIGN_CENTER;
            barcode128.CodeType = Barcode128.CODE128;
            barcode128.Code = this.dataTableSP.Rows[0]["part_no"].ToString().Trim();
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(barcode128.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White));
            imageViewerBmp.Source =  (BitmapSource)(BitmapConversion.ToWpfBitmap(bm));
        }
        private void txtSl_tem_GotFocus(object sender, RoutedEventArgs e)
        {
            this.txtSl_tem.SelectAll();
        }

    }
}
