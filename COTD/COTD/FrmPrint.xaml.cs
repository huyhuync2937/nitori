using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;

namespace COTD
{
    public partial class FrmPrint : Form
    {
        public DataSet DsPrint = new DataSet();

        public FrmPrint()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.GridSearch.ReportGroupName = StartUp.CommandInfo["rep_file"].ToString().Trim();
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
            this.txtDenNgay.Value = (object)DateTime.Today;
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.Xem();
        }

        private void btnin_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.In();
        }

        public double ParseDouble(object obj, double defaultvalue)
        {
            double result = 0.0;
            double.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void Xem()
        {
            if (this.GridSearch.XGReport.ActiveRecord is DataRecord)
            {
                DataSet dataSet = new DataSet();
                DataSet dsPrint = this.DsPrint;
                string str = "";
                if (this.txtTuNgay.IsValueValid && !string.IsNullOrEmpty(this.txtTuNgay.Text))
                    str = "ngay1 >= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtTuNgay.Value) + "'";
                if (this.txtDenNgay.IsValueValid && !string.IsNullOrEmpty(this.txtDenNgay.Text))
                    str = !(str != "") ? str + "ngay1 <= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtDenNgay.Value) + "'" : str + " and ngay1 <= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtDenNgay.Value) + "'";
                dsPrint.Tables[0].DefaultView.RowFilter = str;
                this.GridSearch.DSource = dsPrint;
                this.GridSearch.V_Xem(true);
            }
            this.Close();
        }

        private void In()
        {
            if (this.GridSearch.XGReport.ActiveRecord is DataRecord)
            {
                DataSet dataSet = new DataSet();
                DataSet dsPrint = this.DsPrint;
                string str = "";
                if (this.txtTuNgay.IsValueValid && !string.IsNullOrEmpty(this.txtTuNgay.Text))
                    str = "date0 >= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtTuNgay.Value) + "'";
                if (this.txtDenNgay.IsValueValid && !string.IsNullOrEmpty(this.txtDenNgay.Text))
                    str = !(str != "") ? str + "date0 <= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtDenNgay.Value) + "'" : str + " and date0 <= '" + string.Format("{0:dd-MM-yyyy}", (object)(DateTime)this.txtDenNgay.Value) + "'";
                dsPrint.Tables[0].DefaultView.RowFilter = str;
                this.GridSearch.DSource = dsPrint;
                this.GridSearch.V_In((short)1);
            }
            this.Close();
        }

        private void btnthoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnxem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Xem();
        }

    }
}
