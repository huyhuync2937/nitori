using SasControls;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

namespace Incd1
{
    public partial class Incd1F10 : FormFilter
    {
        private bool isOK = false;
        public DataTable dtOption;

        public DataTable DataOption
        {
            get
            {
                if (this.dtOption != null)
                    return this.dtOption;
                this.dtOption = new DataTable("F10");
                this.dtOption.Columns.Add("group1", typeof(string));
                this.dtOption.Columns.Add("group2", typeof(string));
                this.dtOption.Columns.Add("group3", typeof(string));
                this.dtOption.Columns.Add("sortby", typeof(string));
                this.dtOption.Columns.Add("ps_no1", typeof(double));
                this.dtOption.Columns.Add("ps_no2", typeof(double));
                this.dtOption.Columns.Add("ps_co1", typeof(double));
                this.dtOption.Columns.Add("ps_co2", typeof(double));
                this.dtOption.Columns.Add("no_ck1", typeof(double));
                this.dtOption.Columns.Add("no_ck2", typeof(double));
                this.dtOption.Columns.Add("co_ck1", typeof(double));
                this.dtOption.Columns.Add("co_ck2", typeof(double));
                this.dtOption.Rows.Add((object)"0", (object)"0", (object)"0", (object)"1", (object)0, (object)0, (object)0, (object)0, (object)0, (object)0, (object)0, (object)0);
                return this.dtOption;
            }
            set
            {
                this.dtOption = value;
            }
        }

        public Incd1F10(DataTable GroupSelectedTable)
        {
            if (GroupSelectedTable != null)
                this.dtOption = GroupSelectedTable.Copy();
            this.InitializeComponent();
            this.DataContext = (object)this.DataOption;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            this.isOK = true;
            this.Close();
        }

        public bool ShowDialog()
        {
            base.ShowDialog();
            return this.isOK;
        }

        private void ConfirmGridView_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtGrp1.Focus();
        }

        private void FormFilter_Loaded(object sender, RoutedEventArgs e)
        {
        }

    }
}
