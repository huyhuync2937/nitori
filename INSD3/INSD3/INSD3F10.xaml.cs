using SasFormReport;
using System.Data;
using System.Windows;
using System.Windows.Input;

namespace INSD3
{
    public partial class INSD3F10 : FormFilter
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

        public INSD3F10(DataTable GroupSelectedTable)
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

        private void txtGrp1_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!(this.txtGrp1.Value.ToString() == "0"))
                return;
            this.txtGrp2.Value = (object)"0";
            this.txtGrp3.Value = (object)"0";
        }

        private void txtGrp2_PreviewLostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (!(this.txtGrp2.Value.ToString() == "0"))
                return;
            this.txtGrp3.Value = (object)"0";
        }

    }
}
