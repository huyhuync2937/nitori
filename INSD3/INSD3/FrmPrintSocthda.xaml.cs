using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Data;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace INSD3
{
    public partial class FrmPrintSocthda : Form
    {
        public DataSet DsPrint = new DataSet();

        public FrmPrintSocthda()
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            SysFunc.LoadIcon((Window)this);
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.GridSearch.ReportGroupName = StartUp.commandInfo["rep_file"].ToString().Trim();
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.BtnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.BtnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
            this.GridSearch.DSource = this.DsPrint;
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.Xem();
        }

        private void btnin_Click(object sender, RoutedEventArgs e)
        {
            this.In();
        }

        private void Xem()
        {
            if (!(this.GridSearch.XGReport.ActiveRecord is DataRecord))
                return;
            Decimal num = Math.Ceiling((Decimal)(this.DsPrint.Tables["tbListkho"].Rows.Count + 1) / new Decimal(5));
            for (int index1 = 0; (Decimal)index1 < num; ++index1)
            {
                string str1;
                string index2;
                if (5 * index1 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str1 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1]["ma_kho"].ToString().Trim();
                    index2 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1]["items"].ToString().Trim();
                }
                else if (5 * index1 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str1 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index2 = "T_Nh";
                }
                else
                {
                    str1 = "";
                    index2 = (string)null;
                }
                string str2;
                string index3;
                if (5 * index1 + 1 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str2 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 1]["ma_kho"].ToString().Trim();
                    index3 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 1]["items"].ToString().Trim();
                }
                else if (5 * index1 + 1 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str2 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index3 = "T_Nh";
                }
                else
                {
                    str2 = "";
                    index3 = (string)null;
                }
                string str3;
                string index4;
                if (5 * index1 + 2 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str3 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 2]["ma_kho"].ToString().Trim();
                    index4 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 2]["items"].ToString().Trim();
                }
                else if (5 * index1 + 2 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str3 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index4 = "T_Nh";
                }
                else
                {
                    str3 = "";
                    index4 = (string)null;
                }
                string str4;
                string index5;
                if (5 * index1 + 3 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str4 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 3]["ma_kho"].ToString().Trim();
                    index5 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 3]["items"].ToString().Trim();
                }
                else if (5 * index1 + 3 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str4 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index5 = "T_Nh";
                }
                else
                {
                    str4 = "";
                    index5 = (string)null;
                }
                string str5;
                string index6;
                if (5 * index1 + 4 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str5 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 4]["ma_kho"].ToString().Trim();
                    index6 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 4]["items"].ToString().Trim();
                }
                else if (5 * index1 + 4 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str5 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index6 = "T_Nh";
                }
                else
                {
                    str5 = "";
                    index6 = (string)null;
                }
                this.DsPrint.Tables["tbHeader"].Clear();
                DataRow row1 = this.DsPrint.Tables["tbHeader"].NewRow();
                row1["header1"] = (object)str1;
                row1["header2"] = (object)str2;
                row1["header3"] = (object)str3;
                row1["header4"] = (object)str4;
                row1["header5"] = (object)str5;
                this.DsPrint.Tables["tbHeader"].Rows.Add(row1);
                this.DsPrint.Tables["tbColumn"].Clear();
                DataTable dataTable1 = new DataTable();
                DataTable dataTable2;
                if (index2 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag").Copy();
                else if (index3 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag", index2).Copy();
                else if (index4 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag", index2, index3).Copy();
                else if (index5 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag", index2, index3, index4).Copy();
                else if (index6 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag", index2, index3, index4, index5).Copy();
                else
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "ten_vt2", "dvt", "fsort", "ftag", index2, index3, index4, index5, index6).Copy();
                foreach (DataRow row2 in (InternalDataCollectionBase)dataTable2.Rows)
                {
                    DataRow row3 = this.DsPrint.Tables["tbColumn"].NewRow();
                    row3["stt"] = row2["stt"];
                    row3["ma_vt"] = row2["ma_vt"];
                    row3["ten_vt"] = row2["ten_vt"];
                    row3["ten_vt2"] = row2["ten_vt2"];
                    row3["dvt"] = row2["dvt"];
                    row3["ftag"] = row2["ftag"];
                    row3["fsort"] = row2["fsort"];
                    if (index2 != null)
                        row3["Column1"] = row2[index2];
                    if (index3 != null)
                        row3["Column2"] = row2[index3];
                    if (index4 != null)
                        row3["Column3"] = row2[index4];
                    if (index5 != null)
                        row3["Column4"] = row2[index5];
                    if (index6 != null)
                        row3["Column5"] = row2[index6];
                    this.DsPrint.Tables["tbColumn"].Rows.Add(row3);
                }
                this.DsPrint.Tables["tbColumn"].DefaultView.Sort = StartUp.tbDetail.DefaultView.Sort;
                this.GridSearch.V_Xem(true);
            }
        }

        public Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = new Decimal(0);
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private void In()
        {
            if (!(this.GridSearch.XGReport.ActiveRecord is DataRecord))
                return;
            Decimal num = Math.Ceiling((Decimal)(this.DsPrint.Tables["tbListkho"].Rows.Count + 1) / new Decimal(5));
            for (int index1 = 0; (Decimal)index1 < num; ++index1)
            {
                string str1;
                string index2;
                if (5 * index1 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str1 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1]["ma_kho"].ToString().Trim();
                    index2 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1]["items"].ToString().Trim();
                }
                else if (5 * index1 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str1 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index2 = "T_Nh";
                }
                else
                {
                    str1 = "";
                    index2 = (string)null;
                }
                string str2;
                string index3;
                if (5 * index1 + 1 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str2 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 1]["ma_kho"].ToString().Trim();
                    index3 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 1]["items"].ToString().Trim();
                }
                else if (5 * index1 + 1 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str2 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index3 = "T_Nh";
                }
                else
                {
                    str2 = "";
                    index3 = (string)null;
                }
                string str3;
                string index4;
                if (5 * index1 + 2 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str3 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 2]["ma_kho"].ToString().Trim();
                    index4 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 2]["items"].ToString().Trim();
                }
                else if (5 * index1 + 2 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str3 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index4 = "T_Nh";
                }
                else
                {
                    str3 = "";
                    index4 = (string)null;
                }
                string str4;
                string index5;
                if (5 * index1 + 3 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str4 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 3]["ma_kho"].ToString().Trim();
                    index5 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 3]["items"].ToString().Trim();
                }
                else if (5 * index1 + 3 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str4 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index5 = "T_Nh";
                }
                else
                {
                    str4 = "";
                    index5 = (string)null;
                }
                string str5;
                string index6;
                if (5 * index1 + 4 < this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str5 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 4]["ma_kho"].ToString().Trim();
                    index6 = this.DsPrint.Tables["tbListkho"].Rows[5 * index1 + 4]["items"].ToString().Trim();
                }
                else if (5 * index1 + 4 == this.DsPrint.Tables["tbListkho"].Rows.Count)
                {
                    str5 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["lan"].Equals((object)"Việt") ? "TỔNG VT" : "TOTAL";
                    index6 = "T_Nh";
                }
                else
                {
                    str5 = "";
                    index6 = (string)null;
                }
                this.DsPrint.Tables["tbHeader"].Clear();
                DataRow row1 = this.DsPrint.Tables["tbHeader"].NewRow();
                row1["header1"] = (object)str1;
                row1["header2"] = (object)str2;
                row1["header3"] = (object)str3;
                row1["header4"] = (object)str4;
                row1["header5"] = (object)str5;
                this.DsPrint.Tables["tbHeader"].Rows.Add(row1);
                this.DsPrint.Tables["tbColumn"].Clear();
                DataTable dataTable1 = new DataTable();
                DataTable dataTable2;
                if (index2 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag").Copy();
                else if (index3 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag", index2).Copy();
                else if (index4 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag", index2, index3).Copy();
                else if (index5 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag", index2, index3, index4).Copy();
                else if (index6 == null)
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag", index2, index3, index4, index5).Copy();
                else
                    dataTable2 = this.DsPrint.Tables["tbDetail"].DefaultView.ToTable(true, "stt", "ma_vt", "ten_vt", "dvt", "fsort", "ftag", index2, index3, index4, index5, index6).Copy();
                foreach (DataRow row2 in (InternalDataCollectionBase)dataTable2.Rows)
                {
                    DataRow row3 = this.DsPrint.Tables["tbColumn"].NewRow();
                    row3["stt"] = row2["stt"];
                    row3["ma_vt"] = row2["ma_vt"];
                    row3["ten_vt"] = row2["ten_vt"];
                    row3["dvt"] = row2["dvt"];
                    row3["ftag"] = row2["ftag"];
                    row3["fsort"] = row2["fsort"];
                    if (index2 != null)
                        row3["Column1"] = row2[index2];
                    if (index3 != null)
                        row3["Column2"] = row2[index3];
                    if (index4 != null)
                        row3["Column3"] = row2[index4];
                    if (index5 != null)
                        row3["Column4"] = row2[index5];
                    if (index6 != null)
                        row3["Column5"] = row2[index6];
                    this.DsPrint.Tables["tbColumn"].Rows.Add(row3);
                }
                this.DsPrint.Tables["tbColumn"].DefaultView.Sort = StartUp.tbDetail.DefaultView.Sort;
                this.GridSearch.V_In((short)1);
            }
        }

        private void btnthoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnxem_Click(object sender, RoutedEventArgs e)
        {
            this.Xem();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            IntPtr handle = new WindowInteropHelper((Window)this).Handle;
        }
    }
}
