using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;

namespace Inctpxd
{
    public partial class FrmPrintInctpxd : Form
    {
        public DataSet DsPrint = new DataSet();

        public FrmPrintInctpxd()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.GridSearch.ReportGroupName = StartUpTrans.CommandInfo["rep_file"].ToString();
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.GridSearch.DSource = this.DsPrint;
            if (!this.DsPrint.Tables[1].Columns.Contains("stt"))
                this.DsPrint.Tables[1].Columns.Add(new DataColumn("stt", typeof(int)));
            int num = 1;
            foreach (DataRowView dataRowView in this.DsPrint.Tables[1].DefaultView)
            {
                if (!string.IsNullOrEmpty(dataRowView["ma_vt"].ToString().Trim()))
                {
                    dataRowView["stt"] = (object)num;
                    ++num;
                }
            }
            DataTable phIn = StartUpTrans.GetPhIn();
            if (phIn.Rows.Count == 0)
            {
                DataRow row = phIn.NewRow();
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["stt_rec"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                row["so01"] = (object)0;
                row["so02"] = StartUpTrans.DmctInfo["so_lien"] == DBNull.Value ? (object)1 : StartUpTrans.DmctInfo["so_lien"];
                phIn.Rows.Add(row);
            }
            this.DataContext = (object)phIn;
        }

        private void btnin_Click(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.Value != null)
            {
                int num1 = 1;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                int num2 = 1;
                foreach (DataRowView dataRowView in this.DsPrint.Tables[1].DefaultView)
                {
                    if (!string.IsNullOrEmpty(dataRowView["ma_vt"].ToString().Trim()))
                    {
                        dataRowView["stt"] = (object)num2;
                        ++num2;
                    }
                }
                for (; num1 <= int16; ++num1)
                {
                    this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien"] = (object)int16;
                    this.GridSearch.V_In((short)1);
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void btnin_lt_Click(object sender, RoutedEventArgs e)
        {
            if (StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString() == "1" && ExMessageBox.Show(1075, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes && this.txtlien.Value != null)
            {
                int iRow = FrmInctpxd.iRow;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                for (int index1 = 1; index1 < this.DsPrint.Tables[0].Rows.Count; ++index1)
                {
                    for (int index2 = 1; index2 <= int16; ++index2)
                    {
                        this.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables[0].Rows[index1]["stt_rec"].ToString() + "'";
                        this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables[0].Rows[index1]["stt_rec"].ToString() + "'";
                        this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
                        this.DsPrint.Tables["TablePH"].DefaultView[0]["so_lien"] = (object)int16;
                        this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                        int num = 1;
                        foreach (DataRowView dataRowView in this.DsPrint.Tables[1].DefaultView)
                        {
                            dataRowView["stt"] = dataRowView["stt_rec0"] != (object)"999" ? (object)num : (object)DBNull.Value;
                            ++num;
                        }
                        this.GridSearch.V_In((short)1);
                    }
                }
                this.DsPrint.Tables["TablePH"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables["TablePH"].Rows[iRow]["stt_rec"].ToString() + "'";
                this.DsPrint.Tables["TableCT"].DefaultView.RowFilter = "stt_rec= '" + this.DsPrint.Tables["TablePH"].Rows[iRow]["stt_rec"].ToString() + "'";
                this.DsPrint.Tables["TableCT"].DefaultView.Sort = "stt_rec0";
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void btnxem_Click(object sender, RoutedEventArgs e)
        {
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
            this.GridSearch.V_Xem();
        }

        private void btnthoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void txtlien_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.IsFocusWithin || !(this.txtlien.Value.ToString() == ""))
                return;
            this.txtlien.Value = (object)0;
        }

        private void txtctu0_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtctu0.IsFocusWithin)
                return;
            if (this.txtctu0.Value.ToString() == "")
                this.txtctu0.Value = (object)0;
            this.DsPrint.Tables["TablePH"].DefaultView[0]["so_ct_goc"] = this.txtctu0.Value;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.DsPrint.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.DsPrint.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
        }
        public string GetFilePDFName(DataRow dr)
        {
            string text = "";
            text = string.Format("{0}_{1}_{2}", Convert.ToDateTime(dr["ngay_ct"]).ToString("yyyyMMdd"), dr["ma_qs"].ToString().Trim(), dr["so_ct"].ToString().Trim());
            return SysFunc.DeleteSpecialCharacter(StartupBase.SasObj, text);
        }
        public void ExportPDT_LT(int _loai)
        {
            bool flag = false;
            string text = "";
            IntPtr handle = new WindowInteropHelper((Window)(object)this).Handle;
            if (!Checker.IsValidUnit2Print(GridSearch.DSource))
            {
                ExMessageBox.Show(54321, StartupBase.SasObj, "Chứng từ không thuộc đơn vị hiện thời!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            Record activeRecord = ((DataPresenterBase)GridSearch.XGReport).ActiveRecord;
            //DataRowView dataRowView = (activeRecord as DataRecord).DataItem as DataRowView;
            if (StartupBase.SasObj.GetOption("M_IN_HOI_CK").ToString() == "1" && ExMessageBox.Show(4002, StartupBase.SasObj, "Có chắc chắn xuất PDF tất cả các chứng từ đã được lọc?", "SISERPSME 2024", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                using (System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyComputer;
                    System.Windows.Forms.DialogResult dialogResult = folderBrowserDialog.ShowDialog();
                    if (dialogResult == System.Windows.Forms.DialogResult.OK)
                    {
                        if (_loai == 1)
                        {
                            GridSearch.KhoitaoXuatPdf_LT();
                        }
                        text = folderBrowserDialog.SelectedPath;
                        for (int i = 1; i < DsPrint.Tables[0].Rows.Count; i++)
                        {
                            if (DsPrint.Tables[0].Rows[i]["status"].ToString() == "3")
                            {
                                continue;
                            }
                            string sttlst = DsPrint.Tables[0].Rows[i]["stt_rec"].ToString();
                            string filePDFName = GetFilePDFName(DsPrint.Tables[0].Rows[i]);
                            DsPrint.Tables[0].Rows[i]["so_ct_goc"] = (object)this.txtctu0.Text;
                            DsPrint.Tables[0].AcceptChanges();
                            int num = 1;
                            int num2 = 1;
                            int num3 = 1;
                            int num4 = 1;
                            if (this.txtlien.Value != null)
                            {
                                double num5 = Convert.ToDouble(this.txtlien.Value);
                                num4 = Convert.ToInt32(Math.Ceiling(num5 / Convert.ToDouble((activeRecord as DataRecord).Cells["so_lien"].Value)));
                                num3 = ((Convert.ToInt32(this.txtlien.Value) % num4 <= 0) ? (Convert.ToInt32(this.txtlien.Value) / num4) : (Convert.ToInt32(this.txtlien.Value) / num4 + 1));
                                num = 1;
                                num2 = 1;
                                while (num <= num3)
                                {
                                    DsPrint.Tables[0].DefaultView.RowFilter = "stt_rec= '" + DsPrint.Tables[0].Rows[i]["stt_rec"].ToString() + "'";
                                    DsPrint.Tables[1].DefaultView.RowFilter = "stt_rec= '" + DsPrint.Tables[0].Rows[i]["stt_rec"].ToString() + "'";
                                    DsPrint.Tables[1].DefaultView.Sort = "stt_rec0";
                                    num++;
                                    num2 += 2;
                                }
                            }
                            GridSearch.DSource = DsPrint;
                            GridSearch.V_XuatPdf_LT(text, filePDFName, handle, _loai);
                        }
                        flag = true;
                        if (_loai == 1)
                        {
                            string filePDFName1 = string.Format("{0}_{1}_{2}", DateTime.Now.ToString("yyyyMMddHHmmss"), DsPrint.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim(), DsPrint.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim());
                            string _filename = text + "\\" + filePDFName1 + ".pdf";
                            GridSearch.MergeReport_LT(_filename);
                        }
                        StartUpTrans.SetPhIn(base.DataContext as DataTable);
                    }
                }
            }
            if (flag && ExMessageBox.Show(3901, StartupBase.SasObj, "Chương trình đã thực hiện xong, có muốn mở thư mục chứa tệp PDF?", "SISERPSME 2024", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                Process.Start(text);
            }
        }
        private void Btnpdflt_Click(object sender, RoutedEventArgs e)
        {
            int _gopfile = string.IsNullOrEmpty(StartupBase.SasObj.GetOption("M_EXPORT_LT").ToString()) ? 0 : int.Parse(StartupBase.SasObj.GetOption("M_EXPORT_LT").ToString());
            ExportPDT_LT(_gopfile);
        }
    }
}
