using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;

namespace Poctpxf
{
    public partial class FrmIn : Form
    {
        private DataSet dsSource = new DataSet();
        private bool IsND51;

        public FrmIn(bool isND51)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.IsND51 = isND51;
            if (isND51)
            {
                this.GridSearch.ReportGroupName = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_file"].ToString();
                this.GridSearch.DefaultFilter = "report not like 'SOCTHDA_GTGT_XKho%'";
            }
            else
                this.GridSearch.ReportGroupName = StartUpTrans.CommandInfo["rep_file"].ToString();
            this.dsSource = StartUpTrans.DsTrans.Copy();
            StartUp.GetDmnt(this.dsSource);
            this.dsSource.Tables[1].Columns.Add(new DataColumn("stt", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("tien2", typeof(Decimal))
            {
                DefaultValue = (object)DBNull.Value
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("gia2", typeof(Decimal))
            {
                DefaultValue = (object)DBNull.Value
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("tien_nt2", typeof(Decimal))
            {
                DefaultValue = (object)DBNull.Value
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("gia_nt2", typeof(Decimal))
            {
                DefaultValue = (object)DBNull.Value
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
            {
                DefaultValue = (object)0
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("ban_sao", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
            this.dsSource.Tables[0].TableName = "TablePH";
            this.dsSource.Tables[1].TableName = "TableCT";
            this.dsSource.Tables.Add(this.CreateTableInfo().Copy());
            this.GridSearch.DSource = this.dsSource;
            this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        private DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "TableInfo";
            dataTable.Columns.Add(new DataColumn("M_Ten_CTY", typeof(string))
            {
                DefaultValue = (object)StartupBase.SasObj.GetSysvar("M_Ten_CTY").ToString()
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        private void InsertSubRow()
        {
            this.dsSource.Tables[0].DefaultView[0]["t_tien_sau_ck_nt"] = this.dsSource.Tables[0].DefaultView[0]["t_tien_nt"];
            this.dsSource.Tables[0].DefaultView[0]["t_tien_sau_ck"] = this.dsSource.Tables[0].DefaultView[0]["t_tien"];
            this.dsSource.Tables[0].DefaultView[0]["so_seri"] = (object)this.dsSource.Tables[0].DefaultView[0]["so_seri"].ToString().Trim();
            int num = 1;
            foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
            {
                if (!dataRowView["ma_vt"].ToString().Trim().Equals(string.Empty))
                {
                    dataRowView["stt"] = (object)num;
                    dataRowView["tien2"] = dataRowView["tien"];
                    dataRowView["tien_nt2"] = dataRowView["tien_nt"];
                    dataRowView["gia2"] = dataRowView["gia"];
                    dataRowView["gia_nt2"] = dataRowView["gia_nt"];
                    ++num;
                }
            }
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.Title = SysFunc.Cat_Dau(this.Title);
            DataTable phIn = StartUpTrans.GetPhIn();
            if (phIn.Rows.Count == 0)
            {
                DataRow row = phIn.NewRow();
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["stt_rec"] = (object)this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                row["so01"] = (object)0;
                row["so02"] = StartUpTrans.DmctInfo["so_lien"] == DBNull.Value ? (object)1 : StartUpTrans.DmctInfo["so_lien"];
                phIn.Rows.Add(row);
            }
            if (phIn.Rows.Count == 1)
            {
                DataRow row = phIn.Rows[0];
                if (this.IsND51)
                    row["so02"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
            }
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GridSearch.XGReport.RecordActivated += new EventHandler<RecordActivatedEventArgs>(this.XGReport_RecordActivated)));
            this.DataContext = (object)phIn;
            this.txtctu0.Focus();
        }

        private void XGReport_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null || this.GridSearch.XGReport.ActiveRecord.RecordType != RecordType.DataRecord)
                return;
            if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1")
            {
                this.txtlien.IsReadOnly = true;
                this.txtlien.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
            }
            else
            {
                this.txtlien.IsReadOnly = false;
                this.txtlien.Text = StartUpTrans.DmctInfo["so_lien"].ToString().Trim();
            }
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

        private void BtnIn_Click(object sender, RoutedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null)
                return;
            this.hddt();
            if (this.txtlien.Value != null)
            {
                if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && !StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString().Equals("0"))
                {
                    if (ExMessageBox.Show(465, StartupBase.SasObj, "Hóa đơn đã được in, có muốn in lại hay không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                        return;
                    FrmLogin frmLogin = new FrmLogin();
                    frmLogin.ShowDialog();
                    if (!frmLogin.IsLogined)
                        return;
                }
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)Convert.ToInt16(this.txtlien.Text);
                int num = 1;
                int result1 = 0;
                int result2 = 0;
                int int32 = Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord)this.GridSearch.XGReport.ActiveRecord, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()));
                int.TryParse(this.dsSource.Tables[0].DefaultView[0]["so_lien_hd"].ToString(), out result1);
                int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result2);
                if (int32 > result1)
                    this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                int int16_1 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() != "1")
                    result1 = int16_1;
                int int16_2 = (int)Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]);
                if (int16_2 >= 1)
                    this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                for (; num <= int16_1; ++num)
                {
                    if (int32 <= result1)
                        this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = num > int32 && num <= result1 && int16_2 < 1 ? (object)"" : (object)"BẢN SAO";
                    this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)(result1 == 0 || num % result1 <= 0 ? result1 : num % result1);
                    this.InsertSubRow();
                    this.GridSearch.V_In((short)1, result2 >= num && string.IsNullOrEmpty(this.dsSource.Tables[0].DefaultView[0]["ban_sao"].ToString()));
                }
                if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && this.GridSearch.PrintSuccess)
                {
                    StartUpTrans.UpdateSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString(), ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["id"].ToString(), this.txtlien.Text);
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null)
                return;
            DataRowView dataItem = (this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView;
            string str1 = dataItem["mau_tu_in"].ToString();
            bool flag = false;
            for (int index = 1; index < this.dsSource.Tables[0].Rows.Count; ++index)
            {
                if (!this.dsSource.Tables[0].Rows[index]["tinh_trang_hddt"].ToString().Equals("0"))
                {
                    flag = true;
                    break;
                }
            }
            if (flag)
            {
                int num1 = (int)ExMessageBox.Show(396, StartupBase.SasObj, "Có chứng từ thuộc hóa đơn điện tử, không in liên tục được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else if (StartUp.IsQLHD && str1 == "1")
            {
                int num2 = (int)ExMessageBox.Show(470, StartupBase.SasObj, "Có chứng từ thuộc mẫu hóa đơn tự in, không in liên tục được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                if (StartUpTrans.M_IN_HOI_CK == 1 && ExMessageBox.Show(940, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc ?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                    return;
                if (dataItem["is_hddt"].ToString().Equals("1"))
                    ;
                if (this.txtlien.Value != null)
                {
                    if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && !StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString().Equals("0"))
                    {
                        if (ExMessageBox.Show(480, StartupBase.SasObj, "Hóa đơn đã được in, có muốn in lại hay không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                            return;
                        FrmLogin frmLogin = new FrmLogin();
                        frmLogin.ShowDialog();
                        if (!frmLogin.IsLogined)
                            return;
                        this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                    }
                    int int16_1 = (int)Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]);
                    if (int16_1 >= 1)
                        this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                    int iRow = FrmPoctpxf.iRow;
                    List<int> intList = new List<int>();
                    int result1 = 0;
                    int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result1);
                    int int16_2 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                    for (int index1 = 1; index1 < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index1)
                    {
                        for (int index2 = 1; index2 <= int16_2; ++index2)
                        {
                            string str2 = this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString();
                            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec = '" + str2 + "'";
                            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec = '" + str2 + "'";
                            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                            if (index2 == 1)
                                intList.Add(Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord)this.GridSearch.XGReport.ActiveRecord, this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString())));
                            if (this.dsSource.Tables[0].DefaultView[0]["status"].ToString() != "3")
                            {
                                int num3 = intList[index1 - 1];
                                int result2 = 0;
                                int.TryParse(this.dsSource.Tables[0].DefaultView[0]["so_lien_hd"].ToString(), out result2);
                                if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() != "1")
                                    result2 = int16_2;
                                this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = num3 <= result2 ? (index2 > num3 && index2 <= result2 ? (object)"" : (object)"BẢN SAO") : (object)"BẢN SAO";
                                if (int16_1 >= 1)
                                    this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)(result2 == 0 || index2 % result2 <= 0 ? result2 : index2 % result2);
                                this.InsertSubRow();
                                this.GridSearch.V_In(Convert.ToInt16(1), result1 >= index2 && string.IsNullOrEmpty(this.dsSource.Tables[0].DefaultView[0]["ban_sao"].ToString()));
                                if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && index2 == 1 && this.GridSearch.PrintSuccess)
                                {
                                    StartUpTrans.UpdateSl_in(this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString(), ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["id"].ToString(), this.txtlien.Text);
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
                                }
                            }
                        }
                    }
                    StartUpTrans.SetPhIn(this.DataContext as DataTable);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                }
                this.Close();
            }
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.BtnXem_Click(sender, (RoutedEventArgs)e);
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
            this.hddt();
            string str = StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
            if (!str.Equals("0"))
            {
                this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
            }
            if (Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]) >= (short)1)
                this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
            this.InsertSubRow();
            this.GridSearch.V_Xem();
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.dsSource.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.dsSource.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
        }

        private void hddt()
        {
            if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["is_hddt"].ToString().Equals("1") || !StartUp.M_SD_HDDT.Equals("1") || this.dsSource.Tables["TablePH"].DefaultView[0]["sd_hddt_yn"].ToString().Trim().Equals("0"))
                return;
            if (this.dsSource.Tables["TablePH"].DefaultView[0]["so_seri_hddt"].ToString().Trim() != string.Empty)
                this.dsSource.Tables["TablePH"].DefaultView[0]["so_seri"] = (object)this.dsSource.Tables["TablePH"].DefaultView[0]["so_seri_hddt"].ToString();
            if (this.dsSource.Tables["TablePH"].DefaultView[0]["so_ct_hddt"].ToString().Trim() != string.Empty)
                this.dsSource.Tables["TablePH"].DefaultView[0]["so_ct"] = (object)this.dsSource.Tables["TablePH"].DefaultView[0]["so_ct_hddt"].ToString();
            if (this.dsSource.Tables["TablePH"].DefaultView[0]["mau_hddt"].ToString().Trim() != string.Empty)
                this.dsSource.Tables["TablePH"].DefaultView[0]["mau_hd"] = (object)this.dsSource.Tables["TablePH"].DefaultView[0]["mau_hddt"].ToString();
            this.dsSource.Tables["TablePH"].AcceptChanges();
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
                        for (int i = 1; i < dsSource.Tables[0].Rows.Count; i++)
                        {
                            if (dsSource.Tables[0].Rows[i]["status"].ToString() == "3")
                            {
                                continue;
                            }
                            string sttlst = dsSource.Tables[0].Rows[i]["stt_rec"].ToString();
                            string filePDFName = GetFilePDFName(dsSource.Tables[0].Rows[i]);
                            dsSource.Tables[0].Rows[i]["so_ct_goc"] = (object)this.txtctu0.Text;
                            dsSource.Tables[0].AcceptChanges();
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
                                    dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + dsSource.Tables[0].Rows[i]["stt_rec"].ToString() + "'";
                                    dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + dsSource.Tables[0].Rows[i]["stt_rec"].ToString() + "'";
                                    dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                                    num++;
                                    num2 += 2;
                                }
                            }
                            GridSearch.DSource = dsSource;
                            GridSearch.V_XuatPdf_LT(text, filePDFName, handle, _loai);
                        }
                        flag = true;
                        if (_loai == 1)
                        {
                            string filePDFName1 = string.Format("{0}_{1}_{2}", DateTime.Now.ToString("yyyyMMddHHmmss"), dsSource.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim(), dsSource.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim());
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
