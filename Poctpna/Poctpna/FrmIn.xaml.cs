using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;

namespace Poctpna
{
    public partial class FrmIn : Form
    {
        private DataSet dsSource = new DataSet();

        public FrmIn()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.ShowInTaskbar = false;
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.GridSearch.ReportGroupName = StartUpTrans.CommandInfo["rep_file"].ToString();
            if (StartUpTrans.M_LAN != "V")
                this.Title = "Report form list";
            this.dsSource = StartUpTrans.DsTrans.Copy();
            StartUp.GetDmnt(this.dsSource);
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
            {
                DefaultValue = (object)0
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("ma_nx", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("stt", typeof(int)));
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
            this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            if (this.dsSource.Tables[2].DefaultView.Count > 0)
                this.dsSource.Tables[2].DefaultView[0]["so_ct0"] = (object)this.dsSource.Tables[2].DefaultView[0]["so_ct0"].ToString().Trim();
            int num = 1;
            foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
            {
                if (this.dsSource.Tables[0].DefaultView[0]["ma_gd"].ToString().IndexOfAny(new char[2]
                {
                  '2',
                  '5'
                }) >= 0)
                    dataRowView["tien"] = dataRowView["tien_tt"];
                dataRowView["ma_nx"] = (object)this.dsSource.Tables[0].DefaultView[0]["ma_nx"].ToString();
                dataRowView["stt"] = (object)num;
                ++num;
            }
            this.GridSearch.DSource = this.dsSource;
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Title = SysFunc.Cat_Dau(this.Title);
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
                this.txtctu0.Focus();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
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
            if (this.txtlien.Value != null)
            {
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                int num = 1;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)Convert.ToInt16(this.txtlien.Text.ToString());
                for (; num <= int16; ++num)
                {
                    this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                    this.GridSearch.V_In((short)1);
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.Value != null)
            {
                if (StartUp.M_IN_HOI_CK == 1 && ExMessageBox.Show(340, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc ?", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString().Trim(), MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                    return;
                int iRow = FrmPoctpna.iRow;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)int16;
                for (int index1 = 1; index1 < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index1)
                {
                    for (int index2 = 1; index2 <= int16; ++index2)
                    {
                        this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                        this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString() + "'";
                        int num = 1;
                        foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
                        {
                            if (!string.IsNullOrEmpty(dataRowView["ma_vt"].ToString().Trim()))
                            {
                                dataRowView["stt"] = (object)num;
                                ++num;
                            }
                        }
                        this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                        this.GridSearch.V_In((short)1);
                    }
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
            }
            this.Close();
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
            this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
            this.GridSearch.V_Xem();
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void txtctu0_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtctu0.IsFocusWithin)
                return;
            if (this.txtctu0.Value.ToString() == "")
                this.txtctu0.Value = (object)0;
            this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = this.txtctu0.Value;
        }

        private void GridSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
        }

        private void txtlien_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.IsFocusWithin || !(this.txtlien.Text.ToString() == ""))
                return;
            this.txtlien.Value = (object)0;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.dsSource.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.dsSource.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
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
