using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Data;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace COSXKSX.KSX
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
            this.dsSource = StartUpTrans.DsTrans.Copy();
            StartUp.GetDmnt(this.dsSource);
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
          
            this.dsSource.Tables[1].Columns.Add(new DataColumn("stt", typeof(int)));
            //DataRow row = this.dsSource.Tables[1].NewRow();          
            this.dsSource.Tables[2].Columns.Add(new DataColumn("stt", typeof(int)));
            //DataRow row1 = this.dsSource.Tables[2].NewRow();
            this.dsSource.Tables[3].Columns.Add(new DataColumn("stt", typeof(int)));
            //DataRow row2 = this.dsSource.Tables[3].NewRow();
            this.dsSource.Tables[4].Columns.Add(new DataColumn("stt", typeof(int)));
            //DataRow row3 = this.dsSource.Tables[4].NewRow();
            //row["stt_rec"] = row1["stt_rec"]= row2["stt_rec"]= row3["stt_rec"] =(object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            //this.dsSource.Tables[1].Rows.Add(row);
            //this.dsSource.Tables[2].Rows.Add(row1);
            //this.dsSource.Tables[3].Rows.Add(row2);
            //this.dsSource.Tables[4].Rows.Add(row3);
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            int num1 = 1;
         
            foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
            {               
                dataRowView["stt"] = (object)num1;
                ++num1;
            }
            foreach (DataRowView dataRowView in this.dsSource.Tables[2].DefaultView)
            {
                dataRowView["stt"] = (object)num1;
                ++num1;
            }
            foreach (DataRowView dataRowView in this.dsSource.Tables[3].DefaultView)
            {
                dataRowView["stt"] = (object)num1;
                ++num1;
            }
            foreach (DataRowView dataRowView in this.dsSource.Tables[4].DefaultView)
            {
                dataRowView["stt"] = (object)num1;
                ++num1;
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
                this.txtlien.Focus();
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
                int num = 1;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)int16;
                for (; num <= int16; ++num)
                {
                   // this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
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
                if (StartUp.M_IN_HOI_CK == 1 && ExMessageBox.Show(2210, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc ?", "SASERP 20 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                    return;
                int iRow = FrmPoctpna.iRow;
                int num1 = 1;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)num1;
                for (; num1 <= int16; ++num1)
                {
                    for (int index = 1; index < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index)
                    {
                        this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[3].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[4].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                       // this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                        int num2 = 1;
                        foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
                        {
                            dataRowView["stt"] = (object)num2;
                            ++num2;
                        }
                        this.GridSearch.V_In((short)1);
                    }
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[3].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                StartUpTrans.DsTrans.Tables[4].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
            }
            this.Close();
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
           // this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
            this.GridSearch.V_Xem();
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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
    }
}
