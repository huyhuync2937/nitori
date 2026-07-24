using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasVoucherLib;
using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;

namespace CACTPC1
{
    public partial class FrmIn : Form
    {
        private DataSet dsSource;

        public FrmIn()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.GridSearch.LocalSasObj = StartupBase.SasObj;
            this.GridSearch.ReportGroupName = StartUpTrans.CommandInfo["rep_file"].ToString();
            this.dsSource = StartUpTrans.DsTrans.Copy();
            StartUp.GetDmnt(this.dsSource);
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lienQD1548", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
            {
                DefaultValue = (object)0
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("dien_giai_ct_goc", typeof(string))
            {
                DefaultValue = (object)""
            });
            if (!this.dsSource.Tables[1].Columns.Contains("tien_in_nt"))
                this.dsSource.Tables[1].Columns.Add("tien_in_nt", typeof(Decimal));
            if (!this.dsSource.Tables[1].Columns.Contains("tien_in"))
                this.dsSource.Tables[1].Columns.Add("tien_in", typeof(Decimal));
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
            this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.UpdateTenTKVN_EN();
            this.GridSearch.DSource = this.dsSource;
            this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        private void UpdateTien_in()
        {
            string str = this.dsSource.Tables[0].DefaultView[0]["ma_gd"].ToString();
            foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
            {
                if (str == "8")
                {
                    Decimal dec1 = StartUp.ToDec(dataRowView["tien_tt"]);
                    Decimal dec2 = StartUp.ToDec(dataRowView["tien_tt_nt"]);
                    StartUp.ToDec(dataRowView["thue"]);
                    StartUp.ToDec(dataRowView["thue_nt"]);
                    if (StartUp.ToDec(dataRowView["thue_suat"]) >= new Decimal(0))
                    {
                        dataRowView["tien_in"] = (object)dec1;
                        dataRowView["tien_in_nt"] = (object)dec2;
                    }
                    else
                    {
                        dataRowView["tien_in"] = (object)dec1;
                        dataRowView["tien_in_nt"] = (object)dec2;
                    }
                }
            }
        }

        private void UpdateTenTKVN_EN()
        {
            string str1 = "ph46";
            string str2 = "ct46";
            if (StartUpTrans.Ma_ct != "PC1")
            {
                str1 = "ph56";
                str2 = "ct56";
            }
            string str3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
            SqlCommand cmd = new SqlCommand("select tk_i, ten_tk as ten_tk_i, ten_tk2 as ten_tk_i2  from " + str2 + " a, dmtk b where a.tk_i=tk and stt_rec= '" + str3 + "'");
            DataTable dataTable1 = new DataTable();
            DataTable dataTable2 = DataProvider.FillCommand(StartupBase.SasObj, cmd).Tables[0].Copy();
            for (int index1 = 0; index1 < dataTable2.Rows.Count; ++index1)
            {
                for (int index2 = 0; index2 < this.dsSource.Tables[1].Rows.Count; ++index2)
                {
                    if (this.dsSource.Tables[1].Rows[index2]["tk_i"].ToString().Trim().Equals(dataTable2.Rows[index1]["tk_i"].ToString().Trim()))
                    {
                        this.dsSource.Tables[1].Rows[index2]["ten_tk"] = dataTable2.Rows[index1]["ten_tk_i"];
                        this.dsSource.Tables[1].Rows[index2]["ten_tk2"] = dataTable2.Rows[index1]["ten_tk_i2"];
                    }
                }
            }
            this.dsSource.Tables[1].AcceptChanges();
        }


        private void AddtbPsNo()
        {
            var source1 = this.dsSource.Tables[1].DefaultView.ToTable().AsEnumerable().GroupBy<DataRow, string>((Func<DataRow, string>)(o => o.Field<string>("tk_i"))).Select(g =>
            {
                var data = new
                {
                    TK = g.Key,
                    Tien_nt = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("tien_nt"))),
                    Tien = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("tien")))
                };
                return data;
            });
            var source2 = this.dsSource.Tables[2].DefaultView.ToTable().AsEnumerable().GroupBy<DataRow, string>((Func<DataRow, string>)(o => o.Field<string>("tk_thue_no"))).Select(g =>
            {
                var data = new
                {
                    TK = g.Key,
                    Tien_nt = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_thue_nt"))),
                    Tien = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("t_thue")))
                };
                return data;
            });
            DataTable source3 = this.dsSource.Tables[1].Clone();
            if (source1.ToArray().Length > 0)
            {
                foreach (var data in source1)
                {
                    if (data.TK != null)
                    {
                        DataRow row = source3.NewRow();
                        row["tk_i"] = (object)data.TK.Trim();
                        row["tien_nt"] = (object)data.Tien_nt;
                        row["tien"] = (object)data.Tien;
                        source3.Rows.Add(row);
                    }
                }
            }
            if (source2.ToArray().Length > 0)
            {
                foreach (var data in source2)
                {
                    if (data.TK != null)
                    {
                        DataRow row = source3.NewRow();
                        row["tk_i"] = (object)data.TK.Trim();
                        row["tien_nt"] = (object)data.Tien_nt;
                        row["tien"] = (object)data.Tien;
                        source3.Rows.Add(row);
                    }
                }
            }
            var source4 = source3.AsEnumerable().GroupBy<DataRow, string>((Func<DataRow, string>)(o => o.Field<string>("tk_i"))).Select(g =>
            {
                var data = new
                {
                    TK = g.Key,
                    Tien_nt = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("tien_nt"))),
                    Tien = g.Sum<DataRow>((Func<DataRow, Decimal?>)(p => p.Field<Decimal?>("tien")))
                };
                return data;
            });
            DataTable table = this.dsSource.Tables[1].Clone();
            table.TableName = "TablePsNo";
            if (source4.ToArray().Length > 0)
            {
                foreach (var data in source4)
                {
                    DataRow row = table.NewRow();
                    row["tk_i"] = (object)data.TK.Trim();
                    row["tien_nt"] = (object)data.Tien_nt;
                    row["tien"] = (object)data.Tien;
                    table.Rows.Add(row);
                }
            }
            if (this.dsSource.Tables.Contains("TablePsNo"))
                this.dsSource.Tables.Remove("TablePsNo");
            this.dsSource.Tables.Add(table);
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
                    row["so01"] = (object)StartUp.so_ct0;
                    row["gc01"] = (object)(StartUp.dien_giai0 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dgiai_ct"].ToString());
                    row["so02"] = StartUpTrans.DmctInfo["so_lien"] == DBNull.Value ? (object)1 : StartUpTrans.DmctInfo["so_lien"];
                    phIn.Rows.Add(row);
                }
                this.DataContext = (object)phIn;
                this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
                this.txtdiengiaict0.Focus();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
                this.Close();
            }
        }

        private void AddTbTkPs(string sttrec)
        {
            string str1 = "ph46";
            string str2 = "ct46";
            string str3 = "ct46gt";
            if (StartUpTrans.Ma_ct != "PC1")
            {
                str1 = "ph56";
                str2 = "ct56";
                str3 = "ct56gt";
            }
            SqlCommand cmd = new SqlCommand("select tk, tt, sum(tien) as tien, sum(tien_nt) as tien_nt  from( select tk_i as tk, sum(tien) as tien, sum(tien_nt) as tien_nt, 1 as tt from " + str2 + " c join " + str1 + " p ON c.stt_rec = p.stt_rec where c.stt_rec='" + sttrec + "' group by tk_i, p.ma_gd  union  select tk_thue_no as tk,sum(t_thue) as tien, sum(t_thue_nt) as tien_nt, 2 as tt from " + str3 + " where stt_rec='" + sttrec + "' group by tk_thue_no ) a group by tk, tt order by tt, tk");
            DataTable dataTable = new DataTable();
            DataTable table = DataProvider.FillCommand(StartupBase.SasObj, cmd).Tables[0].Copy();
            table.TableName = "tbTaiKhoanPS";
            if (this.dsSource.Tables.Contains("tbTaiKhoanPS"))
            {
                this.dsSource.Tables["tbTaiKhoanPS"].Clear();
                foreach (DataRow row in (InternalDataCollectionBase)table.Rows)
                    this.dsSource.Tables["tbTaiKhoanPS"].Rows.Add(row.ItemArray);
            }
            else
                this.dsSource.Tables.Add(table);
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
                int num1 = 1;
                int num2 = 1;
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["dien_giai_ct_goc"] = (object)this.txtdiengiaict0.Text;
                if (this.GridSearch.XGReport.ActiveRecord != null)
                {
                    int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                    this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
                    while (num1 <= int16)
                    {
                        this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)num1;
                        this.dsSource.Tables[0].DefaultView[0]["so_lienQD1548"] = (object)num2;
                        this.AddtbPsNo();
                        this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                        this.GridSearch.V_In((short)1);
                        ++num1;
                        num2 += 2;
                    }
                    StartUpTrans.SetPhIn(this.DataContext as DataTable);
                }
            }
            this.Close();
        }

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.M_IN_HOI_CK == 1 && ExMessageBox.Show(820, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc ?", "Fast Book 11 .NET", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                return;
            if (this.txtlien.Value != null)
            {
                int iRow = FrmCACTPC1.iRow;
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["dien_giai_ct_goc"] = (object)this.txtdiengiaict0.Text;
                if (this.GridSearch.XGReport.ActiveRecord != null)
                {
                    int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                    for (int index = 1; index < this.dsSource.Tables[0].Rows.Count; ++index)
                    {
                        int num1 = 1;
                        int num2 = 1;
                        while (num1 <= int16)
                        {
                            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                            this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)num1;
                            this.dsSource.Tables[0].DefaultView[0]["so_lienQD1548"] = (object)num2;
                            this.AddTbTkPs(this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
                            this.AddtbPsNo();
                            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                            this.GridSearch.V_In((short)1);
                            ++num1;
                            num2 += 2;
                        }
                    }
                    StartUpTrans.SetPhIn(this.DataContext as DataTable);
                    StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Sort = "stt_rec0";
                }
            }
            this.Close();
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.BtnXem_Click(sender, (RoutedEventArgs)e);
        }

        private void txtNum_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as NumericTextBox).SelectAll();
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.Value == null)
                return;
            this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
            this.dsSource.Tables[0].DefaultView[0]["dien_giai_ct_goc"] = (object)this.txtdiengiaict0.Text;
            this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)1;
            this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
            this.AddtbPsNo();
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, "TablePsNo");
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
            this.GridSearch.V_Xem();
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Form_Closing(object sender, CancelEventArgs e)
        {
            StartUp.so_ct0 = this.txtctu0.nValue;
            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dgiai_ct"] = (object)(StartUp.dien_giai0 = this.txtdiengiaict0.Text);
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.dsSource.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.dsSource.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
        }
    }
}
