using Infragistics.Windows.DataPresenter;
using SasControls;
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
using System.Windows.Threading;

namespace CACTPC1
{
    public partial class FrmInBN1 : Form
    {
        private DataSet dsSource;

        public FrmInBN1()
        {
            this.InitializeComponent();
            if (DesignerProperties.GetIsInDesignMode((DependencyObject)this))
                return;
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
            this.dsSource.Tables[0].Columns.Add(new DataColumn("dia_chi_nh", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("MDTT", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("HTTT", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("TPTN", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("TPNN", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("CamKet", typeof(string))
            {
                DefaultValue = (object)""
            });
            StartUp.InsertInfoBank(ref this.dsSource);
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

        private void UpdateTenTKVN_EN()
        {
            SqlCommand cmd = new SqlCommand("select tk_i, ten_tk as ten_tk_i, ten_tk2 as ten_tk_i2  from ct56 a, dmtk b where a.tk_i=tk and stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim() + "'");
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
            //where o.Field<decimal?>("tien_nt") != null
            foreach (DataRowView row in this.dsSource.Tables[1].DefaultView)
            {
                if (String.IsNullOrEmpty(row["tien_nt"].ToString()))
                    row["tien_nt"] = 0;
                if (String.IsNullOrEmpty(row["tien"].ToString()))
                    row["tien"] = 0;
            }
            this.dsSource.Tables[1].AcceptChanges();
            foreach (DataRowView row in this.dsSource.Tables[2].DefaultView)
            {
                if (String.IsNullOrEmpty(row["t_thue_nt"].ToString()))
                    row["t_thue_nt"] = 0;
                if (String.IsNullOrEmpty(row["t_thue"].ToString()))
                    row["t_thue"] = 0;
            }

            this.dsSource.Tables[2].AcceptChanges();
            var enumerable = from o in this.dsSource.Tables[1].DefaultView.ToTable().AsEnumerable()
                             group o by o.Field<string>("tk_i") into g
                             select new
                             {
                                 TK = g.Key,
                                 Tien_nt = g.Sum((DataRow p) => p.Field<decimal>("tien_nt")),
                                 Tien = g.Sum((DataRow p) => p.Field<decimal>("tien"))
                             };

            var enumerable2 = from o in this.dsSource.Tables[2].DefaultView.ToTable().AsEnumerable()
                              group o by o.Field<string>("tk_thue_no") into g
                              select new
                              {
                                  TK = g.Key,
                                  Tien_nt = g.Sum((DataRow p) => p.Field<decimal>("t_thue_nt")),
                                  Tien = g.Sum((DataRow p) => p.Field<decimal>("t_thue"))
                              };

            DataTable dataTable = this.dsSource.Tables[1].Clone();

            if (enumerable.ToArray().Length > 0)
            {
                foreach (var data in enumerable)
                {
                    if (data.TK != null)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        dataRow["tk_i"] = data.TK.Trim();
                        dataRow["tien_nt"] = data.Tien_nt;
                        dataRow["tien"] = data.Tien;
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            if (enumerable2.ToArray().Length > 0)
            {
                foreach (var data2 in enumerable2)
                {
                    if (data2.TK != null)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        dataRow["tk_i"] = data2.TK.Trim();
                        dataRow["tien_nt"] = data2.Tien_nt;
                        dataRow["tien"] = data2.Tien;
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }

            var enumerable3 = from o in dataTable.AsEnumerable()
                              group o by o.Field<string>("tk_i") into g
                              select new
                              {
                                  TK = g.Key,
                                  Tien_nt = g.Sum((DataRow p) => p.Field<decimal>("tien_nt")),
                                  Tien = g.Sum((DataRow p) => p.Field<decimal>("tien"))
                              };


            DataTable dataTable2 = this.dsSource.Tables[1].Clone();
            dataTable2.TableName = "TablePsNo";

            if (enumerable3.ToArray().Length > 0)
            {
                foreach (var data3 in enumerable3)
                {
                    DataRow dataRow = dataTable2.NewRow();
                    dataRow["tk_i"] = data3.TK.Trim();
                    dataRow["tien_nt"] = data3.Tien_nt;
                    dataRow["tien"] = data3.Tien;
                    dataTable2.Rows.Add(dataRow);
                }
            }
            if (this.dsSource.Tables.Contains("TablePsNo"))
            {
                this.dsSource.Tables.Remove("TablePsNo");
            }
            this.dsSource.Tables.Add(dataTable2);
        }

        private void AddTT()
        {
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.txtNoiDung.Focus()), DispatcherPriority.Background);
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.AddTT();
            if (DesignerProperties.GetIsInDesignMode((DependencyObject)this))
                return;
            this.Title = SysFunc.Cat_Dau(this.Title);
            DataTable phIn = StartUpTrans.GetPhIn();
            if (phIn.Rows.Count == 0)
            {
                DataRow row = phIn.NewRow();
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["stt_rec"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                row["so01"] = (object)StartUp.so_ct0;
                row["gc01"] = (object)StartUp.dien_giai0;
                row["so02"] = StartUpTrans.DmctInfo["so_lien"] == DBNull.Value ? (object)1 : StartUpTrans.DmctInfo["so_lien"];
                if (this.dsSource.Tables[4].Rows.Count > 0)
                {
                    row["gc02"] = this.dsSource.Tables[4].Rows[0]["ten_nh"];
                    row["gc03"] = this.dsSource.Tables[4].Rows[0]["tk_nh"];
                }
                row["gc06"] = this.dsSource.Tables[0].DefaultView[0]["tk_nh"];
                row["gc07"] = this.dsSource.Tables[0].DefaultView[0]["ten_nh"];
                row["gc08"] = this.dsSource.Tables[0].DefaultView[0]["dia_chi_nh"];
                row["gc10"] = this.dsSource.Tables[0].DefaultView[0]["dien_giai"];
                row["gc11"] = this.dsSource.Tables[0].DefaultView[0]["MDTT"];
                row["gc12"] = this.dsSource.Tables[0].DefaultView[0]["HTTT"];
                row["gc13"] = this.dsSource.Tables[0].DefaultView[0]["TPTN"];
                row["gc14"] = this.dsSource.Tables[0].DefaultView[0]["TPNN"];
                row["gc15"] = this.dsSource.Tables[0].DefaultView[0]["CamKet"];
                phIn.Rows.Add(row);
            }
            else if (this.dsSource.Tables[4].Rows.Count > 0 && this.dsSource.Tables[4].Rows[0].RowState != DataRowState.Added)
            {
                phIn.Rows[0]["gc02"] = this.dsSource.Tables[4].Rows[0]["ten_nh"];
                phIn.Rows[0]["gc03"] = this.dsSource.Tables[4].Rows[0]["tk_nh"];
            }
            phIn.Rows[0]["gc04"] = this.dsSource.Tables[0].DefaultView[0]["ten_kh"];
            phIn.Rows[0]["gc05"] = this.dsSource.Tables[0].DefaultView[0]["dia_chi"];
            phIn.Rows[0]["gc09"] = this.dsSource.Tables[0].DefaultView[0]["tinh_thanh"];
            this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
            this.DataContext = (object)phIn;
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
                DataTable dataContext = this.DataContext as DataTable;
                if (dataContext.Rows.Count > 0)
                {
                    DataRow row = dataContext.Rows[0];
                    if (this.dsSource.Tables[4].Rows.Count > 0)
                    {
                        this.dsSource.Tables[4].Rows[0]["ten_nh"] = row["gc02"];
                        this.dsSource.Tables[4].Rows[0]["tk_nh"] = row["gc03"];
                    }
                    this.dsSource.Tables[0].DefaultView[0]["ten_kh"] = row["gc04"];
                    this.dsSource.Tables[0].DefaultView[0]["dia_chi"] = row["gc05"];
                    this.dsSource.Tables[0].DefaultView[0]["tk_nh"] = row["gc06"];
                    this.dsSource.Tables[0].DefaultView[0]["ten_nh"] = row["gc07"];
                    this.dsSource.Tables[0].DefaultView[0]["dia_chi_nh"] = row["gc08"];
                    this.dsSource.Tables[0].DefaultView[0]["tinh_thanh"] = row["gc09"];
                    this.dsSource.Tables[0].DefaultView[0]["dien_giai"] = row["gc10"];
                    this.dsSource.Tables[0].DefaultView[0]["MDTT"] = row["gc11"];
                    this.dsSource.Tables[0].DefaultView[0]["HTTT"] = row["gc12"];
                    this.dsSource.Tables[0].DefaultView[0]["TPTN"] = row["gc13"];
                    this.dsSource.Tables[0].DefaultView[0]["TPNN"] = row["gc14"];
                    this.dsSource.Tables[0].DefaultView[0]["CamKet"] = row["gc15"];
                }
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());

                while (num1 <= int16)
                {
                    this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)num1;
                    this.dsSource.Tables[0].DefaultView[0]["so_lienQD1548"] = (object)num2;
                    this.AddtbPsNo();
                    this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                    this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, "TablePsNo");
                    this.GridSearch.V_In((short)1);
                    ++num1;
                    num2 += 2;
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void AddTbTkPs(string sttrec)
        {
            SqlCommand cmd = new SqlCommand("select tk, tt, sum(tien) as tien, sum(tien_nt) as tien_nt  from( select tk_i as tk, sum(tien) as tien, sum(tien_nt) as tien_nt, 1 as tt from ct56 c join ph56 p ON c.stt_rec = p.stt_rec where c.stt_rec='" + sttrec + "' group by tk_i, p.ma_gd  union  select tk_thue_no as tk,sum(t_thue) as tien, sum(t_thue_nt) as tien_nt, 2 as tt from ct56gt where stt_rec='" + sttrec + "' group by tk_thue_no ) a group by tk, tt order by tt, tk");
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

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.M_IN_HOI_CK == 1 && ExMessageBox.Show(490, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc ?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                return;
            if (this.txtlien.Value != null)
            {
                int iRow = FrmCACTPC1.iRow;
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["dien_giai_ct_goc"] = (object)this.txtdiengiaict0.Text;
                DataTable dataContext = this.DataContext as DataTable;
                if (dataContext.Rows.Count > 0)
                {
                    DataRow row = dataContext.Rows[0];
                    if (this.dsSource.Tables[4].Rows.Count > 0)
                    {
                        this.dsSource.Tables[4].Rows[0]["ten_nh"] = row["gc02"];
                        this.dsSource.Tables[4].Rows[0]["tk_nh"] = row["gc03"];
                    }
                    this.dsSource.Tables[0].DefaultView[0]["ten_kh"] = row["gc04"];
                    this.dsSource.Tables[0].DefaultView[0]["dia_chi"] = row["gc05"];
                    this.dsSource.Tables[0].DefaultView[0]["tk_nh"] = row["gc06"];
                    this.dsSource.Tables[0].DefaultView[0]["ten_nh"] = row["gc07"];
                    this.dsSource.Tables[0].DefaultView[0]["dia_chi_nh"] = row["gc08"];
                    this.dsSource.Tables[0].DefaultView[0]["tinh_thanh"] = row["gc09"];
                    this.dsSource.Tables[0].DefaultView[0]["dien_giai"] = row["gc10"];
                    this.dsSource.Tables[0].DefaultView[0]["MDTT"] = row["gc11"];
                    this.dsSource.Tables[0].DefaultView[0]["HTTT"] = row["gc12"];
                    this.dsSource.Tables[0].DefaultView[0]["TPTN"] = row["gc13"];
                    this.dsSource.Tables[0].DefaultView[0]["TPNN"] = row["gc14"];
                    this.dsSource.Tables[0].DefaultView[0]["CamKet"] = row["gc15"];
                }
                int num1 = 1;
                int num2 = 1;
                int int16 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                while (num1 <= int16)
                {
                    for (int index = 1; index < this.dsSource.Tables[0].Rows.Count; ++index)
                    {
                        this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                        this.dsSource.Tables[2].DefaultView.RowFilter = "stt_rec= '" + this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString() + "'";
                        this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)int16;
                        this.dsSource.Tables[0].DefaultView[0]["so_lienQD1548"] = (object)num2;
                        this.AddTbTkPs(this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
                        this.AddtbPsNo();
                        this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
                        this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, "TablePsNo");
                        this.GridSearch.V_In((short)1);
                    }
                    ++num1;
                    num2 += 2;
                }
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
            }
            this.Close();
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.BtnXem_Click(sender, (RoutedEventArgs)e);
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.Value == null)
                return;
            this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
            this.dsSource.Tables[0].DefaultView[0]["dien_giai_ct_goc"] = (object)this.txtdiengiaict0.Text;
            this.dsSource.Tables[0].DefaultView[0]["so_lien"] = this.txtlien.Value;
            DataTable dataContext = this.DataContext as DataTable;
            if (dataContext.Rows.Count > 0)
            {
                DataRow row = dataContext.Rows[0];
                if (this.dsSource.Tables[4].Rows.Count > 0)
                {
                    this.dsSource.Tables[4].Rows[0]["ten_nh"] = row["gc02"];
                    this.dsSource.Tables[4].Rows[0]["tk_nh"] = row["gc03"];
                }
                this.dsSource.Tables[0].DefaultView[0]["ten_kh"] = row["gc04"];
                this.dsSource.Tables[0].DefaultView[0]["dia_chi"] = row["gc05"];
                this.dsSource.Tables[0].DefaultView[0]["tk_nh"] = row["gc06"];
                this.dsSource.Tables[0].DefaultView[0]["ten_nh"] = row["gc07"];
                this.dsSource.Tables[0].DefaultView[0]["dia_chi_nh"] = row["gc08"];
                this.dsSource.Tables[0].DefaultView[0]["tinh_thanh"] = row["gc09"];
                this.dsSource.Tables[0].DefaultView[0]["dien_giai"] = row["gc10"];
                this.dsSource.Tables[0].DefaultView[0]["MDTT"] = row["gc11"];
                this.dsSource.Tables[0].DefaultView[0]["HTTT"] = row["gc12"];
                this.dsSource.Tables[0].DefaultView[0]["TPTN"] = row["gc13"];
                this.dsSource.Tables[0].DefaultView[0]["TPNN"] = row["gc14"];
                this.dsSource.Tables[0].DefaultView[0]["CamKet"] = row["gc15"];
            }

            var rows = this.dsSource.Tables[1].Select("[tk_i] is null ");
            foreach (var row in rows)
            { row.Delete(); }
            this.dsSource.Tables[1].AcceptChanges();

            this.AddTbTkPs(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim());
            this.AddtbPsNo();
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, "TablePsNo");
            this.GridSearch.V_Xem();
            StartUpTrans.SetPhIn(this.DataContext as DataTable);
        }

        private void BtnThoat_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MaskedTextBox_PreviewLostKeyboardFocus(
          object sender,
          KeyboardFocusChangedEventArgs e)
        {
            MaskedTextBox maskedTextBox = sender as MaskedTextBox;
            if (!string.IsNullOrEmpty(maskedTextBox.Text.Trim()))
                return;
            maskedTextBox.Text = "1";
        }

        private void Form_Closing(object sender, CancelEventArgs e)
        {
            StartUp.so_ct0 = this.txtctu0.nValue;
            StartUp.dien_giai0 = this.txtdiengiaict0.Text;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.dsSource.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.dsSource.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
        }

    }
}
