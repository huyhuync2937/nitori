using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
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

namespace ARCTHD1
{
    public partial class FrmIn : Form
    {
        private DataSet dsSource = new DataSet();
        private DataSet dsTmp = new DataSet();
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
            this.dsSource.Tables.Add(StartUp.GetDmnt().Copy());
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_lien", typeof(int))
            {
                DefaultValue = (object)1
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("so_ct_goc", typeof(int))
            {
                DefaultValue = (object)0
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("thue_suat", typeof(int))
            {
                DefaultValue = (object)0
            });
            this.dsSource.Tables[0].Columns.Add(new DataColumn("ban_sao", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[1].Columns.Add(new DataColumn("stt", typeof(string))
            {
                DefaultValue = (object)""
            });
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
            this.dsSource.Tables[0].TableName = "TablePH";
            this.dsSource.Tables[1].TableName = "TableCT";
            this.dsSource.Tables.Add(this.CreateTableInfo().Copy());
            this.dsSource.Tables.Add(FrmIn.CreateTableMST(StartUp.M_MA_THUE, "TableMST_NB").Copy());
            this.dsSource.Tables.Add(FrmIn.CreateTableMST(this.dsSource.Tables["TablePH"].Rows[FrmArcthd1.iRow]["ma_so_thue"].ToString().TrimEnd(), "TableMST_NM").Copy());
            if (!this.dsSource.Tables["TableCT"].Columns.Contains("dvt"))
                this.dsSource.Tables["TableCT"].Columns.Add("dvt", typeof(string));
            if (!this.dsSource.Tables["TableCT"].Columns.Contains("so_luong"))
                this.dsSource.Tables["TableCT"].Columns.Add("so_luong", typeof(int));
            int count = this.dsSource.Tables["TableCT"].DefaultView.Count;
            for (int index = 0; index < count; ++index)
                this.dsSource.Tables["TableCT"].DefaultView[index]["stt"] = (object)(index + 1);
            this.dsTmp = this.dsSource.Copy();
            this.dsTmp.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsTmp.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsTmp.Tables[1].DefaultView.Sort = "stt_rec0";
            this.UpdateTenTD();
            this.GridSearch.DSource = this.dsSource;
            this.GridSearch.ReportPreviewMouseDoubleClick += new ControlFilterReport.MouseClick(this.GridSearch_ReportPreviewMouseDoubleClick);
            if (this.BindingSasObj.GetOption("M_LAN").ToString().Equals("V"))
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar("M_EXPORT_SIGN").ToString();
            else
                this.btnExport.Content = (object)this.BindingSasObj.GetSysvar2("M_EXPORT_SIGN").ToString();
        }

        public static DataTable CreateTableMST(string ma_so_thue, string tablename)
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = tablename;
            int length = ma_so_thue.Length;
            int startIndex;
            int num;
            for (startIndex = 0; startIndex < length; ++startIndex)
            {
                num = startIndex + 1;
                dataTable.Columns.Add(new DataColumn("m" + num.ToString(), typeof(string))
                {
                    DefaultValue = (object)ma_so_thue.Substring(startIndex, 1)
                });
            }
            for (; startIndex < 14; ++startIndex)
            {
                num = startIndex + 1;
                dataTable.Columns.Add(new DataColumn("m" + num.ToString(), typeof(string))
                {
                    DefaultValue = (object)""
                });
            }
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        private void UpdateTenTD()
        {
            SqlCommand cmd = new SqlCommand("select ma_phi_i,ma_td_i,ma_td2_i,ma_td3_i,ten_phi, ten_td , ten_td2 ,ten_td3 from v_ct21 where  stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim() + "'");
            DataTable dataTable1 = new DataTable();
            DataTable dataTable2 = DataProvider.FillCommand(StartupBase.SasObj, cmd).Tables[0].Copy();
            for (int index1 = 0; index1 < dataTable2.Rows.Count; ++index1)
            {
                for (int index2 = 0; index2 < this.dsSource.Tables[1].Rows.Count; ++index2)
                {
                    if (this.dsSource.Tables[1].Columns.Contains("ten_phi") && this.dsSource.Tables[1].Rows[index2]["ma_phi_i"].ToString().Trim().Equals(dataTable2.Rows[index1]["ma_phi_i"].ToString().Trim()))
                        this.dsSource.Tables[1].Rows[index2]["ten_phi"] = dataTable2.Rows[index1]["ten_phi"];
                    if (this.dsSource.Tables[1].Columns.Contains("ten_td") && this.dsSource.Tables[1].Rows[index2]["ma_td_i"].ToString().Trim().Equals(dataTable2.Rows[index1]["ma_td_i"].ToString().Trim()))
                        this.dsSource.Tables[1].Rows[index2]["ten_td"] = dataTable2.Rows[index1]["ten_td"];
                    if (this.dsSource.Tables[1].Columns.Contains("ten_td2") && this.dsSource.Tables[1].Rows[index2]["ma_td2_i"].ToString().Trim().Equals(dataTable2.Rows[index1]["ma_td2_i"].ToString().Trim()))
                        this.dsSource.Tables[1].Rows[index2]["ten_td2"] = dataTable2.Rows[index1]["ten_td2"];
                    if (this.dsSource.Tables[1].Columns.Contains("ten_td3") && this.dsSource.Tables[1].Rows[index2]["ma_td3_i"].ToString().Trim().Equals(dataTable2.Rows[index1]["ma_td3_i"].ToString().Trim()))
                        this.dsSource.Tables[1].Rows[index2]["ten_td3"] = dataTable2.Rows[index1]["ten_td3"];
                }
            }
            this.dsSource.Tables[1].AcceptChanges();
        }

        private DataTable CreateTableInfo()
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = "TableInfo";
            dataTable.Columns.Add(new DataColumn("M_PHONE", typeof(string))
            {
                DefaultValue = (object)StartUp.M_PHONE
            });
            dataTable.Columns.Add(new DataColumn("M_MST", typeof(string))
            {
                DefaultValue = (object)StartUp.M_MA_THUE
            });
            dataTable.Columns.Add(new DataColumn("M_Ten_CTY", typeof(string))
            {
                DefaultValue = (object)StartupBase.SasObj.GetSysvar("M_Ten_CTY").ToString().ToUpper()
            });
            dataTable.Columns.Add(new DataColumn("M_DIA_CHI", typeof(string))
            {
                DefaultValue = (object)StartupBase.SasObj.GetSysvar("M_DIA_CHI").ToString()
            });
            dataTable.Columns.Add(new DataColumn("M_TK_NH", typeof(string))
            {
                DefaultValue = (object)StartupBase.SasObj.GetOption("M_TK_NH").ToString()
            });
            dataTable.Columns.Add(new DataColumn("M_CUC_THUE", typeof(string))
            {
                DefaultValue = (object)StartupBase.SasObj.GetOption("M_CUC_THUE").ToString().ToUpper()
            });
            DataRow row = dataTable.NewRow();
            dataTable.Rows.Add(row);
            return dataTable;
        }

        private void InsertSubRow()
        {
            this.dsSource.Tables[0].DefaultView[0]["t_tien_sau_ck_nt"] = this.dsSource.Tables[0].DefaultView[0]["t_tien_nt2"];
            this.dsSource.Tables[0].DefaultView[0]["t_tien_sau_ck"] = this.dsSource.Tables[0].DefaultView[0]["t_tien2"];
            this.dsSource.Tables[0].DefaultView[0]["so_seri"] = this.dsSource.Tables[0].DefaultView[0]["so_seri"];
            foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
            {
                if (this.dsSource.Tables[0].DefaultView[0]["ma_gd"].ToString().IndexOfAny(new char[2]
                {
          '2',
          '5'
                }) >= 0)
                    dataRowView["tien"] = dataRowView["tien_tt"];
            }
            this.dsSource.Tables[0].DefaultView[0]["thue_suat"] = this.dsSource.Tables[1].DefaultView[0]["thue_suati"];
            if (this.dsSource.Tables[1].DefaultView.Cast<DataRowView>().All<DataRowView>((Func<DataRowView, bool>)(x => x["stt_rec0"].ToString() != "-1")) && (Convert.ToDecimal(this.dsSource.Tables[0].DefaultView[0]["t_ck_nt"] is DBNull ? (object)0 : this.dsSource.Tables[0].DefaultView[0]["t_ck_nt"]) > new Decimal(0) || Convert.ToDecimal(this.dsSource.Tables[0].DefaultView[0]["t_ck"] is DBNull ? (object)0 : this.dsSource.Tables[0].DefaultView[0]["t_ck"]) > new Decimal(0)))
            {
                DataRecord activeRecord = this.GridSearch.XGReport.ActiveRecord as DataRecord;
                DataRow row = this.dsSource.Tables[1].NewRow();
                row["stt_rec"] = (object)this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString();
                row["stt_rec0"] = (object)-1;
                row["ten_vt"] = (object)"Chiết khấu";
                if (activeRecord != null)
                    row["dien_giaii"] = activeRecord.Cells["lan"].Value.ToString().Trim().ToLower().Equals("anh") ? (object)"Discount" : (object)"Chiết khấu";
                row["tien_nt2"] = this.dsSource.Tables[0].DefaultView[0]["t_ck_nt"];
                row["tien2"] = this.dsSource.Tables[0].DefaultView[0]["t_ck"];
                this.dsSource.Tables[1].Rows.Add(row);
            }
            this.dsSource.Tables[1].AcceptChanges();
            this.dsSource.Tables[0].DefaultView[0]["thue_suat"] = this.dsSource.Tables[1].DefaultView[0]["thue_suati"];
            this.GridSearch.InsertSubRow(StartUpTrans.Ma_ct, 1);
        }

        private void ResetData()
        {
            this.dsSource = this.dsTmp.Copy();
            this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'";
            this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
            this.UpdateTenTD();
            this.GridSearch.DSource = this.dsSource;
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtHTTT.Visibility = Visibility.Collapsed;
            this.lbhttt.Visibility = Visibility.Collapsed;
            DataTable phIn = StartUpTrans.GetPhIn();
            if (phIn.Rows.Count == 0)
            {
                DataRow row = phIn.NewRow();
                if (String.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim()))
                {
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"] = (object)0;
                }    
                row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                row["stt_rec"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
                row["so01"] = (object)0;
                row["so02"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
                row["gc01"] = (object)FrmArcthd1.hinhthuc_tt;
                phIn.Rows.Add(row);
            }
            if (phIn.Rows.Count == 1)
            {
                DataRow row = phIn.Rows[0];
                row["so02"] = !this.IsND51 ? StartUpTrans.DmctInfo["so_lien"] : (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lien_hd"].ToString().Trim();
                row["gc01"] = (object)FrmArcthd1.hinhthuc_tt;
            }
            this.GridSearch.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.GridSearch.XGReport.RecordActivated += new EventHandler<RecordActivatedEventArgs>(this.XGReport_RecordActivated)));
            this.DataContext = (object)phIn;
            this.txtlien.Focus();
        }

        private void XGReport_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null || this.GridSearch.XGReport.ActiveRecord.RecordType != RecordType.DataRecord)
                return;
            if (this.IsPhieuht())
                this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
            else
            {
                this.GridSearch.DSource = this.dsSource;
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
            if (this.IsPhieuht())
            {
                this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                this.GridSearch.V_Xem();
            }
            else
            {
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
                    this.dsSource.Tables[0].DefaultView[0]["ht_tt"] = (object)this.txtHTTT.Text;
                    this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)Convert.ToInt16(this.txtlien.Text);
                    int num = 1;
                    int result1 = 0;
                    int result2 = 0;
                    int int32 = Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord)this.GridSearch.XGReport.ActiveRecord, StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()));
                    int.TryParse(this.dsSource.Tables[0].DefaultView[0]["so_lien_hd"].ToString(), out result1);
                    int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result2);
                    if (int32 > result1)
                        this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                    int int16_1 = (int)Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]);
                    if (int16_1 >= 1)
                        this.dsSource.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                    int int16_2 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                    if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() != "1")
                        result1 = int16_2;
                    for (; num <= int16_2; ++num)
                    {
                        if (int32 <= result1)
                            this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = num > int32 && num <= result1 && int16_1 < 1 ? (object)"" : (object)"BẢN SAO";
                        this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)(result1 == 0 || num % result1 <= 0 ? result1 : num % result1);
                        this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                        this.dsSource.Tables[0].DefaultView[0]["ht_tt"] = (object)this.txtHTTT.Text;
                        this.InsertSubRow();
                        this.GridSearch.V_In((short)1, result2 >= num && string.IsNullOrEmpty(this.dsSource.Tables[0].DefaultView[0]["ban_sao"].ToString()));
                    }
                    if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && this.GridSearch.PrintSuccess)
                    {
                        StartUpTrans.UpdateSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString(), ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["id"].ToString(), this.txtlien.Text);
                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
                    }
                    this.ResetData();
                    StartUpTrans.SetPhIn(this.DataContext as DataTable);
                }
                this.Close();
            }
        }

        private void BtnInLT_Click(object sender, RoutedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null)
                return;
            string str1 = ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["mau_tu_in"].ToString();
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
                if (StartUpTrans.M_IN_HOI_CK == 1 && ExMessageBox.Show(475, StartupBase.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                    return;
                if (this.IsPhieuht())
                {
                    for (int index = 1; index < this.dsSource.Tables[0].Rows.Count; ++index)
                    {
                        this.GridSearch.DSource = StartUp.GetPhieuht(this.dsSource.Tables[0].Rows[index]["stt_rec"].ToString());
                        this.GridSearch.V_In((short)1);
                    }
                }
                else
                {
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
                        int iRow = FrmArcthd1.iRow;
                        List<int> intList = new List<int>();
                        int result1 = 0;
                        int.TryParse(StartUpTrans.DmctInfo["so_lien_xac_minh"].ToString(), out result1);
                        DataTable table = StartUpTrans.DsTrans.Tables[1];
                        int int16_1 = (int)Convert.ToInt16(Math.Ceiling(Convert.ToDouble(this.txtlien.Text) / Convert.ToDouble((this.GridSearch.XGReport.ActiveRecord as DataRecord).Cells["so_lien"].Value)));
                        for (int index1 = 1; index1 < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index1)
                        {
                            for (int index2 = 1; index2 <= int16_1; ++index2)
                            {
                                string str2 = this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString();
                                this.dsSource.Tables[0].DefaultView.RowFilter = "stt_rec = '" + str2 + "'";
                                this.dsSource.Tables[1].DefaultView.RowFilter = "stt_rec = '" + str2 + "'";
                                this.dsSource.Tables[1].DefaultView.Sort = "stt_rec0";
                                this.dsSource.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                if (index2 == 1)
                                    intList.Add(Convert.ToInt32(StartUpTrans.GetSo_lien((DataRecord)this.GridSearch.XGReport.ActiveRecord, this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString())));
                                if (this.dsSource.Tables[0].DefaultView[0]["status"].ToString() != "3")
                                {
                                    int num3 = intList[index1 - 1];
                                    int result2 = 0;
                                    int.TryParse(this.dsSource.Tables[0].DefaultView[0]["so_lien_hd"].ToString(), out result2);
                                    if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() != "1")
                                        result2 = int16_1;
                                    int int16_2 = (int)Convert.ToInt16(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"]);
                                    if (int16_2 >= 1)
                                        this.dsSource.Tables["TablePH"].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                                    this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = num3 <= result2 ? (index2 > num3 && index2 <= result2 && int16_2 < 1 ? (object)"" : (object)"BẢN SAO") : (object)"BẢN SAO";
                                    this.dsSource.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                    this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                                    this.dsSource.Tables[0].DefaultView[0]["so_lien"] = (object)(result2 == 0 || index2 % result2 <= 0 ? result2 : index2 % result2);
                                    this.InsertSubRow();
                                    int num4 = 1;
                                    foreach (DataRowView dataRowView in this.dsSource.Tables[1].DefaultView)
                                    {
                                        dataRowView["stt"] = dataRowView["stt_rec0"] != (object)"999" ? (object)num4 : (object)DBNull.Value;
                                        ++num4;
                                    }
                                    this.dsSource.Tables[0].DefaultView[0]["stt_rec"].ToString();
                                    this.dsSource.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                    this.GridSearch.V_In(Convert.ToInt16(1), result1 >= index2 && string.IsNullOrEmpty(this.dsSource.Tables[0].DefaultView[0]["ban_sao"].ToString()));
                                    this.dsSource.Tables[1].DefaultView[0]["stt_rec"].ToString();
                                    if (((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["nd51"].ToString() == "1" && index2 == 1 && this.GridSearch.PrintSuccess)
                                    {
                                        StartUpTrans.UpdateSl_in(this.dsSource.Tables[0].Rows[index1]["stt_rec"].ToString(), ((this.GridSearch.XGReport.ActiveRecord as DataRecord).DataItem as DataRowView)["id"].ToString(), this.txtlien.Text);
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
                                    }
                                }
                                this.dsSource.Tables[1].DefaultView[0]["stt_rec"].ToString();
                            }
                        }
                        StartUpTrans.SetPhIn(this.DataContext as DataTable);
                        this.ResetData();
                        StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[iRow]["stt_rec"].ToString() + "'";
                    }
                    this.Close();
                }
            }
        }

        private void GridSearch_ReportPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.BtnXem_Click(sender, (RoutedEventArgs)e);
        }

        private void BtnXem_Click(object sender, RoutedEventArgs e)
        {
            if (this.GridSearch.XGReport.ActiveRecord == null)
                return;
            if (this.IsPhieuht())
            {
                this.GridSearch.DSource = StartUp.GetPhieuht(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                this.GridSearch.V_Xem();
            }
            else
            {
                this.hddt();
                string str = StartUpTrans.GetSl_in(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()).ToString();
                if (!str.Equals("0"))
                {
                    this.dsSource.Tables[0].DefaultView[0]["ban_sao"] = (object)"BẢN SAO";
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
                }
                this.dsSource.Tables[0].DefaultView[0]["so_ct_goc"] = (object)this.txtctu0.Text;
                this.dsSource.Tables[0].DefaultView[0]["ht_tt"] = (object)this.txtHTTT.Text;
                this.InsertSubRow();
                this.GridSearch.V_Xem();
                this.ResetData();
                StartUpTrans.SetPhIn(this.DataContext as DataTable);
                this.ht_tt();
            }
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

        private void txtlien_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtlien.IsFocusWithin || !(this.txtlien.Value.ToString() == ""))
                return;
            this.txtlien.Value = (object)0;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            if (this.dsSource.Tables[0].DefaultView.Count != 1)
                return;
            this.GridSearch.V_XuatPdf(StartUpTrans.GetFileNameExportWithSignature(this.dsSource.Tables[0].DefaultView[0]), new WindowInteropHelper((Window)this).Handle);
        }

        private bool IsPhieuht()
        {
            return (((DataRecord)this.GridSearch.XGReport.ActiveRecord).DataItem as DataRowView)["nhom_bc"].ToString().Trim() == "VcIn";
        }

        private void txtHTTT_LostFocus(object sender, RoutedEventArgs e)
        {
            this.dsSource.Tables[0].DefaultView[0]["ht_tt"] = (object)this.txtHTTT.Text;
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

        private void ht_tt()
        {
            SqlCommand sqlcmd = new SqlCommand("update phin set ht_tt= @ht_tt where stt_rec=@stt_rec");
            sqlcmd.Parameters.Add("@ht_tt", SqlDbType.NVarChar).Value = (object)this.txtHTTT.Text;
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.NVarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
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
