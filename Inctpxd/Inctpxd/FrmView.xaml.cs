using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Linq;
using System.Collections;
using System.Data;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace Inctpxd
{
    public partial class FrmView : Form
    {
        private DataTable tbPH;
        private DataTable tbCT;
        public FrmView(DataTable _tbPH, DataTable _tbCT)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.tbPH = _tbPH;
            this.tbCT = _tbCT;
            string[] strArray = StartUpTrans.CommandInfo["VBrowse1"].ToString().Trim().Split('|');
            if (StartUpTrans.M_LAN == "E")
                strArray = StartUpTrans.CommandInfo["EBrowse1"].ToString().Trim().Split('|');
            string strBrowse1 = strArray[0];
            string strBrowse2 = strArray[1];
            this.GrdBrowse.DataSource = (IEnumerable)this.tbPH.DefaultView;
            this.GrdBrowseCt.DataSource = (IEnumerable)this.tbCT.DefaultView;
            this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, strBrowse1, this.tbPH));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, strBrowse1);
            this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2, this.tbCT));
            SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, strBrowse2);
        }

        private void GrdBrowse_RecordActivated(object sender, RecordActivatedEventArgs e)
        {
            try
            {
                BasicGridView basicGridView = sender as BasicGridView;
                if (basicGridView.ActiveRecord == null || (basicGridView.ActiveRecord.Index < 0 || basicGridView.ActiveRecord.RecordType != RecordType.DataRecord))
                    return;
                this.tbCT.DefaultView.RowFilter = "";
                string str = "1 =1 ";
                this.tbCT.DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object)"" : (object)" and ", (object)"tag", (object)((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)["tag"].ToString());
                this.GrdBrowseCt.DataSource = (IEnumerable)this.tbCT.DefaultView;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               this.Grid0.pnlButton.btnOk.Content = StartUpTrans.M_LAN == "V" ? (object)"Chọn" : (object)"Select";
               if (this.GrdBrowse.Records.Count <= 0)
                   return;
               this.GrdBrowse.ActiveRecord = this.GrdBrowse.Records[0];
           }));
        }

        private void Form_KeyUp(object sender, KeyEventArgs e)
        {
            if (!e.Key.Equals((object)Key.Escape))
                return;
            this.Close();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (this.GrdBrowse.ActiveRecord == null || this.GrdBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
                return;
            try
            {
                if (this.GrdBrowse.ActiveRecord == null || this.GrdBrowseCt.Records.Count < 1)
                {
                    this.Close();
                    return;
                }
                DataRowView dataItem1 = (this.GrdBrowse.ActiveRecord as DataRecord).DataItem as DataRowView;
                if (dataItem1["ma_sp"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_sp"].ToString().Trim()) && dataItem1["ma_bpht"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_bpht"].ToString().Trim()) && dataItem1["so_lsx"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsx"].ToString().Trim()))
                {
                    this.Close();
                    return;
                }
                int count = StartUpTrans.DsTrans.Tables[1].DefaultView.Count;
                for (int index = 0; index < count; ++index)
                    StartUpTrans.DsTrans.Tables[1].DefaultView.Delete(0);
                StartUpTrans.DsTrans.Tables[1].AcceptChanges();
                for (int index = 0; index < this.GrdBrowseCt.Records.Count; ++index)
                {
                    DataRowView dataItem2 = (this.GrdBrowseCt.Records[index] as DataRecord).DataItem as DataRowView;
                    DataRow row = StartUpTrans.DsTrans.Tables[1].NewRow();
                    row["stt_rec"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"];
                    int result = 1;
                    string str = StartUpTrans.DsTrans.Tables[1].AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec") == StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())).Max<DataRow, string>((Func<DataRow, string>)(x => x.Field<string>("stt_rec0")));
                    if (str != null)
                    {
                        int.TryParse(str.ToString(), out result);
                        ++result;
                    }
                    row["stt_rec0"] = (object)string.Format(StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString(), (object)result);
                    row["ma_ct"] = (object)StartUpTrans.Ma_ct;
                    row["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                    row["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                    row["so_luong"] = (object)0;
                    row["gia_nt"] = (object)0;
                    row["tien_nt"] = (object)0;
                    row["gia"] = (object)0;
                    row["tien"] = (object)0;
                    row["ton13"] = (object)DBNull.Value;
                    row["ma_nx_i"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nx"];
                    row["ma_vt"] = dataItem2["ma_vt"];
                    row["ten_vt"] = dataItem2["ten_vt"];
                    row["ten_vt2"] = dataItem2["ten_vt2"];
                    row["ma_sp"] = dataItem2["ma_sp"];
                    row["ma_bpht_i"] = dataItem2["ma_bpht"];
                    row["so_lsx_i"] = dataItem2["so_lsx"];
                    row["sl_dm"] = dataItem2["sl_dm"];
                    row["dvt"] = dataItem2["dvt"];
                    row["ma_nx_i"] = dataItem2["tk_nvl"];
                    row["gia_ton"] = dataItem2["gia_ton"];
                    row["vt_ton_kho"] = dataItem2["vt_ton_kho"];
                    row["sua_tk_vt"] = dataItem2["sua_tk_vt"];
                    row["tk_vt_dmvt"] = dataItem2["tk_vt"];
                    row["tk_vt"] = dataItem2["tk_vt"];
                    row["sl_min"] = dataItem2["sl_min"];
                    row["dvt1"] = dataItem2["dvt1"];
                    row["he_so1"] = dataItem2["he_so1"];
                    row["so_luong1"] = 0;

                    StartUpTrans.DsTrans.Tables[1].Rows.Add(row);
                }
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_sp"] = dataItem1["ma_sp"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_sp"] = dataItem1["ten_sp"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_sp2"] = dataItem1["ten_sp2"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_bpht"] = dataItem1["ma_bpht"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_lsx"] = dataItem1["so_lsx"];
                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lkh"] = dataItem1["ngay_lkh"];
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            this.Close();
        }

        private void Grid0_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
