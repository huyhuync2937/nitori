using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace Poctpxf
{
    /// <summary>
    /// Interaction logic for FrmPoctpxf_PN.xaml
    /// </summary>
    public partial class FrmPoctpxf_PN : Form
    {
        private int indexRow = 0;
        private DataRecord mainrecord = (DataRecord)null;
        private CodeValueBindingObject Voucher_Ma_nt0;
        private string currentMa_nt;
        public FrmPoctpxf_PN(DataTable tbSource, Record recordCT, CodeValueBindingObject VMa_nt0)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.Loaded += new RoutedEventHandler(this.FrmPoctpxf_PN_Loaded);
            this.GrdGia.DataSource = (IEnumerable)tbSource.DefaultView;
            DataRowView drVCT = (recordCT as DataRecord).DataItem as DataRowView;
            this.Title = this.Title = SysFunc.Cat_Dau(StartUpTrans.M_LAN.Equals("V") ? drVCT["ten_vt"].ToString() : drVCT["ten_vt2"].ToString());
            IEnumerable<DataRow> source = tbSource.AsEnumerable().Where<DataRow>((Func<DataRow, bool>)(b => b.Field<string>("stt_rec").Equals(drVCT["stt_rec_pn"].ToString()) && b.Field<string>("stt_rec0").Equals(drVCT["stt_rec0pn"].ToString()))).Select<DataRow, DataRow>((Func<DataRow, DataRow>)(c => c)).Take<DataRow>(1);
            if (source != null && ((IEnumerable<DataRow>)source.ToArray<DataRow>()).Count<DataRow>() > 0)
                this.indexRow = tbSource.Rows.IndexOf(source.ToArray<DataRow>()[0]);
            if (this.indexRow < 0)
                this.indexRow = 0;
            this.mainrecord = recordCT as DataRecord;
            this.currentMa_nt = VMa_nt0.Text;
        }

        private void FrmPoctpxf_PN_Loaded(object sender, RoutedEventArgs e)
        {
            this.Voucher_Ma_nt0 = (CodeValueBindingObject)this.FindResource((object)"Voucher_Ma_nt0");
            this.Voucher_Ma_nt0.Text = this.currentMa_nt;
            if (this.GrdGia.Records.Count > 0)
            {
                this.GrdGia.ActiveRecord = this.GrdGia.Records[this.indexRow];
                this.GrdGia.Focus();
            }
            this.SetVisibleField();
        }

        private void SetVisibleField()
        {
            if (this.Voucher_Ma_nt0.Text.Trim().Equals(StartUpTrans.M_ma_nt0.Trim()))
            {
                this.GrdGia.FieldLayouts[0].Fields["ma_nt"].Visibility = Visibility.Hidden;
                this.GrdGia.FieldLayouts[0].Fields["ma_nt"].Settings.CellMaxWidth = 0.0;
                this.GrdGia.FieldLayouts[0].Fields["ty_giaf"].Visibility = Visibility.Hidden;
                this.GrdGia.FieldLayouts[0].Fields["ty_giaf"].Settings.CellMaxWidth = 0.0;
                this.GrdGia.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Hidden;
                this.GrdGia.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = 0.0;
            }
            else
            {
                this.GrdGia.FieldLayouts[0].Fields["ma_nt"].Visibility = Visibility.Visible;
                this.GrdGia.FieldLayouts[0].Fields["ma_nt"].Settings.CellMaxWidth = this.GrdGia.FieldLayouts[0].Fields["ma_nt"].Width.Value.Value;
                this.GrdGia.FieldLayouts[0].Fields["ty_giaf"].Visibility = Visibility.Visible;
                this.GrdGia.FieldLayouts[0].Fields["ty_giaf"].Settings.CellMaxWidth = this.GrdGia.FieldLayouts[0].Fields["ty_giaf"].Width.Value.Value;
                this.GrdGia.FieldLayouts[0].Fields["gia_nt"].Visibility = Visibility.Visible;
                this.GrdGia.FieldLayouts[0].Fields["gia_nt"].Settings.CellMaxWidth = this.GrdGia.FieldLayouts[0].Fields["gia_nt"].Width.Value.Value;
            }
        }

        private void btnNhan_Click(object sender, RoutedEventArgs e)
        {
            if (this.GrdGia.ActiveRecord == null)
                return;
            DataRowView dataItem = (this.GrdGia.ActiveRecord as DataRecord).DataItem as DataRowView;
            Decimal result1 = new Decimal(0);
            Decimal result2 = new Decimal(0);
            Decimal result3 = new Decimal(0);
            Decimal result4 = new Decimal(0);
            Decimal.TryParse(this.mainrecord.Cells["so_luong"].Value.ToString(), out result1);
            Decimal.TryParse(dataItem["gia"].ToString(), out result2);
            Decimal.TryParse(dataItem["gia_nt"].ToString(), out result3);
            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result4);
            if (StartUpTrans.M_ma_nt0.Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString().Trim()))
            {
                this.mainrecord.Cells["gia_nt"].Value = (object)result2;
                this.mainrecord.Cells["gia"].Value = (object)result2;
                if (result1 != new Decimal(0) && result2 != new Decimal(0))
                    this.mainrecord.Cells["tien"].Value = (object)SysFunc.Round(result1 * result2, StartUpTrans.M_ROUND);
            }
            else
            {
                this.mainrecord.Cells["gia_nt"].Value = (object)result3;
                Decimal num = SysFunc.Round(result3 * result4, StartUpTrans.M_ROUND_GIA);
                this.mainrecord.Cells["gia"].Value = (object)num;
                if (result1 != new Decimal(0) && num != new Decimal(0))
                    this.mainrecord.Cells["tien"].Value = (object)SysFunc.Round(result1 * num, StartUpTrans.M_ROUND);
            }
            this.mainrecord.Cells["stt_rec_pn"].Value = dataItem["stt_rec"];
            this.mainrecord.Cells["stt_rec0pn"].Value = dataItem["stt_rec0"];
            this.Close();
        }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.Close();
        }

    }
}
