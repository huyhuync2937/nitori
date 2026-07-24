using Infragistics.Windows.Editors;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace FxTinhKH
{
    public partial class Ingia_tbLoc : FormFilter
    {
        public bool bResult = false;
        private DateTime _ngay_ct1;
        private DateTime _ngay_ct2;
        private DataTable dtMa_ts;

        public Ingia_tbLoc()
        {
            this.InitializeComponent();
        }

        public string Procedure { get; set; }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            DataTable dataContext = this.DataContext as DataTable;
            int period1 = 1;
            int period2 = 1;
            int result1 = 2017;
            try
            {
                period1 = (int)dataContext.Rows[0]["ky1"];
                period2 = (int)dataContext.Rows[0]["ky2"];
                result1 = (int)dataContext.Rows[0]["nam1"];
            }
            catch (Exception ex)
            {
            }
            int.TryParse(this.txtYear1.Value.ToString(), out result1);
            if (period1 == 0 || result1 == 0 || period2 == 0)
            {
                int num = (int)ExMessageBox.Show(1145, StartupBase.SasObj, "Kỳ từ 1 đến 12!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                if (period1 == 0)
                    this.txtPeriod1.Focus();
                else if (period2 == 0)
                {
                    this.txtPeriod2.Focus();
                }
                else
                {
                    if (result1 != 0)
                        return;
                    this.txtYear1.Focus();
                }
            }
            else
            {
                try
                {
                    this._ngay_ct1 = NgayTC.StartOfPeriod(result1, period1);
                    this._ngay_ct2 = NgayTC.StartOfPeriod(result1, period2);
                }
                catch (Exception ex)
                {
                    int num = (int)ExMessageBox.Show(1150, StartupBase.SasObj, "Kỳ tính giá không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtPeriod1.Focus();
                    return;
                }
                if (this._ngay_ct1 < StartUp.M_ngay_ct0)
                {
                    int num = (int)ExMessageBox.Show(1155, StartupBase.SasObj, "Năm phải lớn hơn hoặc bằng năm của kỳ mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtYear1.Focus();
                }
                else if (this._ngay_ct1 < StartUp.M_ngay_ks)
                {
                    int num = (int)ExMessageBox.Show(1160, StartupBase.SasObj, "Không thể tính khấu hao vào kỳ khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtYear1.Focus();
                }
                else if (this._ngay_ct1 > this._ngay_ct2)
                {
                    int num = (int)ExMessageBox.Show(1161, StartupBase.SasObj, "Từ kỳ đến kỳ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    this.txtPeriod1.Focus();
                }
                else
                {
                    StartUp.KeyFilter = string.IsNullOrEmpty(this.txtMaTs.Text) ? "1=1" : "dbo.InList(ma_cc_me, '" + this.txtMaTs.Text.Trim() + "', ',') = 1";
                    SqlCommand sqlcmd = new SqlCommand("SELECT SUM(Gt_Kh_ky) FROM CTKHCC WHERE Nam= @nam AND Ky Between @Ky1 and @Ky2 AND ma_dvcs LIKE @Ma_dvcs" + (string.IsNullOrEmpty(StartUp.KeyFilter) ? "" : " and " + StartUp.KeyFilter));
                    sqlcmd.Parameters.Add("@nam", SqlDbType.Int).Value = (object)result1;
                    sqlcmd.Parameters.Add("@ky1", SqlDbType.Int).Value = (object)period1;
                    sqlcmd.Parameters.Add("@ky2", SqlDbType.Int).Value = (object)period2;
                    sqlcmd.Parameters.Add("@Ma_dvcs", SqlDbType.VarChar).Value = (object)string.Format("{0}%", (object)this.M_MA_DVCS.ToString().Trim());
                    Decimal result2 = new Decimal(0);
                    Decimal.TryParse(StartupBase.SasObj.ExcuteScalar(sqlcmd).ToString(), out result2);
                    if (result2 > new Decimal(0))
                    {
                        if (ExMessageBox.Show(1165, this.BindingSasObj, string.Format("Đã tính khấu hao từ kỳ [{0}] đến kỳ [{1}], có tính lại không?", (object)period1, (object)period2), "Tinh khau hao tai san", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes)
                        {
                            this.bResult = false;
                        }
                        else
                        {
                            this.BindingSasObj.SetSysvar("M_THANG1", (object)period1);
                            this.bResult = true;
                        }
                    }
                    else
                    {
                        this.BindingSasObj.SetSysvar("M_THANG1", (object)period1);
                        this.bResult = true;
                    }
                    this.Hide();
                }
            }
        }

        private void FrmIngia_tbLoc_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtYear1.Focus();
            this.DateChanged();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               this.txtMa_dvcs.SearchInit();
               this.txtMa_dvcs_LostFocus((object)this.txtMa_dvcs, (RoutedEventArgs)null);
           }));
            this.dtMa_ts = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT CAST(0 as BIT) as tag, ma_cc_me, ten_cc, ten_cc2, ma_dvcs FROM dmcc ORDER BY ma_cc_me")).Tables[0];
            this.txtMaTs.Text = "";
        }

        protected override bool IsEnterToPassObject(object sender)
        {
            return sender is XamComboEditor || base.IsEnterToPassObject(sender);
        }

        protected override void OnClosed(EventArgs e)
        {
        }

        private void txt_LostFocus(object sender, RoutedEventArgs e)
        {
            this.DateChanged();
        }

        private void DateChanged()
        {
            DataTable dataContext = this.DataContext as DataTable;
            int result1 = 0;
            int result2 = 0;
            int result3 = 0;
            if (!int.TryParse(dataContext.Rows[0]["ky1"].ToString(), out result1) || !int.TryParse(dataContext.Rows[0]["ky2"].ToString(), out result3) || !int.TryParse(dataContext.Rows[0]["nam1"].ToString(), out result2))
                return;
            if (result1 <= 0 || result1 > 12 || (result3 <= 0 || result3 > 12) || result2 < 1900)
                return;
            try
            {
                this._ngay_ct1 = NgayTC.StartOfPeriod(result2, result1);
                this._ngay_ct2 = NgayTC.EndOfPeriod(result2, result3);
            }
            catch (Exception ex)
            {
                return;
            }
            this.lblNgay_ct1.Text = this._ngay_ct1.ToString((IFormatProvider)Thread.CurrentThread.CurrentCulture).Substring(0, 10);
            this.lblNgay_ct2.Text = this._ngay_ct2.ToString((IFormatProvider)Thread.CurrentThread.CurrentCulture).Substring(0, 10);
        }

        private void txtGotFocus(object sender, RoutedEventArgs e)
        {
            (sender as XamMaskedEditor).SelectAll();
        }

        private void txtDk_cl_LostFocus(object sender, RoutedEventArgs e)
        {
            MaskedTextBox maskedTextBox = sender as MaskedTextBox;
            if (!(maskedTextBox.Text.Trim() == ""))
                return;
            maskedTextBox.Text = maskedTextBox.InputMask.Substring(0, 1);
        }

        private void txtTinh_giatb_LostFocus(object sender, RoutedEventArgs e)
        {
            MaskedTextBox maskedTextBox = sender as MaskedTextBox;
            if (!(maskedTextBox.Text.Trim() == ""))
                return;
            maskedTextBox.Text = maskedTextBox.InputMask.Substring(0, 1);
        }

        private void txtMa_dvcs_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((sender as AutoCompleteTextBox).RowResult == null)
            {
                this.txtTen_dvcs.Text = "";
            }
            else
            {
                try
                {
                    this.txtTen_dvcs.Text = !(StartupBase.M_LAN == "V") ? this.txtMa_dvcs.RowResult["ten_dvcs2"].ToString() : this.txtMa_dvcs.RowResult["ten_dvcs"].ToString();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            }
        }

        private void txtMaTs_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.F11)
                return;
            this.OnF11();
        }

        private void OnF11()
        {
            string text = this.txtMaTs.Text;
            char[] chArray = new char[1] { ',' };
            foreach (object obj in text.Split(chArray))
            {
                foreach (DataRow dataRow in this.dtMa_ts.Select(string.Format("ma_cc_me = '{0}'", obj)))
                    dataRow["tag"] = (object)true;
            }
            DataTable dataTable = this.dtMa_ts.Copy();
            COTKTH2Dvcs cotktH2Dvcs = new COTKTH2Dvcs();
            dataTable.DefaultView.RowFilter = !string.IsNullOrEmpty(this.txtMa_dvcs.Text.Trim()) ? "ma_dvcs LIKE '" + this.txtMa_dvcs.Text + "'" : "1=1";
            cotktH2Dvcs.GrdCt.DataSource = (IEnumerable)dataTable.DefaultView;
            if (StartupBase.M_LAN != "V")
                cotktH2Dvcs.Title = "FX list";
            bool? nullable = cotktH2Dvcs.ShowDialog();
            if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                return;
            DataRow[] dataRowArray = dataTable.Select("tag = 1");
            string str = "";
            foreach (DataRow dataRow in dataRowArray)
                str = str + (str == "" ? "" : ",") + dataRow["ma_cc_me"].ToString().Trim();
            this.txtMaTs.Text = str;
        }

        private void txtPeriod1_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtPeriod1.Text.Trim()))
                return;
            this.txtPeriod1.Text = "1";
        }

        private void txtPeriod2_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(this.txtPeriod2.Text.Trim()))
                return;
            this.txtPeriod2.Text = "1";
        }
    }
}
