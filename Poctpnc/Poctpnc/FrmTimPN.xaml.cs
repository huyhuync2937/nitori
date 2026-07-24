using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using SasLib;
using System;
using System.Data;
using System.Windows;
using System.Windows.Input;

namespace Poctpnc
{
    public partial class FrmTimPN : Form
    {
        public bool isOk = false;
        public string makh = string.Empty;
        public FrmView frm;
        public DataSet dsPn;

        public FrmTimPN(string makh)
        {
            this.InitializeComponent();
            this.BindingSasObj = StartupBase.SasObj;
            this.makh = makh;
            SysFunc.LoadIcon((Window)this);
        }

        public SasObject SasObj
        {
            get
            {
                return (SasObject)this.GetValue(Form.SasObjProperty);
            }
            set
            {
                this.SetValue(Form.SasObjProperty, (object)value);
            }
        }

        private void FrmTim_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtNgay_ct2.Value = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
            this.txtNgay_ct1.Focus();
        }

        private string GetPhFilterExpr()
        {
            int databaseFieldLength = this.BindingSasObj.GetDatabaseFieldLength("so_ct");
            string str = " 1=1 " + " and ma_dvcs like '" + StartupBase.SasObj.DmdvcsInfo.Rows[0][0].ToString() + "%'";
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text))
                str = str + " and ngay_ct >= " + this.ConvertDataToSql(this.txtNgay_ct1.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text))
                str = str + " and ngay_ct <= " + this.ConvertDataToSql(this.txtNgay_ct2.Value, typeof(DateTime));
            if (!string.IsNullOrEmpty(this.txtSo_ct1.Text))
                str = str + " and so_ct >= '" + this.txtSo_ct1.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtSo_ct2.Text))
                str = str + " and so_ct <= '" + this.txtSo_ct2.Text.Trim().PadLeft(databaseFieldLength, ' ') + "'";
            if (!string.IsNullOrEmpty(this.txtMa_kh.Text))
                str = str + " and ma_kh = " + this.ConvertDataToSql((object)this.txtMa_kh.Text.Trim(), typeof(string));
            return str;
        }

        public string ConvertDataToSql(object value, Type ValueType)
        {
            string str;
            switch (ValueType.ToString())
            {
                case "System.String":
                    str = string.Format("'{0}'", (object)(value as string).Replace("'", "'"));
                    break;
                case "System.DateTime":
                    str = string.Format("'{0}'", (object)((DateTime)value).ToString("yyyyMMdd"));
                    break;
                default:
                    str = string.Format("'{0}'", value);
                    break;
            }
            return str;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (!this.IsHopLe())
                return;
            this.Close();
            this.frm = new FrmView(this.GetPhFilterExpr());
            if (StartUpTrans.M_LAN != "V")
                this.frm.Title = "Voucher list";
            this.frm.ShowDialog();
            this.isOk = this.frm.isOk;
            this.dsPn = this.frm.dsPn;
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            this.Close();
        }

        private Decimal ParseDecimal(object obj, Decimal defaultvalue)
        {
            Decimal result = new Decimal(0);
            Decimal.TryParse(obj != null ? obj.ToString() : defaultvalue.ToString(), out result);
            return result;
        }

        private bool IsHopLe()
        {
            if (!string.IsNullOrEmpty(this.txtNgay_ct1.Text) && !this.txtNgay_ct1.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(105, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct1.Focus();
                this.txtNgay_ct1.SelectAll();
                return false;
            }
            if (!string.IsNullOrEmpty(this.txtNgay_ct2.Text) && !this.txtNgay_ct2.IsValueValid)
            {
                int num = (int)ExMessageBox.Show(110, StartupBase.SasObj, "Ngày lọc chứng từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                this.txtNgay_ct2.Focus();
                this.txtNgay_ct2.SelectAll();
                return false;
            }
            if (string.IsNullOrEmpty(this.txtNgay_ct1.Text) || string.IsNullOrEmpty(this.txtNgay_ct2.Text) || !((DateTime)this.txtNgay_ct1.Value > (DateTime)this.txtNgay_ct2.Value))
                return true;
            int num1 = (int)ExMessageBox.Show(980, StartupBase.SasObj, "Ngày lọc chứng từ chưa hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            this.txtNgay_ct1.Focus();
            this.txtNgay_ct1.SelectAll();
            return false;
        }

        private void FrmTim_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape)
                return;
            this.isOk = false;
            this.Close();
        }

        private void txtMa_kh_LostFocus(object sender, RoutedEventArgs e)
        {
            if (this.txtMa_kh.RowResult == null)
                this.tblten_kh.Text = "";
            else
                this.tblten_kh.Text = !(StartUpTrans.M_LAN == "V") ? this.txtMa_kh.RowResult["ten_kh2"].ToString() : this.txtMa_kh.RowResult["ten_kh"].ToString();
        }
    }
}
