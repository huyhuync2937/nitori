using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Poctpnc
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmPoctpnc.IsInEditMode == null || !FrmPoctpnc.IsInEditMode.Value || values == null)
                return (object)true;
            string[] strArray = parameter.ToString().Trim().Split(';');
            switch (strArray[0].ToUpper())
            {
                case "PH":
                    bool result1 = false;
                    bool.TryParse(values[0].ToString(), out result1);
                    switch (strArray[1])
                    {
                        case "ong_ba":
                            return result1 && StartUpTrans.M_ong_ba.Equals("1") ? (object)false : (object)true;
                        case "ngay_lct":
                            return result1 && StartUpTrans.M_ngay_lct.Equals("1") ? (object)false : (object)true;
                    }
                    break;
                case "CT":
                    switch (strArray[1])
                    {
                        case "tien":
                            if (!values[0].Equals(DependencyProperty.UnsetValue) && !values[1].Equals(DependencyProperty.UnsetValue))
                            {
                                Decimal num1 = System.Convert.ToDecimal(values[0]);
                                Decimal num2 = new Decimal(1);
                                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"] != DBNull.Value)
                                    num2 = System.Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"]);
                                if (num1 * num2 == new Decimal(0))
                                    return (object)false;
                            }
                            return !values[2].Equals(DependencyProperty.UnsetValue) && (bool)values[2] ? (object)false : (object)true;
                        case "tien_nt0":
                            Decimal result2 = new Decimal(0);
                            Decimal result3 = new Decimal(0);
                            bool result4 = false;
                            bool.TryParse(values[2].ToString(), out result4);
                            Decimal.TryParse(values[0].ToString(), out result2);
                            Decimal.TryParse(values[1].ToString(), out result3);
                            return result2 == new Decimal(0) || result4 ? (object)false : (object)true;
                        case "cp_nt":
                            return (object)false;
                        case "cp":
                            Decimal result5 = new Decimal(0);
                            bool result6 = false;
                            Decimal.TryParse(values[0].ToString(), out result5);
                            bool.TryParse(values[1].ToString(), out result6);
                            return result6 || result5 == new Decimal(0) ? (object)false : (object)true;
                        case "tk_vt":
                            Decimal result7 = new Decimal(0);
                            Decimal.TryParse(values[0].ToString(), out result7);
                            bool result8 = false;
                            bool.TryParse(values[1].ToString(), out result8);
                            return result7 == new Decimal(1) && result8 ? (object)false : (object)true;
                        case "ma_kh2":
                            Decimal result9 = new Decimal(0);
                            Decimal.TryParse(values[0].ToString(), out result9);
                            return result9 == new Decimal(1) ? (object)false : (object)true;
                    }
                    break;
                case "GT":
                    string empty1;
                    switch (strArray[1])
                    {
                        case "ten_kh":
                            empty1 = string.Empty;
                            return string.IsNullOrEmpty(values[0].ToString().Trim()) ? (object)false : (object)true;
                        case "dia_chi":
                            empty1 = string.Empty;
                            string empty2 = string.Empty;
                            string str1 = values[2].ToString().Trim();
                            return string.IsNullOrEmpty(values[0].ToString().Trim()) || string.IsNullOrEmpty(str1) ? (object)false : (object)true;
                        case "ma_so_thue":
                            empty1 = string.Empty;
                            string empty3 = string.Empty;
                            string str2 = values[2].ToString().Trim();
                            return string.IsNullOrEmpty(values[0].ToString().Trim()) || string.IsNullOrEmpty(str2) ? (object)false : (object)true;
                    }
                    break;
            }
            return (object)false;
        }

        public object[] ConvertBack(
          object value,
          Type[] targetTypes,
          object parameter,
          CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
