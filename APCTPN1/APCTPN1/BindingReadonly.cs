using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace APCTPN1
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmAPCTPN1.IsInEditMode == null || !FrmAPCTPN1.IsInEditMode.Value || values == null)
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
                        case "gia_nt":
                            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                            {
                                bool result2 = false;
                                bool.TryParse(values[1].ToString(), out result2);
                                if (result2)
                                {
                                    Decimal result3 = new Decimal(0);
                                    Decimal.TryParse(values[0].ToString(), out result3);
                                    if (result3 != new Decimal(0))
                                        return (object)false;
                                }
                            }
                            return (object)true;
                        case "gia":
                            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue)
                            {
                                bool result2 = false;
                                bool.TryParse(values[3].ToString(), out result2);
                                if (result2)
                                {
                                    Decimal result3 = new Decimal(0);
                                    Decimal result4 = new Decimal(0);
                                    Decimal result5 = new Decimal(0);
                                    bool result6 = false;
                                    Decimal.TryParse(values[0].ToString(), out result3);
                                    Decimal.TryParse(values[1].ToString(), out result4);
                                    bool.TryParse(values[2].ToString(), out result6);
                                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result5);
                                    if (result3 != new Decimal(0) && (result4 * result5 == new Decimal(0) || result6))
                                        return (object)false;
                                }
                            }
                            return (object)true;
                        case "tien":
                            bool flag = true;
                            if (!values[0].Equals(DependencyProperty.UnsetValue) && !values[1].Equals(DependencyProperty.UnsetValue))
                            {
                                Decimal num1 = System.Convert.ToDecimal(values[0]);
                                Decimal num2 = new Decimal(1);
                                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"] != DBNull.Value)
                                    num2 = System.Convert.ToDecimal(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"]);
                                if (num1 * num2 == new Decimal(0))
                                    flag = false;
                            }
                            if (!values[2].Equals(DependencyProperty.UnsetValue) && (bool)values[2])
                                flag = false;
                            return (object)flag;
                    }
                    break;
                case "GT":
                    string empty1;
                    switch (strArray[1])
                    {
                        case "ma_kh2":
                            if (values[0] != null)
                                return (object)!values[0].ToString().Trim().Equals("1");
                            break;
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
