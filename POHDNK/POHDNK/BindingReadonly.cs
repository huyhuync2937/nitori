using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace POHDNK
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmPoctpna.IsInEditMode != null && FrmPoctpna.IsInEditMode.Value && values != null)
            {
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
                            case "gia_nt0":
                                Decimal result2 = new Decimal(0);
                                Decimal.TryParse(values[0].ToString(), out result2);
                                return result2 != new Decimal(0) ? (object)false : (object)true;
                            case "tien_nt0":
                                Decimal result3 = new Decimal(0);
                                Decimal result4 = new Decimal(0);
                                bool result5 = false;
                                bool.TryParse(values[2].ToString(), out result5);
                                Decimal.TryParse(values[0].ToString(), out result3);
                                Decimal.TryParse(values[1].ToString(), out result4);
                                return result3 == new Decimal(0) || result3 * result4 == new Decimal(0) || result5 ? (object)false : (object)true;
                            case "db":
                            case "nk":
                                Decimal result6 = new Decimal(0);
                                Decimal result7 = new Decimal(0);
                                bool result8 = false;
                                bool.TryParse(values[1].ToString(), out result8);
                                Decimal.TryParse(values[0].ToString(), out result6);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result7);
                                return result6 * result7 == new Decimal(0) || result8 ? (object)false : (object)true;
                            case "gia0":
                                Decimal result9 = new Decimal(0);
                                Decimal result10 = new Decimal(0);
                                Decimal result11 = new Decimal(0);
                                bool result12 = false;
                                bool.TryParse(values[2].ToString(), out result12);
                                Decimal.TryParse(values[0].ToString(), out result9);
                                Decimal.TryParse(values[1].ToString(), out result10);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result11);
                                return result9 != new Decimal(0) && (result10 * result11 == new Decimal(0) || result12) ? (object)false : (object)true;
                            case "tien0":
                                Decimal result13 = new Decimal(0);
                                Decimal result14 = new Decimal(0);
                                bool result15 = false;
                                Decimal.TryParse(values[0].ToString(), out result13);
                                bool.TryParse(values[1].ToString(), out result15);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result14);
                                return result15 || result13 * result14 == new Decimal(0) ? (object)false : (object)true;
                            case "thue":
                                Decimal result16 = new Decimal(0);
                                Decimal result17 = new Decimal(0);
                                bool result18 = false;
                                Decimal.TryParse(values[0].ToString(), out result16);
                                Decimal.TryParse(values[1].ToString(), out result17);
                                bool.TryParse(values[2].ToString(), out result18);
                                return result18 || result16 * (result17 / new Decimal(100)) == new Decimal(0) ? (object)false : (object)true;
                        }
                        break;
                    case "CP":
                        switch (strArray[1])
                        {
                            case "tien":
                                if (!values[0].Equals(DependencyProperty.UnsetValue) && !values[1].Equals(DependencyProperty.UnsetValue))
                                {
                                    double num1 = System.Convert.ToDouble(values[0]);
                                    double num2 = 1.0;
                                    if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"] != DBNull.Value)
                                        num2 = System.Convert.ToDouble(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"]);
                                    if (num1 * num2 == 0.0)
                                        return (object)false;
                                }
                                return !values[2].Equals(DependencyProperty.UnsetValue) && (bool)values[2] ? (object)false : (object)true;
                        }
                        break;
                }
            }
            return (object)true;
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
