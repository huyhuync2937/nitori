using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace QCCNCC
{
    public class BindingReadOnly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            if (parameter != null)
            {
                switch (parameter.ToString())
                {
                    case "gia_nt0":
                        Decimal result1 = new Decimal(0);
                        Decimal.TryParse(values[0].ToString(), out result1);
                        if (result1 != new Decimal(0))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tien_nt0":
                        Decimal result2 = new Decimal(0);
                        Decimal result3 = new Decimal(0);
                        bool result4 = false;
                        bool.TryParse(values[2].ToString(), out result4);
                        Decimal.TryParse(values[0].ToString(), out result2);
                        Decimal.TryParse(values[1].ToString(), out result3);
                        if (result2 == new Decimal(0) || result2 * result3 == new Decimal(0) || result4)
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "gia0":
                        Decimal result5 = new Decimal(0);
                        Decimal result6 = new Decimal(0);
                        Decimal result7 = new Decimal(0);
                        bool result8 = false;
                        bool.TryParse(values[2].ToString(), out result8);
                        Decimal.TryParse(values[0].ToString(), out result5);
                        Decimal.TryParse(values[1].ToString(), out result6);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result7);
                        if (result5 != new Decimal(0) && (result6 * result7 == new Decimal(0) || result8))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tien0":
                        Decimal result9 = new Decimal(0);
                        Decimal result10 = new Decimal(0);
                        bool result11 = false;
                        Decimal.TryParse(values[0].ToString(), out result9);
                        bool.TryParse(values[1].ToString(), out result11);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result10);
                        if (result11 || result9 * result10 == new Decimal(0))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tk_vt":
                        Decimal result12 = new Decimal(0);
                        Decimal.TryParse(values[0].ToString(), out result12);
                        bool result13 = false;
                        bool.TryParse(values[1].ToString(), out result13);
                        if (result12 == new Decimal(1) && result13)
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "cp_nt":
                        flag = false;
                        break;
                    case "cp":
                        Decimal result14 = new Decimal(0);
                        bool result15 = false;
                        Decimal.TryParse(values[0].ToString(), out result14);
                        bool.TryParse(values[1].ToString(), out result15);
                        if (result15 || result14 == new Decimal(0))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "ma_kh2":
                        Decimal result16 = new Decimal(0);
                        Decimal.TryParse(values[0].ToString(), out result16);
                        if (result16 == new Decimal(1))
                        {
                            flag = false;
                            break;
                        }
                        break;
                }
            }
            return (object)flag;
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
