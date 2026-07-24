using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace INCTPND
{
    internal class BindingReadOnly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            if (FrmINCTPND.IsInEditMode == null || !FrmINCTPND.IsInEditMode.Value)
                return (object)true;
            if (parameter != null)
            {
                switch (parameter.ToString())
                {
                    case "gia_nt":
                        Decimal result1 = new Decimal(0);
                        Decimal result2 = new Decimal(1);
                        bool result3 = false;
                        bool result4 = false;
                        Decimal.TryParse(values[0].ToString(), out result1);
                        bool.TryParse(values[1].ToString(), out result3);
                        bool.TryParse(values[2].ToString(), out result4);
                        Decimal.TryParse(values[3].ToString(), out result2);
                        if ((!(result2 == new Decimal(1)) && !(result2 == new Decimal(4)) || !result4) && (result1 != new Decimal(0) || result3))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tien_nt":
                        Decimal result5 = new Decimal(0);
                        Decimal result6 = new Decimal(0);
                        Decimal result7 = new Decimal(1);
                        bool result8 = false;
                        bool result9 = false;
                        Decimal.TryParse(values[0].ToString(), out result5);
                        Decimal.TryParse(values[1].ToString(), out result6);
                        bool.TryParse(values[2].ToString(), out result8);
                        bool.TryParse(values[3].ToString(), out result9);
                        Decimal.TryParse(values[4].ToString(), out result7);
                        if ((!(result7 == new Decimal(1)) && !(result7 == new Decimal(4)) || !result9) && (result5 * result6 == new Decimal(0) || result8))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "gia":
                        Decimal result10 = new Decimal(0);
                        Decimal result11 = new Decimal(0);
                        Decimal result12 = new Decimal(0);
                        Decimal result13 = new Decimal(1);
                        bool result14 = false;
                        bool result15 = false;
                        Decimal.TryParse(values[0].ToString(), out result10);
                        Decimal.TryParse(values[1].ToString(), out result11);
                        bool.TryParse(values[2].ToString(), out result14);
                        bool.TryParse(values[3].ToString(), out result15);
                        Decimal.TryParse(values[4].ToString(), out result13);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result12);
                        if ((!(result13 == new Decimal(1)) && !(result13 == new Decimal(4)) || !result15) && (result10 != new Decimal(0) && result11 == new Decimal(0) || result10 != new Decimal(0) && result11 != new Decimal(0) && result14 || result10 != new Decimal(0) && result11 * result12 == new Decimal(0)))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tien":
                        Decimal result16 = new Decimal(0);
                        Decimal result17 = new Decimal(0);
                        Decimal result18 = new Decimal(1);
                        bool result19 = false;
                        bool result20 = false;
                        Decimal.TryParse(values[0].ToString(), out result16);
                        bool.TryParse(values[1].ToString(), out result19);
                        bool.TryParse(values[2].ToString(), out result20);
                        Decimal.TryParse(values[3].ToString(), out result18);
                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result17);
                        if ((!(result18 == new Decimal(1)) && !(result18 == new Decimal(4)) || !result20) && (result16 * result17 == new Decimal(0) || result19))
                        {
                            flag = false;
                            break;
                        }
                        break;
                    case "tk_vt":
                        Decimal result21 = new Decimal(0);
                        Decimal.TryParse(values[0].ToString(), out result21);
                        bool result22 = false;
                        bool.TryParse(values[1].ToString(), out result22);
                        if (result21 == new Decimal(1) && result22)
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
