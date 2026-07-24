using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ARCTHD1
{
    internal class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            if (FrmArcthd1.IsInEditMode != null && !FrmArcthd1.IsInEditMode.Value)
                return (object)true;
            string[] strArray = parameter.ToString().Split(';');
            if (strArray.Length < 2)
                return (object)true;
            switch (strArray[0])
            {
                case "PH":
                    switch (strArray[1])
                    {
                        case "ong_ba":
                            if (values[0] != DependencyProperty.UnsetValue)
                            {
                                bool result = false;
                                bool.TryParse(values[0].ToString(), out result);
                                if (result && StartUpTrans.M_ong_ba.Equals("1"))
                                    flag = false;
                                break;
                            }
                            break;
                        case "ngay_lct":
                            if (values[0] != DependencyProperty.UnsetValue)
                            {
                                bool result = false;
                                bool.TryParse(values[0].ToString(), out result);
                                if (result && StartUpTrans.M_ngay_lct.Equals("1"))
                                    flag = false;
                                break;
                            }
                            break;
                    }
                    break;
                case "CT":
                    if (FrmArcthd1.IsInEditMode != null && !FrmArcthd1.IsInEditMode.Value)
                        return (object)true;
                    switch (strArray[1])
                    {
                        case "ma_kh_i":
                        case "ngay_ct0":
                        case "so_ct0":
                        case "so_seri0":
                        case "han_tt":
                            int result1 = 0;
                            int.TryParse(values[0].ToString(), out result1);
                            if (result1 == 1)
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "gia_nt2":
                            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                            {
                                bool result2 = false;
                                bool.TryParse(values[1].ToString(), out result2);
                                if (result2)
                                {
                                    Decimal result3 = new Decimal(0);
                                    Decimal.TryParse(values[0].ToString(), out result3);
                                    if (result3 != new Decimal(0))
                                        flag = false;
                                }
                                break;
                            }
                            break;
                        case "tien_nt2":
                            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue)
                            {
                                bool result2 = false;
                                bool.TryParse(values[3].ToString(), out result2);
                                if (result2)
                                {
                                    Decimal result3 = new Decimal(0);
                                    Decimal result4 = new Decimal(0);
                                    bool result5 = false;
                                    bool.TryParse(values[2].ToString(), out result5);
                                    Decimal.TryParse(values[0].ToString(), out result3);
                                    Decimal.TryParse(values[1].ToString(), out result4);
                                    if (result3 == new Decimal(0) || result3 * result4 == new Decimal(0) || result5)
                                        flag = false;
                                }
                                break;
                            }
                            break;
                        case "gia2":
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
                                        flag = false;
                                }
                                break;
                            }
                            break;
                        case "tien2":
                            Decimal result7 = new Decimal(0);
                            Decimal result8 = new Decimal(0);
                            bool result9 = false;
                            bool.TryParse(values[1].ToString(), out result9);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result8);
                            Decimal.TryParse(values[0].ToString(), out result7);
                            if (result7 * result8 == new Decimal(0) || result9)
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "tk_ck":
                            if (values[0] != null && values[0] != DependencyProperty.UnsetValue)
                                return (object)(double.Parse(values[0].ToString()) == 0.0);
                            break;
                        case "ck":
                            Decimal result10 = new Decimal(0);
                            bool result11 = false;
                            bool.TryParse(values[1].ToString(), out result11);
                            Decimal.TryParse(values[0].ToString(), out result10);
                            if (result10 == new Decimal(0) || result11)
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "ck_nt":
                            Decimal result12 = new Decimal(0);
                            Decimal.TryParse(values[0].ToString(), out result12);
                            if (result12 != new Decimal(0))
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "thue_nt":
                            Decimal result13 = new Decimal(0);
                            bool result14 = false;
                            bool.TryParse(values[1].ToString(), out result14);
                            Decimal.TryParse(values[0].ToString(), out result13);
                            if (result13 == new Decimal(0) || result14)
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "thue":
                            Decimal result15 = new Decimal(0);
                            bool result16 = false;
                            bool result17 = false;
                            bool.TryParse(values[1].ToString(), out result17);
                            bool.TryParse(values[2].ToString(), out result16);
                            Decimal.TryParse(values[0].ToString(), out result15);
                            if (result15 == new Decimal(0) || result16 && result16)
                            {
                                flag = false;
                                break;
                            }
                            break;
                        case "tk_thue_i":
                            bool result18 = false;
                            int result19 = 1;
                            bool.TryParse(values[0].ToString(), out result18);
                            int.TryParse(values[2].ToString(), out result19);
                            if (result18 || result19 == 0)
                            {
                                flag = false;
                                break;
                            }
                            break;
                    }
                    break;
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
