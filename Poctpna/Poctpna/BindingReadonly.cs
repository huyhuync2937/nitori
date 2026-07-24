using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Poctpna
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmPoctpna.IsInEditMode == null || !FrmPoctpna.IsInEditMode.Value)
                return (object)true;
            if (values == null)
                return (object)false;
            string[] strArray = parameter.ToString().Trim().Split(';');
            bool flag = true;
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
                        case "gia0":
                            Decimal result6 = new Decimal(0);
                            Decimal result7 = new Decimal(0);
                            Decimal result8 = new Decimal(0);
                            bool result9 = false;
                            bool.TryParse(values[2].ToString(), out result9);
                            Decimal.TryParse(values[0].ToString(), out result6);
                            Decimal.TryParse(values[1].ToString(), out result7);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result8);
                            return result6 != new Decimal(0) && (result7 * result8 == new Decimal(0) || result9) ? (object)false : (object)true;
                        case "tien0":
                            Decimal result10 = new Decimal(0);
                            Decimal result11 = new Decimal(0);
                            bool result12 = false;
                            Decimal.TryParse(values[0].ToString(), out result10);
                            bool.TryParse(values[1].ToString(), out result12);
                            Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result11);
                            return result12 || result10 * result11 == new Decimal(0) ? (object)false : (object)true;
                        case "tk_vt":
                            Decimal result13 = new Decimal(0);
                            Decimal.TryParse(values[0].ToString(), out result13);
                            bool result14 = false;
                            bool.TryParse(values[1].ToString(), out result14);
                            return result13 == new Decimal(1) && result14 ? (object)false : (object)true;
                        case "cp_nt":
                            return (object)false;
                        case "cp":
                            Decimal result15 = new Decimal(0);
                            bool result16 = false;
                            Decimal.TryParse(values[0].ToString(), out result15);
                            bool.TryParse(values[1].ToString(), out result16);
                            return result16 || result15 == new Decimal(0) ? (object)false : (object)true;
                        case "ma_kh2":
                            Decimal result17 = new Decimal(0);
                            Decimal.TryParse(values[0].ToString(), out result17);
                            return result17 == new Decimal(1) ? (object)false : (object)true;
                        case "ck_nt":
                            if (values[1] != DependencyProperty.UnsetValue)
                            {
                                bool result = false;
                                bool.TryParse(values[1].ToString(), out result);
                                if (result)
                                    flag = false;
                                break;
                            }
                            break;

                        case "ck":
                            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue) && values[4] != DependencyProperty.UnsetValue)
                            {
                                bool result111 = false;
                                bool.TryParse(values[2].ToString(), out result111);
                                if (result111)
                                {
                                    Decimal result21 = new Decimal(0);
                                    Decimal result31 = new Decimal(0);
                                    bool result41 = false;
                                    Decimal result51 = new Decimal(0);
                                    Decimal.TryParse(values[0].ToString(), out result21);
                                    bool.TryParse(values[1].ToString(), out result41);
                                    Decimal.TryParse(values[4].ToString(), out result51);
                                    Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result31);
                                    if ((result41 || result21 * result31 == new Decimal(0)) && result51 != new Decimal(0))
                                        flag = false;
                                }
                                break;
                            }
                            break;
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
