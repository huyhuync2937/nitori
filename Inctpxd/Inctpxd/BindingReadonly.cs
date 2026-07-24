using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Inctpxd
{
    internal class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            try
            {
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
                            case "so_luong_sp":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result = false;
                                    bool.TryParse(values[1].ToString(), out result);
                                    string str = values[0].ToString().Trim();
                                    if (result && !string.IsNullOrEmpty(str))
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
                        switch (strArray[1])
                        {
                            case "ma_vt":
                            case "ma_sp":
                            case "ma_bpht":
                            case "so_lsx":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[1].ToString(), out result1);
                                    Decimal result2 = new Decimal(-1);
                                    Decimal.TryParse(values[0].ToString(), out result2);
                                    if (result1 && result2 <= new Decimal(0))
                                        flag = false;
                                    break;
                                }
                                break;
                            case "so_luong":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[2].ToString(), out result1);
                                    Decimal result2 = new Decimal(-1);
                                    Decimal.TryParse(values[1].ToString(), out result2);
                                    if (result1 && result2 <= new Decimal(0))
                                        flag = false;
                                    break;
                                }
                                break;
                            case "gia_nt":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue) && values[4] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[3].ToString(), out result1);
                                    if (result1)
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal.TryParse(values[0].ToString(), out result2);
                                        Decimal result3 = new Decimal(0);
                                        Decimal.TryParse(values[1].ToString(), out result3);
                                        bool result4 = false;
                                        bool.TryParse(values[2].ToString(), out result4);
                                        if (result2 != new Decimal(0) && (result4 && (result3 == new Decimal(1) || result3 == new Decimal(4)) || result3 == new Decimal(2)) && values[4].ToString() == "1")
                                            flag = false;
                                        if (result2 != new Decimal(0) && result3 == new Decimal(3) && values[4].ToString() == "1" && StartUp.M_SL0_NTXT_CK == "1")
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "tien_nt":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue) && (values[4] != DependencyProperty.UnsetValue && values[5] != DependencyProperty.UnsetValue) && values[6] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[4].ToString(), out result1);
                                    if (result1)
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal result3 = new Decimal(0);
                                        bool result4 = false;
                                        bool.TryParse(values[3].ToString(), out result4);
                                        bool result5 = false;
                                        bool.TryParse(values[2].ToString(), out result5);
                                        Decimal.TryParse(values[0].ToString(), out result2);
                                        Decimal.TryParse(values[1].ToString(), out result3);
                                        Decimal result6 = new Decimal(0);
                                        Decimal.TryParse(values[6].ToString(), out result6);
                                        if ((result2 == new Decimal(0) || result2 * result3 == new Decimal(0) || result4) && (result5 && (result6 == new Decimal(1) || result6 == new Decimal(4)) || result6 == new Decimal(2)) && values[5].ToString() == "1")
                                            flag = false;
                                        if ((result2 == new Decimal(0) || result2 * result3 == new Decimal(0) || result4) && (result6 == new Decimal(3) && values[5].ToString() == "1") && StartUp.M_SL0_NTXT_CK == "1")
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "gia":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue) && (values[4] != DependencyProperty.UnsetValue && values[5] != DependencyProperty.UnsetValue && values[6] != DependencyProperty.UnsetValue) && values[7] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[4].ToString(), out result1);
                                    if (result1)
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal result3 = new Decimal(0);
                                        Decimal result4 = new Decimal(0);
                                        bool result5 = false;
                                        bool result6 = false;
                                        Decimal.TryParse(values[0].ToString(), out result2);
                                        Decimal.TryParse(values[1].ToString(), out result3);
                                        bool.TryParse(values[2].ToString(), out result6);
                                        bool.TryParse(values[3].ToString(), out result5);
                                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result4);
                                        Decimal result7 = new Decimal(0);
                                        Decimal.TryParse(values[7].ToString(), out result7);
                                        if (result2 != new Decimal(0) && values[6].ToString() == "1" && (result5 && (result7 == new Decimal(1) || result7 == new Decimal(4)) || result7 == new Decimal(2)) && (result6 || result3 * result4 == new Decimal(0)))
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "tien":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue) && (values[4] != DependencyProperty.UnsetValue && values[5] != DependencyProperty.UnsetValue) && values[6] != DependencyProperty.UnsetValue)
                                {
                                    bool result1 = false;
                                    bool.TryParse(values[3].ToString(), out result1);
                                    if (result1)
                                    {
                                        Decimal result2 = new Decimal(0);
                                        Decimal result3 = new Decimal(0);
                                        bool result4 = false;
                                        bool result5 = false;
                                        Decimal.TryParse(values[0].ToString(), out result2);
                                        bool.TryParse(values[1].ToString(), out result5);
                                        bool.TryParse(values[2].ToString(), out result4);
                                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result3);
                                        Decimal result6 = new Decimal(0);
                                        Decimal.TryParse(values[6].ToString(), out result6);
                                        if ((result5 || result2 * result3 == new Decimal(0)) && (result4 && (result6 == new Decimal(1) || result6 == new Decimal(4)) || result6 == new Decimal(2)) && values[5].ToString() == "1")
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "tk_vt":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue)
                                {
                                    bool result = false;
                                    bool.TryParse(values[2].ToString(), out result);
                                    if (result)
                                    {
                                        string tk = values[1].ToString().Trim();
                                        if (values[3].ToString().Trim() == "1" && (values[0].ToString() == "1" || tk == "" || StartUp.IsTkMe(tk)))
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                        }
                        return (object)flag;
                }
            }
            catch (Exception ex)
            {
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
