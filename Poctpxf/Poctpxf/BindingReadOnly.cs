using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Poctpxf
{
    public class BindingReadOnly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            if (parameter != null)
            {
                string[] strArray = parameter.ToString().Split(';');
                switch (strArray[0])
                {
                    case "PH":
                        bool result1 = false;
                        bool.TryParse(values[0].ToString(), out result1);
                        switch (strArray[1].ToString())
                        {
                            case "ong_ba":
                                if (result1 && StartUpTrans.M_ong_ba.Equals("1"))
                                    return (object)false;
                                break;
                            case "ngay_lct":
                                if (result1 && StartUpTrans.M_ngay_lct.Equals("1"))
                                    return (object)false;
                                break;
                        }
                        return (object)true;
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
                                        double result3 = 0.0;
                                        double.TryParse(values[0].ToString(), out result3);
                                        if (result3 != 0.0)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "tien_nt":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[3].ToString(), out result2);
                                    if (result2)
                                    {
                                        double result3 = 0.0;
                                        double result4 = 0.0;
                                        bool result5 = false;
                                        bool.TryParse(values[2].ToString(), out result5);
                                        double.TryParse(values[0].ToString(), out result3);
                                        double.TryParse(values[1].ToString(), out result4);
                                        if (result3 == 0.0 || result3 * result4 == 0.0 || result5)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "gia":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[2].ToString(), out result2);
                                    if (result2)
                                    {
                                        Decimal result3 = new Decimal(0);
                                        Decimal result4 = new Decimal(0);
                                        bool result5 = false;
                                        bool.TryParse(values[1].ToString(), out result5);
                                        Decimal.TryParse(values[0].ToString(), out result3);
                                        Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result4);
                                        if (result3 * result4 == new Decimal(0) || result5)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "tien":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue && values[3] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[3].ToString(), out result2);
                                    if (result2)
                                    {
                                        double result3 = 0.0;
                                        double result4 = 0.0;
                                        bool result5 = false;
                                        double.TryParse(values[0].ToString(), out result3);
                                        double.TryParse(values[1].ToString(), out result4);
                                        bool.TryParse(values[2].ToString(), out result5);
                                        if (result5 || result3 * result4 == 0.0)
                                            flag = false;
                                        break;
                                    }
                                    break;
                                }
                                if (values[0] == DependencyProperty.UnsetValue)
                                {
                                    flag = false;
                                    break;
                                }
                                break;
                            case "tk_vt":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[1].ToString(), out result2);
                                    if (result2)
                                    {
                                        double result3 = 0.0;
                                        double.TryParse(values[0].ToString(), out result3);
                                        if (result3 == 1.0)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "ma_kh2_ph":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[1].ToString(), out result2);
                                    if (result2 && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tk_thue_co_cn"].ToString() == "1")
                                        flag = false;
                                    break;
                                }
                                break;
                            case "t_thue_nt_readOnly":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[1].ToString(), out result2);
                                    if (result2)
                                    {
                                        bool result3 = false;
                                        bool.TryParse(values[0].ToString(), out result3);
                                        if (result3)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "t_thue_nt_tabStop":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[1].ToString(), out result2);
                                    if (result2)
                                    {
                                        bool result3 = false;
                                        bool.TryParse(values[0].ToString(), out result3);
                                        if (!result3)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "t_thue_readOnly":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[2].ToString(), out result2);
                                    if (result2)
                                    {
                                        bool result3 = false;
                                        bool.TryParse(values[0].ToString(), out result3);
                                        bool result4 = false;
                                        bool.TryParse(values[1].ToString(), out result4);
                                        if (result3 && result4)
                                            flag = false;
                                    }
                                    break;
                                }
                                break;
                            case "t_thue_tabStop":
                                if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[2] != DependencyProperty.UnsetValue)
                                {
                                    bool result2 = false;
                                    bool.TryParse(values[2].ToString(), out result2);
                                    if (result2)
                                    {
                                        bool result3 = false;
                                        bool.TryParse(values[0].ToString(), out result3);
                                        bool result4 = false;
                                        bool.TryParse(values[1].ToString(), out result4);
                                        flag = result3 && result4;
                                    }
                                    break;
                                }
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
