using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Invt
{
    public class IsEnabledTon_kho : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = false;
            if (parameter != null)
            {
                switch (parameter.ToString())
                {
                    case "vt_ton_kho":
                    case "sua_tk_vt":
                        if (values[0] != DependencyProperty.UnsetValue && values[0].ToString().Trim() != string.Empty)
                        {
                            flag = true;
                            break;
                        }
                        break;
                    case "gia_ton":
                    case "tk_dtnb":
                    case "tk_cl_vt":
                    case "tk_ck":
                    case "tk_nvl":
                    case "tk_spdd":
                    case "sl":
                        if (values[0] != DependencyProperty.UnsetValue)
                        {
                            string str = values[0].ToString().Trim();
                            int result = 0;
                            int.TryParse(values[1].ToString(), out result);
                            if (str != string.Empty && result != 0)
                                flag = true;
                            break;
                        }
                        break;
                    case "tk_vt":
                        if (values[0] != DependencyProperty.UnsetValue && values[0].ToString().Trim() == "0")
                        {
                            flag = true;
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
