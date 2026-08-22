using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SasVoucherLib
{
    public class BindingVisibility : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values != null && values[0] != DependencyProperty.UnsetValue && values[0].GetType().Equals(typeof(bool)))
            {
                string[] strArray = parameter.ToString().Trim().Split(';');
                if (strArray.Length == 1)
                {
                    string str = parameter.ToString().Trim();
                    bool flag = (bool)values[0];
                    switch (str)
                    {
                        case "MaNt0":
                        case "LANV":
                            return (object)(Visibility)(flag ? 0 : 2);
                        case "NotMaNt0":
                        case "LANE":
                            return (object)(Visibility)(flag ? 2 : 0);
                    }
                }
                else
                {
                    bool flag = true;
                    for (int index = 0; index < values.Length; ++index)
                    {
                        object obj = values[index];
                        flag = flag && System.Convert.ToBoolean(obj).Equals(System.Convert.ToBoolean(strArray[index + 1]));
                    }
                    switch (strArray[0].ToString().Trim())
                    {
                        case "MaNt0":
                            return (object)(Visibility)(flag ? 0 : 1);
                        case "NotMaNt0":
                            return (object)(Visibility)(flag ? 1 : 0);
                    }
                }
            }
            return (object)Visibility.Visible;
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
