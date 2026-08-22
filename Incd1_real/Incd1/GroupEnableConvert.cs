using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Incd1
{
    public class GroupEnableConvert : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = true;
            foreach (object obj in values)
            {
                if (obj != null && obj != DependencyProperty.UnsetValue && obj.ToString() == "0")
                    flag = false;
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
