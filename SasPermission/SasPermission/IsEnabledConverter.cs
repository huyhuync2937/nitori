using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SasPermission
{
    public class IsEnabledConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = false;
            if (parameter != null)
            {
                switch (parameter.ToString())
                {
                    case "Write":
                    case "Edit":
                    case "Delete":
                    case "Read":
                        if (values[0] != DependencyProperty.UnsetValue && values[0] != null)
                        {
                            flag = (bool)values[0];
                            break;
                        }
                        break;
                    case "Print":
                        if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && (values[2] != DependencyProperty.UnsetValue && values[0] != null) && values[1] != null && (values[1].ToString() != "" || values[2] != null))
                        {
                            flag = (bool)values[0];
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
