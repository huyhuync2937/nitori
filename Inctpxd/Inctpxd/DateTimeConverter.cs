using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Inctpxd
{
    public class DateTimeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != DependencyProperty.UnsetValue && !string.IsNullOrEmpty(value.ToString().Trim()) ? (object)string.Format("{0:dd-MM-yyyy}", value) : (object)"";
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
