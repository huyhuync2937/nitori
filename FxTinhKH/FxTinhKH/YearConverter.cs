using System;
using System.Globalization;
using System.Windows.Data;

namespace FxTinhKH
{
    [ValueConversion(typeof(int), typeof(int))]
    internal class YearConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (object)((int)value + 1);
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            int result;
            return int.TryParse(value.ToString(), out result) ? (object)(result - 1) : value;
        }
    }
}
