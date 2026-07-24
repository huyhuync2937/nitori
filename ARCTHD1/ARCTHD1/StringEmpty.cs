using System;
using System.Globalization;
using System.Windows.Data;

namespace ARCTHD1
{
    internal class StringEmpty : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (object)string.IsNullOrEmpty(value.ToString().Trim()); 
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
