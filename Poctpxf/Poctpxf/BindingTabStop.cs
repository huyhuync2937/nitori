using System;
using System.Globalization;
using System.Windows.Data;

namespace Poctpxf
{
    public class BindingTabStop : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value != null ? (object)value.ToString().Trim().Equals("1") : (object)true;
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
