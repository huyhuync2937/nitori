using System;
using System.Globalization;
using System.Windows.Data;

namespace SasPermission
{
    public class IsEnableCheckBoxConverter : IValueConverter
    {
        object IValueConverter.Convert(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            int result = 0;
            int.TryParse(value.ToString(), out result);
            return (object)(result == 1);
        }

        object IValueConverter.ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
