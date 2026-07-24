using System;
using System.Globalization;
using System.Windows.Data;

namespace INCTPND
{
    public class CheckBoxConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int result = 0;
            int.TryParse(value.ToString(), out result);
            return (object)(result == 1);
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            return (object)(bool.Parse(value.ToString()) ? 1 : 0);
        }
    }
}
