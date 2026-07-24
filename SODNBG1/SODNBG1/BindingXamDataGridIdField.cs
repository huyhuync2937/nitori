using System;
using System.Globalization;
using System.Windows.Data;

namespace SODNBG1
{
    public class BindingXamDataGridIdField : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (object)(System.Convert.ToInt32(value) + 1);
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            return value;
        }
    }
}
