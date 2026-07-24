using System;
using System.Globalization;
using System.Windows.Data;

namespace Poctpna
{
    public class IdRecordConverterXamDataGrid : IValueConverter
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
            return value;
        }
    }
}
