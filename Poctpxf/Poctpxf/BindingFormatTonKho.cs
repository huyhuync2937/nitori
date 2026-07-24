using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace Poctpxf
{
    public class BindingFormatTonKho : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return new FormatMaskNumberConverter().Convert(value, targetType, parameter, culture);
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
