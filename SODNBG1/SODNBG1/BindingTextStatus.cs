using System;
using System.Globalization;
using System.Windows.Data;

namespace SODNBG1
{
    public class BindingTextStatus : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? (object)"Xử lý" : (object)"Trạng thái";
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
