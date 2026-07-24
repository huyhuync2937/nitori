using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Poctpna
{
    public class BindingHeightCP : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            double result1 = 0.0;
            double result2 = 0.0;
            double.TryParse(values[2].ToString(), out result2);
            switch ((Visibility)values[1])
            {
                case Visibility.Visible:
                    if (double.TryParse(values[0].ToString(), out result1))
                    {
                        result2 -= result1;
                        break;
                    }
                    break;
            }
            return (object)result2;
        }

        public object[] ConvertBack(
          object value,
          Type[] targetTypes,
          object parameter,
          CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
