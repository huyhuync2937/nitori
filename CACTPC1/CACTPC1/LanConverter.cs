using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace CACTPC1
{
    public class LanConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            Visibility visibility = Visibility.Visible;
            if (values != null && (values[0] != DBNull.Value && values[0] != DependencyProperty.UnsetValue))
            {
                switch (parameter.ToString().Trim())
                {
                    case "LANV":
                        visibility = StartUpTrans.M_LAN.Equals("V") ? Visibility.Visible : Visibility.Collapsed;
                        break;
                    case "LANE":
                        visibility = StartUpTrans.M_LAN.Equals("E") ? Visibility.Visible : Visibility.Collapsed;
                        break;
                }
            }
            return (object)visibility;
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
