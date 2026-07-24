using SasVoucherLib;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Socthda
{
    internal class BindingFilter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            string filter = "1=1";
            string[] strArray = parameter.ToString().Split(';');
            if (strArray.Length < 2)
                return filter;
            switch (strArray[0])
            {
                case "CT":
                    switch (strArray[1])
                    {
                        case "dvt1":
                            if (values[0] != DependencyProperty.UnsetValue)
                            {
                                if (string.IsNullOrEmpty(values[0].ToString()))
                                    filter = "1=0";
                                else
                                    filter += " AND ma_vt <> '' and ltrim(rtrim(ma_vt)) LIKE '" + values[0].ToString().Trim() + "'";
                            }
                            break;
                    }
                    break;
            }
            return filter;
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
