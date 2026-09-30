using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace INSD3
{
    public class GroupValidConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            string[] strArray = new string[3] { "1", "2", "3" };
            string str1 = "";
            if (values[0] != DependencyProperty.UnsetValue && values[1] != DependencyProperty.UnsetValue && values[0] != null && values[1] != null)
            {
                foreach (string str2 in ((IEnumerable<string>)strArray).Except<string>((IEnumerable<string>)new string[2]
                {
          values[0].ToString(),
          values[1].ToString()
                }))
                    str1 = str1 + "," + str2;
            }
            return (object)("0" + str1);
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
