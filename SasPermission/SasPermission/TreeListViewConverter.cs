using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace SasPermission
{
    public class TreeListViewConverter : IValueConverter
    {
        public const double Indentation = 10.0;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return (object)null;
            if (targetType != typeof(double) || !typeof(DependencyObject).IsAssignableFrom(value.GetType()))
                throw new NotSupportedException(string.Format("Cannot convert from <{0}> to <{1}> using <TreeListViewConverter>.", (object)value.GetType(), (object)targetType));
            DependencyObject reference = value as DependencyObject;
            int num = -1;
            for (; reference != null; reference = VisualTreeHelper.GetParent(reference))
            {
                if (typeof(TreeViewItem).IsAssignableFrom(reference.GetType()))
                    ++num;
            }
            return (object)(10.0 * (double)num);
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            throw new NotSupportedException("This method is not supported.");
        }
    }
}
