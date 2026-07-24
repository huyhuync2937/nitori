using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace PODMHDM
{
    public class BindingStatusVoucher : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !string.IsNullOrEmpty(value.ToString().Trim()) ? (object)StartUpTrans.tbStatus.Rows.IndexOf(StartUpTrans.tbStatus.Select("Ma_post =" + value)[0]) : (object)"";
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            if (value != null)
            {
                int result = -1;
                int.TryParse(value.ToString(), out result);
                if (result != -1)
                    return StartUpTrans.tbStatus.Rows[result]["ma_post"];
            }
            return StartUpTrans.DmctInfo["ma_post"];
        }
    }
}
