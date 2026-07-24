using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Poctpxf
{
    public class BindingStatusVoucher : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!string.IsNullOrEmpty(value.ToString().Trim()))
            {
                DataRow[] dataRowArray = StartUpTrans.tbStatus.Select("Ma_post =" + value);
                if (((IEnumerable<DataRow>)dataRowArray).Count<DataRow>() > 0)
                {
                    DataRow row = dataRowArray[0];
                    return (object)StartUpTrans.tbStatus.Rows.IndexOf(row);
                }
            }
            DataRow[] dataRowArray1 = StartUpTrans.tbStatus.Select("Default = 1");
            return dataRowArray1.Length > 0 ? (object)StartUpTrans.tbStatus.Rows.IndexOf(dataRowArray1[0]) : (object)-1;
        }

        public object ConvertBack(
          object value,
          Type targetType,
          object parameter,
          CultureInfo culture)
        {
            return value == null ? StartUpTrans.DmctInfo["ma_post"] : value;
        }
    }
}
