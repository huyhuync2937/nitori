using SasControls;
using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace Socthda
{
  internal class NotMaNt0Converter : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
      return StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString() == StartupBase.SasObj.GetOption("M_MA_NT0").ToString() ? (object) "Hidden" : (object) "Visible";
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
