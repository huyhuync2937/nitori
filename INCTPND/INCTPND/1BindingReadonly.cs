// Decompiled with JetBrains decompiler
// Type: INCTPND.BindingReadonly
// Assembly: INCTPND, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 266C292A-FBFF-416E-8263-964721EC8538
// Assembly location: E:\PM_KETOAN\FA11R08\Program\INCTPND.exe

using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace INCTPND
{
  public class BindingReadonly : IValueConverter
  {
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
      bool result = false;
      bool.TryParse(value.ToString(), out result);
      string[] strArray = parameter.ToString().Split(';');
      if (strArray.Length < 2)
        return (object) true;
      switch (strArray[0])
      {
        case "PH":
          switch (strArray[1])
          {
            case "ong_ba":
              if (result && StartUpTrans.M_ong_ba.Equals("1"))
                return (object) false;
              break;
            case "ngay_lct":
              if (result && StartUpTrans.M_ngay_lct.Equals("1"))
                return (object) false;
              break;
            case "ma_bp":
              if (result && StartUp.M_bp_bh == 1)
                return (object) false;
              break;
          }
          break;
      }
      return (object) true;
    }

    public object ConvertBack(
      object value,
      Type targetType,
      object parameter,
      CultureInfo culture)
    {
      return (object) null;
    }
  }
}
