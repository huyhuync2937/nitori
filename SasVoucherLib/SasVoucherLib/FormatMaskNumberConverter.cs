using SasControls;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace SasVoucherLib
{
    /// <summary>
    /// Lớp converter chuyển định dạng trường tiền, số lượng, giá, tý giả cho control.
    /// </summary>
    public class FormatMaskNumberConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string str1 = "";
            if (StartupBase.SasObj != null)
                str1 = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
            if (!value.Equals(DependencyProperty.UnsetValue) && !string.IsNullOrEmpty(value.ToString()) && (parameter != null && StartupBase.SasObj != null))
            {
                string OptionName = parameter.ToString().Trim();
                string str2 = value.ToString();
                switch (OptionName)
                {
                    case "M_IP_SL":
                        str1 = StartupBase.SasObj.GetOption(OptionName).ToString();
                        break;
                    case "M_IP_TIEN":
                        str1 = StartupBase.SasObj.GetOption(OptionName).ToString();
                        break;
                    case "M_IP_TIEN_NT":
                        string str3 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                        if (str3 != null)
                        {
                            str1 = str2.ToString().Equals(str3) ? StartupBase.SasObj.GetOption("M_IP_TIEN").ToString() : StartupBase.SasObj.GetOption(OptionName).ToString();
                            break;
                        }
                        break;
                    case "M_IP_GIA":
                        str1 = StartupBase.SasObj.GetOption(OptionName).ToString();
                        break;
                    case "M_IP_GIA_NT":
                        string str4 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                        if (str4 != null)
                        {
                            str1 = str2.ToString().Equals(str4) ? StartupBase.SasObj.GetOption("M_IP_GIA").ToString() : StartupBase.SasObj.GetOption(OptionName).ToString();
                            break;
                        }
                        break;
                    case "M_IP_TY_GIA":
                        str1 = StartupBase.SasObj.GetOption(OptionName).ToString();
                        break;
                }
            }
            return (object)str1;
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
