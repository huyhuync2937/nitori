using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace CACTPC1
{
    public class BindingVisibility : IMultiValueConverter
    {
        private string[] listBtPb = new string[3]
        {
      "2",
      "3",
      "9"
        };
        private string[] listGdCt = new string[8]
        {
      "2",
      "3",
      "4",
      "5",
      "6",
      "7",
      "9",
      ""
        };
        private string[] listGdCp = new string[1] { "8" };
        private string[] listGdHd = new string[1] { "1" };

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            Visibility visibility = Visibility.Visible;
            try
            {
                string[] strArray = parameter.ToString().Trim().Split(';');
                if (values != null && values[0] != DependencyProperty.UnsetValue)
                {
                    switch (strArray[0])
                    {
                        case "MaNt0":
                        case "NotMaNt0":
                            return new SasVoucherLib.BindingVisibility().Convert(values, targetType, parameter, culture);
                        case "FieldCT":
                            switch (strArray[1])
                            {
                                case "ma_gd":
                                    int result = 2;
                                    int.TryParse(values[0].ToString(), out result);
                                    if (result == 2 || result == 3)
                                    {
                                        visibility = Visibility.Visible;
                                        break;
                                    }
                                    if (result == 9)
                                    {
                                        visibility = Visibility.Collapsed;
                                        break;
                                    }
                                    break;
                                case "ma_kh":
                                    visibility = !(values[0].ToString() == "3") ? Visibility.Collapsed : Visibility.Visible;
                                    break;
                                case "ty_gia_ht2":
                                    return values[0].ToString() == "2" ? (object)Visibility.Visible : (object)Visibility.Collapsed;
                                case "thue":
                                case "tien":
                                    values[0].ToString();
                                    bool flag1 = true;
                                    if (values.Length > 1 && values[1] != DependencyProperty.UnsetValue)
                                        flag1 = (bool)values[1];
                                    return flag1 ? (object)Visibility.Collapsed : (object)Visibility.Visible;
                            }
                            if (strArray.Length > 2)
                            {
                                switch (strArray[2])
                                {
                                    case "ten_kh_i":
                                        visibility = !(values[0].ToString() == "3") || !StartUpTrans.M_LAN.Equals("V") ? Visibility.Collapsed : Visibility.Visible;
                                        break;
                                    case "ten_kh2_i":
                                        visibility = !(values[0].ToString() == "3") || !StartUpTrans.M_LAN.Equals("E") ? Visibility.Collapsed : Visibility.Visible;
                                        break;
                                }
                                break;
                            }
                            break;
                        case "FieldCTCHI":
                            string str1 = values[0].ToString();
                            bool flag2 = true;
                            if (values.Length > 1 && values[1] != DependencyProperty.UnsetValue)
                                flag2 = (bool)values[1];
                            switch (strArray[1])
                            {
                                case "ty_gia_ht2":
                                    return str1 == "3" || str1 == "9" || flag2 ? (object)Visibility.Collapsed : (object)Visibility.Visible;
                                case "tien":
                                    return str1 == "3" || flag2 ? (object)Visibility.Collapsed : (object)Visibility.Visible;
                            }
                            break;
                        case "KindCT":
                            string str2 = values[0].ToString();
                            string str3 = strArray[1].ToString();
                            if (((IEnumerable<string>)this.listGdCt).Contains<string>(str2) && str3 == "CHI")
                                return (object)Visibility.Visible;
                            if (((IEnumerable<string>)this.listGdHd).Contains<string>(str2) && str3 == "HD")
                                return (object)Visibility.Visible;
                            return ((IEnumerable<string>)this.listGdCp).Contains<string>(str2) && str3 == "CP" ? (object)Visibility.Visible : (object)Visibility.Collapsed;
                        case "KindBT":
                            string str4 = values[0].ToString();
                            string str5 = strArray[1].ToString();
                            return ((IEnumerable<string>)this.listBtPb).Contains<string>(str4) && str5 == "PB" ? (object)Visibility.Visible : (object)Visibility.Collapsed;
                        case "KindTG":
                            string str6 = values[0].ToString();
                            string str7 = values[1].ToString();
                            object obj = (object)values[2].ToString();
                            if (strArray[1].ToString() != "GS" || str7 == StartUpTrans.M_ma_nt0)
                                return (object)Visibility.Collapsed;
                            if (((IEnumerable<string>)StartUp.M_Gd_2Tg_List).Contains<string>(str6))
                                return (object)Visibility.Visible;
                            return obj != DependencyProperty.UnsetValue && obj.ToString() != "0" ? (object)Visibility.Visible : (object)Visibility.Collapsed;
                        case "KindText":
                            string str8 = values[0].ToString();
                            string str9 = values[1].ToString();
                            string str10 = strArray[1].ToString();
                            return ((IEnumerable<string>)StartUp.M_Gd_2Tg_List).Contains<string>(str8) && str10 == "GS" && str9 != StartUpTrans.M_ma_nt0 ? (object)"TGGD" : (object)"TGGS1";
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
            return (object)visibility;
        }

        public object[] ConvertBack(
          object value,
          Type[] targetTypes,
          object parameter,
          CultureInfo culture)
        {
            string[] strArray = parameter.ToString().Trim().Split(';');
            switch (strArray[0])
            {
                case "FieldCT":
                    switch (strArray[1])
                    {
                        case "ma_kh":
                            return (object[])new string[1]
                            {
                FrmCACTPC1.Ma_GD_Value.Text
                            };
                    }
                    break;
            }
            throw new NotImplementedException();
        }
    }
}
