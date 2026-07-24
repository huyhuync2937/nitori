using SasVoucherLib;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace CACTPC1
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmCACTPC1.IsInEditMode != null && FrmCACTPC1.IsInEditMode.Value && values != null)
            {
                string[] strArray = parameter.ToString().Trim().Split(';');
                switch (strArray[0])
                {
                    case "PH":
                        switch (strArray[1].ToString())
                        {
                            case "ong_ba":
                                bool result1 = false;
                                bool.TryParse(values[0].ToString(), out result1);
                                if (result1 && StartUpTrans.M_ong_ba.Equals("1"))
                                    return (object)false;
                                break;
                            case "ngay_lct":
                                bool result2;
                                bool.TryParse(values[0].ToString(), out result2);
                                if (result2 && StartUpTrans.M_ngay_lct.Equals("1"))
                                    return (object)false;
                                break;
                            case "ma_gd":
                                if (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                                {
                                    if (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0 && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[2].DefaultView[0]["so_ct0"].ToString()))
                                        return (object)true;
                                    if (StartUpTrans.DsTrans.Tables[1].DefaultView[0]["tk_i"] != DBNull.Value && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[1].DefaultView[0]["tk_i"].ToString().Trim()))
                                        return (object)true;
                                }
                                return (object)false;
                            case "sua_tggs":
                                bool result3;
                                bool.TryParse(values[1].ToString(), out result3);
                                if (!result3)
                                    return (object)false;
                                return values[0].ToString() == "3" ? (object)false : (object)true;
                            case "TGGS1":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || (values[2] == DependencyProperty.UnsetValue || values[3] == DependencyProperty.UnsetValue) || values[4] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                if ((bool)values[1])
                                    return (object)true;
                                if ((bool)values[2])
                                    return (object)false;
                                return values[4].ToString() == "2" ? (object)false : (object)true;
                            case "TGGD":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || (values[2] == DependencyProperty.UnsetValue || values[3] == DependencyProperty.UnsetValue) || values[4] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                if ((bool)values[1])
                                    return (object)true;
                                if (values[3].ToString() == "8" || values[3].ToString() == "9")
                                    return (object)false;
                                return values[4].ToString() != "0" ? (object)false : (object)true;
                            case "ETGGS1":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || (values[2] == DependencyProperty.UnsetValue || values[3] == DependencyProperty.UnsetValue) || values[4] == DependencyProperty.UnsetValue)
                                    return (object)false;
                                if (!(bool)values[0])
                                    return (object)false;
                                if ((bool)values[1])
                                    return (object)false;
                                switch (values[3].ToString())
                                {
                                    case "1":
                                    case "2":
                                    case "3":
                                        if (!(bool)values[2] && values[4].ToString() == "0")
                                            return (object)false;
                                        if (values[4].ToString() != "0")
                                            return (object)false;
                                        break;
                                    case "8":
                                    case "9":
                                        if (values[4].ToString() != "0")
                                            return (object)false;
                                        if (!(bool)values[2])
                                            return (object)false;
                                        break;
                                }
                                return (object)true;
                            case "ETGGD":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || (values[2] == DependencyProperty.UnsetValue || values[3] == DependencyProperty.UnsetValue) || values[4] == DependencyProperty.UnsetValue)
                                    return (object)false;
                                if (!(bool)values[0])
                                    return (object)false;
                                if ((bool)values[1])
                                    return (object)false;
                                switch (values[3].ToString())
                                {
                                    case "1":
                                    case "2":
                                    case "3":
                                        return values[4].ToString() == "0" ? (object)false : (object)true;
                                    case "8":
                                    case "9":
                                        return (object)true;
                                    default:
                                        return (object)false;
                                }
                        }
                        return (object)true;
                    case "CT":
                        if (values[0] == DependencyProperty.UnsetValue)
                            return (object)true;
                        string str = values[0].ToString().Trim();
                        if (((IEnumerable<string>)new string[11]
                        {
              "ma_ms",
              "ngay_ct0",
              "so_ct0",
              "ma_kh_t",
              "ten_kh_t",
              "dia_chi_t",
              "mst_t",
              "ten_vt_t",
              "ma_thue_i",
              "tk_thue_i",
              "ghi_chu_t"
                        }).Contains<string>(strArray[1]))
                            return (object)(str == "0");
                        switch (strArray[1])
                        {
                            case "tien_tt":
                            case "tien":
                                if (values[1] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                return !(bool)values[1] ? (object)true : (object)false;
                            case "thue_nt":
                            case "thue":
                                if (values[1] == DependencyProperty.UnsetValue || values[2] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                if (values[2].ToString().Trim() == "0")
                                    return (object)true;
                                return !(bool)values[1] ? (object)true : (object)false;
                            default:
                                return (object)false;
                        }
                    case "CTCHI":
                        switch (strArray[1])
                        {
                            case "tien_tt":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || values[2] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                if (!(bool)values[1])
                                    return (object)true;
                                return values[2].ToString() == "8" ? (object)true : (object)false;
                            case "tien":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                return !(bool)values[1] ? (object)true : (object)false;
                            case "ty_gia_ht2":
                                if (values[0] == DependencyProperty.UnsetValue || values[1] == DependencyProperty.UnsetValue || values[2] == DependencyProperty.UnsetValue)
                                    return (object)true;
                                if (!(bool)values[0])
                                    return (object)true;
                                if ((bool)values[1])
                                    return (object)false;
                                int result4 = 0;
                                int.TryParse(values[2].ToString(), out result4);
                                return result4 == 2 ? (object)false : (object)true;
                            default:
                                return (object)true;
                        }
                    case "CTHD":
                        switch (strArray[1])
                        {
                            case "tien":
                                if (!values[0].Equals(DependencyProperty.UnsetValue) && !values[1].Equals(DependencyProperty.UnsetValue))
                                {
                                    double num1 = System.Convert.ToDouble(values[0]);
                                    double num2 = 1.0;
                                    if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"] != DBNull.Value && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString()))
                                        num2 = System.Convert.ToDouble(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"]);
                                    if (num1 * num2 == 0.0)
                                        return (object)false;
                                }
                                return !values[2].Equals(DependencyProperty.UnsetValue) && (bool)values[2] ? (object)false : (object)true;
                            case "TT_QD":
                                return !values[0].Equals(DependencyProperty.UnsetValue) && !values[1].Equals(DependencyProperty.UnsetValue) ? (object)(values[0].ToString() == values[1].ToString()) : (object)false;
                        }
                        break;
                    case "CTGT":
                        switch (strArray[1].ToString())
                        {
                            case "TTTHUE":
                                bool result5 = false;
                                bool.TryParse(values[0].ToString(), out result5);
                                if (!result5)
                                    return (object)true;
                                return values[1].Equals((object)"1") ? (object)true : (object)false;
                        }
                        break;
                }
            }
            return (object)true;
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
