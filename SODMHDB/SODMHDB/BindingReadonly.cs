using SasVoucherLib;
using System;
using System.Globalization;
using System.Windows.Data;

namespace SODMHDB
{
    public class BindingReadonly : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (FrmPoctpna.IsInEditMode != null && FrmPoctpna.IsInEditMode.Value && values != null)
            {
                string[] strArray = parameter.ToString().Trim().Split(';');
                switch (strArray[0].ToUpper())
                {
                    case "PH":
                        bool result1 = false;
                        bool.TryParse(values[0].ToString(), out result1);
                        switch (strArray[1])
                        {
                            case "ong_ba":
                                if (result1 && StartUpTrans.M_ong_ba.Equals("1"))
                                    return (object)false;
                                break;
                            case "ngay_lct":
                                if (result1 && StartUpTrans.M_ngay_lct.Equals("1"))
                                    return (object)false;
                                break;
                        }
                        break;
                    case "CT":
                        switch (strArray[1])
                        {
                            case "gia_nt2":
                                Decimal result2 = new Decimal(0);
                                Decimal.TryParse(values[0].ToString(), out result2);
                                if (result2 != new Decimal(0))
                                    return (object)false;
                                break;
                            case "tien_nt2":
                                Decimal result3 = new Decimal(0);
                                Decimal result4 = new Decimal(0);
                                bool result5 = false;
                                bool.TryParse(values[2].ToString(), out result5);
                                Decimal.TryParse(values[0].ToString(), out result3);
                                Decimal.TryParse(values[1].ToString(), out result4);
                                if (result3 == new Decimal(0) || result3 * result4 == new Decimal(0) || result5)
                                    return (object)false;
                                break;
                            case "tl_ck":
                                Decimal result6 = new Decimal(0);
                                Decimal.TryParse(values[0].ToString(), out result6);
                                if (result6 != new Decimal(0))
                                    return (object)false;
                                break;
                            case "ck_nt":
                                Decimal result7 = new Decimal(0);
                                Decimal result8 = new Decimal(0);
                                bool result9 = false;
                                bool.TryParse(values[2].ToString(), out result9);
                                Decimal.TryParse(values[0].ToString(), out result7);
                                Decimal.TryParse(values[1].ToString(), out result8);
                                if (result7 * result8 != new Decimal(0) || result9)
                                    return (object)false;
                                break;
                            case "thue_nt":
                                return (object)false;
                            case "gia_nt":
                                Decimal result10 = new Decimal(0);
                                Decimal.TryParse(values[0].ToString(), out result10);
                                if (result10 != new Decimal(0))
                                    return (object)false;
                                break;
                            case "tien_nt":
                                Decimal result11 = new Decimal(0);
                                Decimal result12 = new Decimal(0);
                                bool result13 = false;
                                bool.TryParse(values[2].ToString(), out result13);
                                Decimal.TryParse(values[0].ToString(), out result11);
                                Decimal.TryParse(values[1].ToString(), out result12);
                                if (result11 == new Decimal(0) || result11 * result12 == new Decimal(0) || result13)
                                    return (object)false;
                                break;
                            case "gia2":
                                Decimal result14 = new Decimal(0);
                                Decimal result15 = new Decimal(0);
                                bool result16 = false;
                                Decimal.TryParse(values[0].ToString(), out result14);
                                bool.TryParse(values[1].ToString(), out result16);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result15);
                                if (result16 || result14 * result15 == new Decimal(0))
                                    return (object)false;
                                break;
                            case "tien2":
                                Decimal result17 = new Decimal(0);
                                Decimal result18 = new Decimal(0);
                                bool result19 = false;
                                Decimal.TryParse(values[0].ToString(), out result17);
                                bool.TryParse(values[1].ToString(), out result19);
                                Decimal.TryParse(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ty_gia"].ToString(), out result18);
                                if (result19 || result17 * result18 == new Decimal(0))
                                    return (object)false;
                                break;
                            case "thue":
                                return (object)false;
                            case "ck":
                                Decimal result20 = new Decimal(0);
                                Decimal result21 = new Decimal(0);
                                bool result22 = false;
                                bool.TryParse(values[2].ToString(), out result22);
                                Decimal.TryParse(values[0].ToString(), out result20);
                                Decimal.TryParse(values[1].ToString(), out result21);
                                if (result20 * result21 != new Decimal(0) || result22)
                                    return (object)false;
                                break;
                            case "gia":
                                Decimal result23 = new Decimal(0);
                                Decimal.TryParse(values[0].ToString(), out result23);
                                bool result24 = false;
                                bool.TryParse(values[1].ToString(), out result24);
                                if (result24)
                                    return (object)false;
                                break;
                            case "tien":
                                Decimal result25 = new Decimal(0);
                                Decimal result26 = new Decimal(0);
                                bool result27 = false;
                                bool.TryParse(values[2].ToString(), out result27);
                                Decimal.TryParse(values[0].ToString(), out result25);
                                Decimal.TryParse(values[1].ToString(), out result26);
                                if (result25 == new Decimal(0) || result25 * result26 == new Decimal(0) || result27)
                                    return (object)false;
                                break;
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
