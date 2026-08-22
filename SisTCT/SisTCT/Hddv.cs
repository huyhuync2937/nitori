using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using SasControls;

namespace SisTCT
{
    public class Hddv
    {
        //private const string url_api_hoadondientu = "https://hoadondientu.gdt.gov.vn:30000";
        private const string url_api_hoadondientu = "https://hoadondientu.gdt.gov.vn/api";
        public static bool CanGetInvoice(DateTime ngay1, DateTime ngay2)
        {
            bool flag;
            try
            {
                    DataSet dataSet = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT DISTINCT CONVERT(varchar(7), ngay_ct0, 120) as namthang FROM ph92"));
                    if (dataSet == null)
                    {
                        flag = false;
                        goto label_13;
                    }
                    else
                    {
                        DataTable table = dataSet.Tables[0];
                        table.DefaultView.Sort = "namthang";
                        string str1 = ngay1.ToString("yyyy-MM");
                        string str2 = ngay2.ToString("yyyy-MM");
                        if (table.DefaultView.FindRows((object)str1).Length == 0)
                            table.Rows.Add((object)str1);
                        if (table.DefaultView.FindRows((object)str2).Length == 0)
                            table.Rows.Add((object)str2);
                        if (table.Rows.Count > 3)
                        {
                            flag = false;
                            goto label_13;
                        }
                    }
            }
            catch (Exception ex)
            {
                flag = false;
                goto label_13;
            }
            flag = true;
            label_13:
            return flag;
        }

        public static CaptchaInfo GetTaxCaptcha()
        {
            try
            {
                CaptchaInfo captchaInfo = new CaptchaInfo();
                string @string = Web.GetString("https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + "/captcha", "");
                if (string.IsNullOrEmpty(@string))
                    return CaptchaInfo.captchaInfo_0;
                string key = Ult.GetValueFromJsonString(@string, "key");
                string str = Ult.GetValueFromJsonString(@string, "content");
                return new CaptchaInfo(str, key, CaptchaHelper.Detect(str));
            }
            catch (Exception ex)
            {
                string err = ex.Message + "\r\n--------------------\r\n" + ex.StackTrace;
                if (ex.InnerException != null)
                    err = err + "\r\n--------------------\r\n" + ex.InnerException.Message + "\r\n--------------------\r\n" + ex.InnerException.StackTrace;
                return new CaptchaInfo(err);
            }
        }

        public static string GetTaxToken(
          string userName,
          string password,
          string key,
          string captcha,
          string value)
        {
            try
            {
                string str = Ult.FormatAuthorizationToken(Ult.GetValueFromJsonString(Web.PostAndGetString("https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + "/security-taxpayer/authenticate", Ult.ConvertObjectToJSon((object)new Dictionary<string, string>()
        {
          {
            "username",
            userName
          },
          {
            nameof (password),
            password
          },
          {
            "cvalue",
            value
          },
          {
            "ckey",
            key
          }
        })), "token"));
                if (!string.IsNullOrEmpty(str))
                    CaptchaHelper.Update(captcha, value);
                return str;
            }
            catch (Exception ex)
            {
                return "Err: " + ex.Message;
            }
        }

        public static string GetTaxInvoiceList(
          string authorizationToken,
          DateTime dateFrom,
          DateTime dateTo,
          string dataTaxCode,
          string dataForm,
          string dataSeri,
          string dataNumber,
          string dataInvoiceStatus,
          string dataProcessStatus,
          string pageState)
        {
            try
            {
                string str1 = "";
                string str2 = "purchase";
                string str3 = "query";
                if (!string.IsNullOrEmpty(dataTaxCode))
                    str1 = str1 + ";nbmst==" + Hddv.ReplaceVariable(dataTaxCode);
                if (!string.IsNullOrEmpty(dataInvoiceStatus))
                    str1 = str1 + ";tthai==" + Hddv.ReplaceVariable(dataInvoiceStatus);
                if (!string.IsNullOrEmpty(dataProcessStatus))
                {
                    str1 = str1 + ";ttxly==" + Hddv.ReplaceVariable(dataProcessStatus);
                    if (dataProcessStatus == "8")
                        str3 = "sco-query";
                }
                if (!string.IsNullOrEmpty(dataForm))
                    str1 = str1 + ";khmshdon==" + Hddv.ReplaceVariable(dataForm);
                if (!string.IsNullOrEmpty(dataNumber))
                    str1 = str1 + ";shdon==" + Hddv.ReplaceVariable(dataNumber);
                if (!string.IsNullOrEmpty(dataSeri))
                    str1 = str1 + ";khhdon==" + Hddv.ReplaceVariable(dataSeri);
                string str4 = "";
                if (!string.IsNullOrEmpty(pageState))
                    str4 = "&state=" + pageState;
                return Web.GetString("https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/{5}/invoices/{0}?sort=tdlap:desc&size=50{4}&search=tdlap=ge={1}T00:00:00;tdlap=le={2}T23:59:59{3}", (object)str2, (object)dateFrom.ToString("dd/MM/yyyy").Replace('-', '/'), (object)dateTo.ToString("dd/MM/yyyy").Replace('-', '/'), (object)str1, (object)str4, (object)str3), authorizationToken);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static string DownloadFile(
          string authorizationToken,
          string dataTaxCode,
          string dataForm,
          string dataSeri,
          string dataNumber,
          string filename)
        {
            return Hddv.DownloadFile(authorizationToken, dataTaxCode, dataForm, dataSeri, dataNumber, filename, "6");
        }

        public static string DownloadFile(
          string authorizationToken,
          string dataTaxCode,
          string dataForm,
          string dataSeri,
          string dataNumber,
          string filename,
          string ttxly)
        {
            string url = "";
            try
            {
                string str1 = Hddv.ReplaceVariable(dataTaxCode);
                string str2 = Hddv.ReplaceVariable(dataForm);
                string str3 = Hddv.ReplaceVariable(dataNumber);
                string str4 = Hddv.ReplaceVariable(dataSeri);
                if (ttxly == "8")
                    url = "https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/sco-query/invoices/export-xml?nbmst={0}&khhdon={1}&shdon={2}&khmshdon={3}", str1, str4, str3, str2);
                else
                    url = "https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/query/invoices/export-xml?nbmst={0}&khhdon={1}&shdon={2}&khmshdon={3}", str1, str4, str3, str2);
                Web.DownloadFile(url, authorizationToken, filename);
                return "OK";
            }
            catch (Exception ex)
            {
                return ex.Message + "\n" + url;
            }
        }

        public static string GetInvDetail(
          string authorizationToken,
          string dataTaxCode,
          string dataForm,
          string dataSeri,
          string dataNumber,
          string ttxly)
        {
            string url = "";
            try
            {
                string str1 = Hddv.ReplaceVariable(dataTaxCode);
                string str2 = Hddv.ReplaceVariable(dataForm);
                string str3 = Hddv.ReplaceVariable(dataNumber);
                string str4 = Hddv.ReplaceVariable(dataSeri);
                if (ttxly == "8")
                    url = "https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/sco-query/invoices/detail?nbmst={0}&khhdon={1}&shdon={2}&khmshdon={3}", str1, str4, str3, str2);
                else
                    url = "https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/query/invoices/detail?nbmst={0}&khhdon={1}&shdon={2}&khmshdon={3}", str1, str4, str3, str2);
                return Web.GetString(url, authorizationToken);
            }
            catch (Exception ex)
            {
                return ex.Message + "\n" + url;
            }
        }

        public static string Base64Decode(string base64EncodedData)
        {
            return Encoding.UTF8.GetString(Convert.FromBase64String(base64EncodedData));
        }

        private static string ReplaceVariable(string value)
        {
            value = value.Replace("'", "").Replace(";", "");
            return value;
        }

        public static string GetTaxInvoiceListOut(
          string authorizationToken,
          DateTime dateFrom,
          DateTime dateTo,
          string dataTaxCode,
          string dataForm,
          string dataSeri,
          string dataNumber,
          string dataInvoiceStatus,
          string dataProcessStatus,
          string pageState)
        {
            try
            {
                string str1 = "";
                string str2 = "sold";
                string str3 = "query";
                if (!string.IsNullOrEmpty(dataTaxCode))
                    str1 = str1 + ";nmmst==" + Hddv.ReplaceVariable(dataTaxCode);
                if (!string.IsNullOrEmpty(dataInvoiceStatus))
                    str1 = str1 + ";tthai==" + Hddv.ReplaceVariable(dataInvoiceStatus);
                if (!string.IsNullOrEmpty(dataProcessStatus))
                {
                    str1 = str1 + ";ttxly==" + Hddv.ReplaceVariable(dataProcessStatus);
                    if (dataProcessStatus == "8")
                        str3 = "sco-query";
                }
                if (!string.IsNullOrEmpty(dataForm))
                    str1 = str1 + ";khmshdon==" + Hddv.ReplaceVariable(dataForm);
                if (!string.IsNullOrEmpty(dataNumber))
                    str1 = str1 + ";shdon==" + Hddv.ReplaceVariable(dataNumber);
                if (!string.IsNullOrEmpty(dataSeri))
                    str1 = str1 + ";khhdon==" + Hddv.ReplaceVariable(dataSeri);
                string str4 = "";
                if (!string.IsNullOrEmpty(pageState))
                    str4 = "&state=" + pageState;
                return Web.GetString("https://hoadondientu.gdt.gov.vn/api".TrimEnd('/') + string.Format("/{5}/invoices/{0}?sort=tdlap:desc&size=50{4}&search=tdlap=ge={1}T00:00:00;tdlap=le={2}T23:59:59{3}", str2, dateFrom.ToString("dd/MM/yyyy").Replace('-', '/'), dateTo.ToString("dd/MM/yyyy").Replace('-', '/'), str1, str4, str3), authorizationToken);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}

