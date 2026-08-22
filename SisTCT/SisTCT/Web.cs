using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace SisTCT
{
    public class Web
    {
        private const SecurityProtocolType securityProtocolType_0 = (SecurityProtocolType)3072;
        private const SecurityProtocolType securityProtocolType_1 = (SecurityProtocolType)768;

        static Web()
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)4080; 
        }

        public static string GetString(string url, string token, DateTime? ngay_ct = null)
        {
            string str = string.Empty;
            using (WebClient webClient = (WebClient)new WebClientWithTimeout())
            {
                webClient.UseDefaultCredentials = true;
                webClient.Encoding = Encoding.UTF8;
                webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/91.0.864.59");
                webClient.Headers[HttpRequestHeader.ContentType] = "application/json";
                if (!string.IsNullOrEmpty(token))
                    webClient.Headers.Add(HttpRequestHeader.Authorization, token);
                str = webClient.DownloadString(url);
            }
            if (ngay_ct.HasValue)
            {
                //string newValue = ngay_ct.Value.AddHours(-7.0).ToString("yyyy-MM-ddTHH:mm:ss") + "Z";
                string newValue = ngay_ct.Value.ToString("yyyy-MM-ddTHH:mm:ss") + "Z";
                str = Regex.Replace(str, "\"tdlap\"\\s*:\\s*\"[^\"]*\"", "\"tdlap\":\"" + newValue + "\"");
            }
            return str;
        }

        public static void DownloadFile(string url, string token, string filename)
        {
            using (WebClient webClient = (WebClient)new WebClientWithTimeout())
            {
                webClient.UseDefaultCredentials = true;
                webClient.Encoding = Encoding.UTF8;
                webClient.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36  Edg/91.0.864.59");
                if (!string.IsNullOrEmpty(token))
                    webClient.Headers.Add(HttpRequestHeader.Authorization, token);
                webClient.DownloadFile(url, filename);
            }
        }

        public static string PostAndGetString(string url, string jsondata)
        {
            string str = string.Empty;
            using (WebClient webClient = (WebClient)new WebClientWithTimeout())
            {
                webClient.Encoding = Encoding.UTF8;
                webClient.Headers[HttpRequestHeader.ContentType] = "application/json";
                str = webClient.UploadString(url, jsondata);
            }
            return str;
        }
    }
}
