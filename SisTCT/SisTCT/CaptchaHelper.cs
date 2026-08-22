using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SisTCT
{
    internal class CaptchaHelper
    {
        internal static string Detect(string svg)
        {
            GetMd5Hash(svg);
            return "";
        }

        internal static void Update(string svg, string value)
        {
            string md5Hash = GetMd5Hash(svg);
            //Sys.Func.UploadCapt(svg, value);
        }
        public static string GetMd5Hash(string str)
        {
            MD5 mD = MD5.Create();
            byte[] array = mD.ComputeHash(Encoding.UTF8.GetBytes(str));
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < array.Length; i++)
            {
                stringBuilder.Append(array[i].ToString("x2"));
            }
            return stringBuilder.ToString();
        }
    }
}