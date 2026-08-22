using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Runtime.CompilerServices;

namespace SisTCT
{
    public class Ult
    {
        internal static string GetValueFromJsonString(string data, string name)
        {
            try
            {
                JObject jObject = JObject.Parse(data);
                return jObject.SelectToken(name).ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        internal static string FormatAuthorizationToken(string authorizationToken)
        {
            try
            {
                string result = "";
                if (!string.IsNullOrEmpty(authorizationToken))
                {
                    result = "Bearer " + authorizationToken;
                }
                return result;
            }
            catch (Exception)
            {
                return "";
            }
        }

        internal static string ConvertObjectToJSon(object o)
        {
            return JsonConvert.SerializeObject(RuntimeHelpers.GetObjectValue(o));
        }
	}
}
