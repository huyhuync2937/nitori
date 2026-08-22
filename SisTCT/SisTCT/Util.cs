using HtmlAgilityPack;
using Newtonsoft.Json;
using SasControls;
using SasControls.ControlLib;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows;
using System.CodeDom.Compiler;
using System.CodeDom;
using System.Xml.Serialization;
using System.Web.Services.Description;
using Newtonsoft.Json.Linq;

namespace SisTCT
{
    public class Util
	{
		public static DataTable getInfoMST(string mst)
		{
			DataTable dataTable = new DataTable();
			string text = "<table class=\"table-taxinfo\">";
			string text2 = "</table>";
			string url = string.Format("https://masothue.vn/Search/?q={0}&type=auto", mst);
			string webCode = getWebCode(url);
			if (!webCode.Contains("table-taxinfo"))
			{
				dataTable = ((!webCode.Contains("text-center")) ? null : convToDataTable(webCode, mst, text, text2));
			}
			else
			{
				webCode = webCode.Substring(webCode.ToLower().IndexOf(text));
				webCode = webCode.Substring(0, webCode.ToLower().IndexOf(text2) + text2.Length);
				dataTable = convert_HTML_to_DataTable(webCode);
			}
			try
			{
				DataTable mstInfo = Util.GetMstInfo(mst);
				if (mstInfo != null && mstInfo.Rows.Count > 0)
				{
					if (dataTable == null)
					{
						dataTable = new DataTable("taxInfo");
						dataTable.Columns.Add("key");
						dataTable.Columns.Add("value");
					}
					foreach (DataRow r in mstInfo.Rows)
					{
						DataRow dataRow = dataTable.Rows.Cast<DataRow>().FirstOrDefault((DataRow x) => x["key"].ToString().Trim().Equals(r["key"].ToString().Trim(), StringComparison.OrdinalIgnoreCase));
						if (dataRow == null)
						{
							dataTable.Rows.Add(r["key"], r["value"]);
						}
						else
						{
							dataRow["value"] = r["value"];
						}
					}
				}
			}
			catch (Exception)
			{
			}
			return dataTable;
		}
		private static DataTable convToDataTable(string html, string mst, string sKey, string eKey)
		{
			HtmlDocument htmlDocument = new HtmlDocument();
			htmlDocument.LoadHtml(html);
			HtmlNode htmlNode = (from x in htmlDocument.DocumentNode.Descendants("a")
								 where x.InnerText.Equals(mst)
								 select x).FirstOrDefault();
			DataTable result = null;
			if (htmlNode != null)
			{
				string text = "https://masothue.vn";
				string url = text + htmlNode.Attributes["href"].Value;
				string webCode = getWebCode(url);
				if (webCode.Contains("table-taxinfo"))
				{
					webCode = webCode.Substring(webCode.ToLower().IndexOf(sKey));
					webCode = webCode.Substring(0, webCode.ToLower().IndexOf(eKey) + eKey.Length);
					result = convert_HTML_to_DataTable(webCode);
				}
				else
				{
					convToDataTable(webCode, mst, sKey, eKey);
				}
			}
			return result;
		}
		private static string getWebCode(string url)
		{
			string result = string.Empty;
			using (WebClient webClient = new WebClient())
			{
				ServicePointManager.SecurityProtocol = (SecurityProtocolType)4080;
				ServicePointManager.DefaultConnectionLimit = 9999;
				string text = RandomString(10, false);
				string value = text + RandomNumber(1, 999999999);
				webClient.Encoding = Encoding.UTF8;
				string value2 = StartupBase.SasObj.GetSysvar("M_USER_AGENT").ToString();
				webClient.Headers.Add("user-agent", value2);
				webClient.Headers.Add(HttpRequestHeader.UserAgent, value);
				try
				{
					result = webClient.DownloadString(url);
				}
				catch (Exception ex)
				{
					if (ex.Message.Contains("(403) Forbidden."))
					{
						ExMessageBox.Show(202006012, StartupBase.SasObj, "Không có quyền truy cập máy chủ.", "SIS ERP SME", MessageBoxButton.OK, MessageBoxImage.Asterisk);
					}
				}
			}
			return result;
		}
		private static int RandomNumber(int min, int max)
		{
			Random random = new Random();
			return random.Next(min, max);
		}
		private static string RandomString(int size, bool lowerCase)
		{
			StringBuilder stringBuilder = new StringBuilder();
			Random random = new Random();
			for (int i = 0; i < size; i++)
			{
				char value = Convert.ToChar(Convert.ToInt32(Math.Floor(26.0 * random.NextDouble() + 65.0)));
				stringBuilder.Append(value);
			}
			if (lowerCase)
			{
				return stringBuilder.ToString().ToLower();
			}
			return stringBuilder.ToString();
		}
		private static DataTable convert_HTML_to_DataTable(string html)
		{
			HtmlDocument htmlDocument = new HtmlDocument();
			htmlDocument.LoadHtml(html);
			HtmlNodeCollection htmlNodeCollection = htmlDocument.DocumentNode.SelectNodes("//table/thead/tr");
			HtmlNodeCollection htmlNodeCollection2 = htmlDocument.DocumentNode.SelectNodes("//table/tbody/tr");
			DataTable dataTable = new DataTable("taxInfo");
			dataTable.Columns.Add("key");
			dataTable.Columns.Add("value");
			if (htmlNodeCollection != null)
			{
				IEnumerable<string> enumerable = from th in htmlNodeCollection[0].Elements("th")
												 select th.InnerText.Trim();
				foreach (string item in enumerable)
				{
					DataRow dataRow = dataTable.NewRow();
					dataRow["key"] = "ten_cong_ty";
					dataRow["value"] = item;
					if (!string.IsNullOrEmpty(dataRow["key"].ToString().Trim()) || !string.IsNullOrEmpty(dataRow["value"].ToString().Trim()))
					{
						dataTable.Rows.Add(dataRow);
					}
				}
			}
			if (htmlNodeCollection2 != null)
			{
				IEnumerable<string[]> enumerable2 = htmlNodeCollection2.Select((HtmlNode tr) => (from td in tr.Elements("td")
																								 select td.InnerText.Trim()).ToArray());
				IEnumerable<HtmlNode>[] array = htmlNodeCollection2.Select((HtmlNode tr) => tr.Elements("td")).ToArray();
				int num = 0;
				foreach (string[] item2 in enumerable2)
				{
					DataRow dataRow2 = dataTable.NewRow();
					dataRow2["key"] = ControlFunction.Cat_Dau(item2[0].Trim()).ToLower().Replace(" ", "_");
					if (item2.Length > 1)
					{
						if (dataRow2["key"].ToString().Trim().Equals("nguoi_dai_dien"))
						{
							string[] array2 = array[num].Select((HtmlNode td) => td.InnerHtml).ToArray();
							HtmlDocument htmlDocument2 = new HtmlDocument();
							htmlDocument2.LoadHtml(array2[1]);
							HtmlNodeCollection htmlNodeCollection3 = htmlDocument.DocumentNode.SelectNodes("//a");
							dataRow2["value"] = htmlNodeCollection3[0].InnerText.Trim();
						}
						else
						{
							dataRow2["value"] = item2[1];
						}
					}
					if (!string.IsNullOrEmpty(dataRow2["key"].ToString().Trim()) || !string.IsNullOrEmpty(dataRow2["value"].ToString().Trim()))
					{
						dataTable.Rows.Add(dataRow2);
					}
					num++;
				}
			}
			return dataTable;
		}

		public static DataTable GetMstInfo(string mst)
		{
			
			return smethod_3(mst, "GetCompanyInformation");
		}
		private static Dictionary<string, string> dictionary_0;

		private static Dictionary<string, string> dictionary_1;
		private static string string_0;

		private static string string_1;
		static Util()
		{
			dictionary_0 = new Dictionary<string, string>();
			dictionary_1 = new Dictionary<string, string>();
			string_0 = "https://api.fast.com.vn";
			string_1 = "GetCompanyInformation";
			dictionary_1.Add("Name", "ten_cong_ty");
			dictionary_1.Add("Address", "dia_chi");
			dictionary_1.Add("Status", "tinh_trang");
		}
		private static void smethod_1(params TaxInfo[] taxInfo_0)
		{
			if (taxInfo_0 != null && taxInfo_0.Length != 0)
			{
				if (taxInfo_0.Length > 50)
				{
					taxInfo_0 = taxInfo_0.Take(50).ToArray();
				}
				string string_ = string.Join("\n", taxInfo_0.Select((TaxInfo taxInFo_0) => smethod_2(taxInFo_0)).ToArray());
				StartupBase.SasObj.ExcuteNonQuery(new System.Data.SqlClient.SqlCommand(string_));
			}
		}
		private static string smethod_2(TaxInfo taxInfo_0)
		{
			if (!string.IsNullOrEmpty(taxInfo_0.Name))
			{
				taxInfo_0.Name = taxInfo_0.Name.Replace("'", "''");
			}
			return string.Format("EXEC dbo.UpdateOrInsertMst '{0}', N'{1}', N'{2}'", taxInfo_0.TaxCode, taxInfo_0.Name, smethod_6(taxInfo_0.Status));
		}
		private static DataTable smethod_3(string mst, string string_3)
		{
			if (string.IsNullOrEmpty(mst))
			{
				return null;
			}
			try
			{
				mst = mst.Trim();
				if (dictionary_0.Keys.Any((string string_1) => mst.Equals(string_1)))
				{
					return smethod_5(dictionary_0[mst]);
				}
				
				object obj = Util.smethod_0(string.Format("{0}/AppService/Service.{1}.asmx", string_0, string_1), "APIService", string_3, new object[1] { mst });
				if (obj == null)
				{
					return null;
				}
				string text = obj.ToString().Trim();
				smethod_1(JsonConvert.DeserializeObject<TaxInfo>(text));
				DataTable dataTable = smethod_5(text);
				if (dataTable != null)
				{
					dictionary_0.Add(mst, text);
				}
				return dataTable;
			}
			catch (Exception)
			{
			}
			return null;
		}
		private static string smethod_4(string string_2)
		{
			if (dictionary_1.Keys.Any((string string_1) => string_1 == string_2))
			{
				return dictionary_1[string_2];
			}
			return string_2;
		}

		private static DataTable smethod_5(string string_2)
		{
			if (string.IsNullOrEmpty(string_2))
			{
				return null;
			}
			DataTable dataTable = new DataTable("info");
			dataTable.Columns.Add("key");
			dataTable.Columns.Add("value");
			try
			{
				JObject jObject = JsonConvert.DeserializeObject<JObject>(string_2);
				new Dictionary<string, string>();
				foreach (KeyValuePair<string, JToken> item in jObject)
				{
					string text = item.Value.ToString();
					if (!string.IsNullOrEmpty(text) && !"taxcode".Equals(item.Key, StringComparison.OrdinalIgnoreCase))
					{
						if ("status".Equals(item.Key, StringComparison.OrdinalIgnoreCase))
						{
							dataTable.Rows.Add("status", item.Value.ToString());
							dataTable.Rows.Add(smethod_4(item.Key), smethod_6(item.Value.ToString()));
						}
						else
						{
							dataTable.Rows.Add(smethod_4(item.Key), text);
						}
					}
				}
			}
			catch (Exception)
			{
				return null;
			}
			return dataTable;
		}
		private static string smethod_6(string string_2)
		{
			if (string_2 != null)
			{
				switch (string_2.Length)
				{
					case 5:
						switch (string_2[1])
						{
							case '4':
								if (string_2 == "04-01")
								{
									return "Vi phạm pháp luật đã chuyển cơ quan công an";
								}
								break;
							case '0':
								if (string_2 == "00-01")
								{
									return "Vi phạm đã chuyển cơ quan công an";
								}
								break;
						}
						break;
					case 2:
						switch (string_2[1])
						{
							case '0':
								if (string_2 == "00")
								{
									return "Đang hoạt động (đã được cấp GCN ĐKT)";
								}
								break;
							case '1':
								if (string_2 == "01")
								{
									return "Ngừng hoạt động và đã đóng MST";
								}
								break;
							case '2':
								if (string_2 == "02")
								{
									return "Đã chuyển cơ quan thuế quản lý";
								}
								break;
							case '3':
								if (string_2 == "03")
								{
									return "Ngừng hoạt động nhưng chưa hoàn thành thủ tục đóng MST";
								}
								break;
							case '4':
								if (string_2 == "04")
								{
									return "Đang hoạt động (được cấp thông báo MST)";
								}
								break;
							case '5':
								if (string_2 == "05")
								{
									return "Tạm nghỉ kinh doanh có thời hạn";
								}
								break;
							case '6':
								if (string_2 == "06")
								{
									return "Không hoạt động tại địa chỉ đã đăng ký";
								}
								break;
						}
						break;
				}
			}
			return string_2;
		}



		public static object smethod_0(string string_0, string string_1, string string_2, object[] object_0)
		{
			try
			{
				WebClient webClient = new WebClient();
				webClient.Proxy = WebRequest.GetSystemWebProxy();
				Stream stream = webClient.OpenRead(string_0 + "?wsdl");
				ServiceDescription serviceDescription = ServiceDescription.Read(stream);
				ServiceDescriptionImporter serviceDescriptionImporter = new ServiceDescriptionImporter();
				serviceDescriptionImporter.ProtocolName = "Soap12";
				serviceDescriptionImporter.AddServiceDescription(serviceDescription, null, null);
				serviceDescriptionImporter.Style = ServiceDescriptionImportStyle.Client;
				serviceDescriptionImporter.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties;
				CodeNamespace codeNamespace = new CodeNamespace();
				CodeCompileUnit codeCompileUnit = new CodeCompileUnit();
				codeCompileUnit.Namespaces.Add(codeNamespace);
				if (serviceDescriptionImporter.Import(codeNamespace, codeCompileUnit) == (ServiceDescriptionImportWarnings)0)
				{
					CodeDomProvider codeDomProvider = CodeDomProvider.CreateProvider("C#");
					string[] assemblyNames = new string[5] { "System.dll", "System.Web.Services.dll", "System.Web.dll", "System.Xml.dll", "System.Data.dll" };
					CompilerParameters compilerParameters = new CompilerParameters(assemblyNames);
					compilerParameters.GenerateInMemory = true;
					CompilerResults compilerResults = codeDomProvider.CompileAssemblyFromDom(compilerParameters, codeCompileUnit);
					if (compilerResults.Errors.Count > 0)
					{
						return null;
					}
					object object_ = compilerResults.CompiledAssembly.CreateInstance(string_1);
					smethod_1(ref object_, "Timeout", 60000000);
					MethodInfo method = object_.GetType().GetMethod(string_2);
					return method.Invoke(object_, object_0);
				}
				return null;
			}
			catch (Exception)
			{
				return null;
			}
		}
		private static void smethod_1(ref object object_0, string string_0, object object_1)
		{
			if (object_0 == null)
			{
				return;
			}
			try
			{
				PropertyInfo property = object_0.GetType().GetProperty(string_0);
				if (property != null)
				{
					property.SetValue(object_0, object_1, null);
				}
			}
			catch (Exception)
			{
			}
		}
	}
	public class TaxInfo
	{
		public string TaxCode;

		public string Name;

		public string Address;

		public string Status;
	}
}
