using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Diagnostics;
using SasControls;
using System.Windows;

namespace SasIeCt
{
    public class C_GetDataExcel
	{
		public static string StrBrowse = "";

		public static string StrBrowseFieldNull = "";

		public static string TCNV3String = "ÊÈèÉÌéÐÒÕúóÓÔíÝáãìõâêòµ\u00b8¶·¹\u00a8»¾¼½Æ©ÇË®ÎÏÑªÖ×ØÜÞßäôù«åæç¬ëîïñ­øö÷ýûüþ¡¢§£¤¥¦";

		public static string UnicodeString = "ấẩốẫèộéềếỳúểễớíỏóỡừõờũàáảãạăằắẳẵặâầậđẻẽẹêệìỉĩịòọụựôồổỗơởợùủưứửữýỷỹỵĂÂĐÊÔƠƯ";

		public static DataSet GetData(ImportInfo info, string convertFont)
		{
			DataSet dtImport = new DataSet();
			try
			{
				DataTable dataFromExcel = GetDataFromExcel(info.FileName, info.FieldNotNull);
				DataTable structTable = GetStructTable(info.TableTemplate);
				if (dataFromExcel == null || structTable == null)
				{
					return null;
				}
				dtImport.Tables.Add(dataFromExcel.Copy());
				dtImport.Tables.Add(structTable.Copy());
				ConverDateTime(ref dtImport);
				if (!(convertFont.Trim() == "2"))
				{
					return dtImport;
				}
				ConverFont(ref dtImport);
				return dtImport;
			}
			catch (Exception ex)
			{
				if (StartUp.waiting != null)
				{
					StartUp.waiting.Close();
				}
				MessageBox.Show("Lỗi cột column");
				ExMessageBox.Show(140, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				return null;
			}
		}

		public void setupPackage()
		{
		}

		public static DataTable GetDataFromExcel(string _fileName, string field_not_null)
		{
			string text = "Yes";
			string text2 = "";
			text2 = ((!_fileName.Contains(".xlsx")) ? ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=\"Excel 8.0;HDR=" + text + ";IMEX=1\"") : ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileName + ";Extended Properties=\"Excel 12.0;HDR=" + text + ";IMEX=1\""));
			DataSet dataSet = new DataSet();
			using (OleDbConnection oleDbConnection = new OleDbConnection(text2))
			{
				bool flag = false;
				try
				{
					oleDbConnection.Open();
					flag = true;
				}
				catch (Exception ex)
				{
					if (ex.Message.Contains("OLEDB"))
					{
						try
						{
							string text3 = "AccessDatabaseEngine.exe";
							StartupBase.SasObj.SynchroFile(".", text3);
							Process process = new Process();
							process.StartInfo.FileName = text3;
							process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
							process.StartInfo.Arguments = "/quiet";
							process.Start();
							process.WaitForExit();
							oleDbConnection.Open();
							flag = true;
						}
						catch (Exception ex2)
						{
							ExMessageBox.Show(210, StartupBase.SasObj, "[" + ex2.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
							return null;
						}
					}
				}
				if (flag)
				{
					try
					{
						DataTable oleDbSchemaTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[4]
						{
							null,
							null,
							null,
							"TABLE"
						});
						IEnumerator enumerator = oleDbSchemaTable.Rows.GetEnumerator();
						try
						{
							if (enumerator.MoveNext())
							{
								DataRow dataRow = (DataRow)enumerator.Current;
								string text4 = dataRow["TABLE_NAME"].ToString();
								OleDbCommand oleDbCommand = new OleDbCommand("SELECT * FROM [" + text4 + "]", oleDbConnection);
								oleDbCommand.CommandType = CommandType.Text;
								DataTable dataTable = new DataTable(text4);
								dataTable.Columns.Add("Stt_(stt):IV", typeof(int));
								dataTable.Columns[0].AutoIncrement = true;
								dataTable.Columns[0].AutoIncrementSeed = 2L;
								dataTable.Columns[0].AutoIncrementStep = 1L;
								dataSet.Tables.Add(dataTable);
								new OleDbDataAdapter(oleDbCommand).Fill(dataTable);
							}
						}
						finally
						{
							IDisposable disposable = enumerator as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
					}
					catch (Exception ex3)
					{
						if (StartUp.waiting != null)
						{
							StartUp.waiting.Close();
						}
						ExMessageBox.Show(150, StartupBase.SasObj, "[" + ex3.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
						oleDbConnection.Close();
						return null;
					}
					oleDbConnection.Close();
				}
			}
			if (dataSet.Tables.Count == 0)
			{
				return null;
			}
			DataTable tbExcel = XoaDongTrang(dataSet.Tables[0]);
			XoaCotTrang(ref tbExcel);
			tbExcel.TableName = "DataExcel";
			return SetColumnName(tbExcel, field_not_null);
		}

		public static DataTable XoaDongTrang(DataTable tbExcel)
		{
			bool flag = false;
			DataTable dataTable = tbExcel.Clone();
			for (int i = 0; i < tbExcel.Rows.Count; i++)
			{
				flag = false;
				for (int j = 1; j < tbExcel.Columns.Count; j++)
				{
					if (tbExcel.Rows[i][j].ToString().Trim() != "")
					{
						flag = true;
						break;
					}
				}
				if (flag)
				{
					dataTable.ImportRow(tbExcel.Rows[i]);
				}
			}
			return dataTable;
		}

		public static void XoaCotTrang(ref DataTable tbExcel)
		{
			string text = "";
			int result = 0;
			for (int i = 0; i < tbExcel.Columns.Count; i++)
			{
				text = tbExcel.Columns[i].ColumnName.Trim();
				if (text.IndexOf("F") == 0 && int.TryParse(text.Substring(1, text.Length - 1), out result))
				{
					tbExcel.Columns.Remove(text);
					i = -1;
				}
			}
		}

		private static DataTable SetColumnName(DataTable tb, string field_not_null)
		{
			try
			{
				StrBrowse = "";
				StrBrowseFieldNull = "stt:H=Dòng";
				string text = "";
				string[] source = field_not_null.Split(';');
				for (int i = 0; i < tb.Columns.Count; i++)
				{
					string text2 = tb.Columns[i].ColumnName.ToString();
					int num = text2.LastIndexOf('(');
					int num2 = text2.LastIndexOf(')');
					if (num == -1 || num2 == -1)
					{
						if (StartUp.waiting != null)
						{
							StartUp.waiting.Close();
						}
						ExMessageBox.Show(160, StartupBase.SasObj, $"Tên cột << [{text2}] >> trong file  excel không đúng định dạng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
						return null;
					}
					string text3 = text2.Substring(num + 1, num2 - num - 1).Trim();
					string text4 = text3 + ":H=" + text2.Substring(0, num - 1).Trim() + text2.Substring(num2 + 1, text2.Length - num2 - 1).Trim();
					text = text + text4 + ";";
					tb.Columns[i].ColumnName = text3;
					if (tb.Columns[i].DataType == typeof(DateTime))
					{
						text4 += ":D";
					}
					if (source.Contains(text3))
					{
						StrBrowseFieldNull = StrBrowseFieldNull + ";" + text4;
					}
				}
				StrBrowse = text.Substring(0, text.Length - 1);
				StrBrowse = StrBrowse.Replace("#", ".");
				StrBrowseFieldNull = StrBrowseFieldNull.Replace("#", ".");
				return tb;
			}
			catch (Exception ex)
			{
				ExMessageBox.Show(165, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				return null;
			}
		}

		public static DataTable GetStructTable(string tableName)
		{
			if (tableName == null || tableName == "")
			{
				return null;
			}
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.CommandText += $"SELECT * FROM information_schema.columns WHERE table_name like '{tableName.Trim()}'";
			DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlCommand);
			if (dataSet == null || dataSet.Tables.Count == 0)
			{
				return null;
			}
			dataSet.Tables[0].TableName = "StrucImex";
			return dataSet.Tables[0];
		}

		public static void ConverFont(ref DataSet dtImport)
		{
			for (int i = 0; i < dtImport.Tables["DataExcel"].Rows.Count; i++)
			{
				for (int j = 0; j < dtImport.Tables["DataExcel"].Columns.Count; j++)
				{
					DataRow[] array = dtImport.Tables["StrucImex"].Select("column_name = '" + dtImport.Tables["DataExcel"].Columns[j] + "'");
					if (array.Length <= 0)
					{
						continue;
					}
					int result = 0;
					if (int.TryParse(array[0]["character_maximum_length"].ToString(), out result) && result != -1)
					{
						string value = ConvertTcvn3ToUnicode(dtImport.Tables["DataExcel"].Rows[i][j].ToString().Trim());
						if (!string.IsNullOrEmpty(value))
						{
							dtImport.Tables["DataExcel"].Rows[i][j] = value;
						}
					}
				}
			}
		}

		private static void ConverDateTime(ref DataSet dtImport)
		{
			for (int i = 0; i < dtImport.Tables["DataExcel"].Rows.Count; i++)
			{
				for (int j = 0; j < dtImport.Tables["DataExcel"].Columns.Count; j++)
				{
					DataRow[] array = dtImport.Tables["StrucImex"].Select("column_name = '" + dtImport.Tables["DataExcel"].Columns[j] + "'");
					if (array.Length > 0 && array[0]["data_type"].ToString().Trim() == "smalldatetime" && !dtImport.Tables["DataExcel"].Rows[i][j].GetType().FullName.Equals("System.DateTime"))
					{
						string[] array2 = dtImport.Tables["DataExcel"].Rows[i][j].ToString().Replace(" ", "").Replace("/", "-")
							.Split('-');
						if (array2.Length == 3)
						{
							dtImport.Tables["DataExcel"].Rows[i][j] = array2[2].Substring(0, 4) + ((array2[1].Length > 1) ? array2[1] : ("0" + array2[1])) + ((array2[0].Length > 1) ? array2[0] : ("0" + array2[0]));
						}
					}
				}
			}
		}

		public static List<char> GetAllChar(string s)
		{
			List<char> list = new List<char>();
			int length = s.Length;
			for (int i = 0; i < length; i++)
			{
				if (!list.Contains(s[i]) && TCNV3String.Contains(s[i]))
				{
					list.Add(s[i]);
				}
			}
			int count = list.Count;
			for (int j = 0; j < count - 1; j++)
			{
				for (int k = j + 1; k < count; k++)
				{
					if (TCNV3String.IndexOf(list[j]) > TCNV3String.IndexOf(list[k]))
					{
						char value = list[j];
						list[j] = list[k];
						list[k] = value;
					}
				}
			}
			return list;
		}

		public static string ConvertTcvn3ToUnicode(string input)
		{
			input = input.Trim();
			List<char> allChar = GetAllChar(input);
			int num = allChar.Count();
			for (int i = 0; i < num; i++)
			{
				int num2 = TCNV3String.IndexOf(allChar[i]);
				if (num2 >= 0)
				{
					input = input.Replace(TCNV3String[num2], UnicodeString[num2]);
				}
			}
			return input;
		}
	}
}
