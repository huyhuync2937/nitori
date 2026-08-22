using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SisTCT
{
	public class AesOperation
	{
		public static string Encrypt(string StrInput)
		{
			if (string.IsNullOrEmpty(StrInput.ToString().Trim()))
			{
				return "";
			}
			return EncryptString("b14ca5898a4e4133bbce2ea2315a1916", StrInput);
		}

		public static string Decrypt(string StrInput)
		{
			if (string.IsNullOrEmpty(StrInput.ToString().Trim()))
			{
				return "";
			}
			return DecryptString("b14ca5898a4e4133bbce2ea2315a1916", StrInput);
		}

		public static string EncryptString(string key, string plainText)
		{
			byte[] iV = new byte[16];
			byte[] inArray;
			using (Aes aes = Aes.Create())
			{
				aes.Key = Encoding.UTF8.GetBytes(key);
				aes.IV = iV;
				ICryptoTransform transform = aes.CreateEncryptor(aes.Key, aes.IV);
				using (MemoryStream memoryStream = new MemoryStream())
				{
					using (CryptoStream stream = new CryptoStream(memoryStream, transform, CryptoStreamMode.Write))
					{
						using (StreamWriter streamWriter = new StreamWriter(stream))
						{
							streamWriter.Write(plainText);
						}
						inArray = memoryStream.ToArray();
					}
				}
			}
			return Convert.ToBase64String(inArray);
		}

		public static string DecryptString(string key, string cipherText)
		{
			byte[] iV = new byte[16];
			byte[] buffer = Convert.FromBase64String(cipherText);
			using (Aes aes = Aes.Create())
			{
				aes.Key = Encoding.UTF8.GetBytes(key);
				aes.IV = iV;
				ICryptoTransform transform = aes.CreateDecryptor(aes.Key, aes.IV);
				using (MemoryStream stream = new MemoryStream(buffer))
				{
					using (CryptoStream stream2 = new CryptoStream(stream, transform, CryptoStreamMode.Read))
					{
						using (StreamReader streamReader = new StreamReader(stream2))
						{
							return streamReader.ReadToEnd();
						}
					}
				}
			}
		}
	}

}
