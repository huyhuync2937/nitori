using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Threading;
using System.Threading;
using SasControls;
using SasFormBrowes;
using SasErrorLib;

namespace SasIeCt
{
    public class C_ImportVoucher
    {
		public static string strException = "";

		public static DataTable tbStt_rec = new DataTable();

		public static DataTable tbDuLieuTrung = new DataTable();

		public static DataTable tb_Post_Error = new DataTable();

		public static DataTable tb_Post_Ok = new DataTable();

		public static bool _flag_post = true;

		public static bool UploadTable(string tableName, DataSet db, string strbrowse)
		{
			strException = "";
			DeleteTableOld(tableName);
			DataTable dataTable = db.Tables["DataExcel"].Clone();
			dataTable.Columns.Add("ly_do");
			foreach (DataRow row in db.Tables["DataExcel"].Rows)
			{
				SqlConnection sqlConnection = new SqlConnection(GetConnectionString(10000));
				sqlConnection.Open();
				try
				{
					SqlCommand uploadCommand = GetUploadCommand(tableName, row);
					uploadCommand.Connection = sqlConnection;
					uploadCommand.ExecuteNonQuery();
				}
				catch (SqlException ex)
				{
					strException = strException + ex.Message + "\n";
					DataRow dataRow2 = dataTable.NewRow();
					dataRow2.ItemArray = row.ItemArray;
					dataRow2["ly_do"] = ex.Message;
					dataTable.Rows.Add(dataRow2);
				}
				catch (Exception ex2)
				{
					ErrorLog.CatchMessage(ex2);
				}
				finally
				{
					sqlConnection.Close();
				}
			}
			if (dataTable.Rows.Count > 0)
			{
				DeleteTableOld(tableName);
				if (StartUp.waiting != null)
				{
					StartUp.waiting.Close();
				}
				FormBrowse br = new FormBrowse(StartupBase.SasObj, dataTable.DefaultView, (StartupBase.M_LAN.Equals("V") ? "ly_do:320:H=Lý do lỗi;" : "ly_do:320:H=Error description;") + strbrowse);
				br.frmBrw.Title = (StartupBase.M_LAN.Equals("V") ? "So lieu loi" : "The data errors");
				br.frmBrw.LanguageID = "SasIeCt_5";
				br.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)delegate
				{
					br.frmBrw.ChangeLanguage();
				});
				br.ShowDialog();
				return false;
			}
			return true;
		}

		private static SqlCommand GetUploadCommand(string tableName, DataRow row)
		{
			string text = "";
			SqlCommand sqlCommand = new SqlCommand();
			try
			{
				for (int i = 0; i < row.Table.Columns.Count; i++)
				{
					string text2 = row.Table.Columns[i].ColumnName.Trim();
					text = ((!(text == "")) ? (text + $",@{text2}") : $"@{text2}");
					sqlCommand.Parameters.Add(new SqlParameter($"@{text2}", row[text2]));
					sqlCommand.CommandText = string.Format("INSERT INTO {0} ({1}) VALUES ({2})", tableName, text.Replace("@", ""), text);
				}
			}
			catch (Exception ex)
			{
				ErrorLog.CatchMessage(ex);
				throw;
			}
			string debugCommandText = sqlCommand.CommandText;
			foreach (SqlParameter parameter in sqlCommand.Parameters)
			{
				debugCommandText = debugCommandText.Replace(parameter.ParameterName, "'" + parameter.Value + "'");
			}
			Console.WriteLine(debugCommandText);
			return sqlCommand;
		}

		private static void DeleteTableOld(string tableName)
		{
			SqlConnection sqlConnection = new SqlConnection(GetConnectionString(10000));
			sqlConnection.Open();
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.CommandText = $"Delete from {tableName}";
			sqlCommand.Connection = sqlConnection;
			sqlCommand.ExecuteNonQuery();
			sqlConnection.Close();
		}

		public static string GetConnectionString(int _time)
		{
			string text = "";
			new DataSet();
			try
			{
				text = StartupBase.SasObj.M_ConnectString;
				if (!text.ToUpper().Contains("TIMEOUT"))
				{
					return text + ";Connect Timeout = " + _time;
				}
				return text;
			}
			catch (Exception ex)
			{
				ErrorLog.CatchMessage(ex);
				return "";
			}
		}

		public static bool Check_Data(ImportInfo info)
		{
			bool flag = false;
			try
			{
				SqlConnection sqlConnection = new SqlConnection(GetConnectionString(500000));
				SqlCommand sqlCommand = new SqlCommand();
				sqlCommand.CommandText = $"EXEC {info.PostProc.Split(',')[0].ToString()} {StartUp._User_id}, '{info.Ma_qs}', '{StartupBase.SasObj.M_ma_dvcs.Trim()}', '{info.Xy_ly.Trim()}'";
				sqlCommand.Connection = sqlConnection;
				sqlCommand.CommandTimeout = 600000;
				string a = StartupBase.SasObj.SqlString(sqlCommand);
				SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
				DataSet dataSet = new DataSet();
				sqlDataAdapter.Fill(dataSet);
				sqlConnection.Close();
				flag = ShowError(dataSet, info.Ma_Imex);
				if (!flag)
				{
					return flag;
				}
				tbStt_rec = dataSet.Tables[dataSet.Tables.Count - 1].Copy();
				tbDuLieuTrung = dataSet.Tables[dataSet.Tables.Count - 2].Copy();
				return flag;
			}
			catch (Exception ex)
			{
				if (StartUp.waiting != null)
				{
					StartUp.waiting.Close();
				}
				ExMessageBox.Show(873, StartupBase.SasObj, ex.Message, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				return false;
			}
		}

		public static bool Post(ImportInfo info)
		{
			tb_Post_Error = tbStt_rec.Clone();
			tb_Post_Error.Columns.Add("ly_do_loi", typeof(string));
			int count = tbStt_rec.Rows.Count;
			StartUp.waiting.pgValue = count;
			StartUp.waiting.Set(0.0);
			bool flag = true;
			SqlConnection sqlConnection = new SqlConnection(GetConnectionString(500000));
			SqlCommand sqlCommand = new SqlCommand();
			sqlCommand.CommandTimeout = 600000;
			sqlCommand.Connection = sqlConnection;
			sqlConnection.Open();
			int num = 0;
			foreach (DataRow row in tbStt_rec.Rows)
			{
				StartUp.waiting.Set(num);
				try
				{
					sqlCommand.CommandText = string.Format("EXEC {0} '{1}'", info.PostProc.Split(',')[1].ToString(), row["stt_rec"].ToString().Trim());
					sqlCommand.ExecuteNonQuery();
				}
				catch (Exception ex)
				{
					DataRow dataRow2 = tb_Post_Error.NewRow();
					dataRow2.ItemArray = row.ItemArray;
					dataRow2["ly_do_loi"] = ex.Message;
					tb_Post_Error.Rows.Add(dataRow2);
					flag = false;
				}
				finally
				{
					num++;
					StartUp.waiting.Set(num);
				}
			}
			StartUp.waiting.Set(StartUp.waiting.pgValue);
			if (!flag)
			{
				int num2 = count;
				foreach (DataRow row2 in tbStt_rec.Rows)
				{
					sqlCommand.CommandText = string.Format("EXEC [dbo].[DeleteVoucher] '{0}', '{1}'", info.Ma_ct.Trim(), row2["stt_rec"].ToString().Trim());
					sqlCommand.ExecuteNonQuery();
					StartUp.waiting.Set_Post_Error(num2--);
				}
				if (StartUp.waiting != null)
				{
					StartUp.waiting.Close();
				}
				FormBrowse formBrowse = new FormBrowse(StartupBase.SasObj, tb_Post_Error.DefaultView, StartupBase.M_LAN.Equals("V") ? "ly_do_loi:H=Lý do lỗi;ngay_ct:H=Ngày c.từ:130;ma_qs:H=Mã quyển c.từ  :130;so_ct:H=Số c.từ:130;stt_rec:0:H=stt_rec" : "ly_do_loi:H=Error description;ngay_ct:H=Voucher date:130;ma_qs:H=Book code:130;so_ct:H=Voucher no.:130;stt_rec:0:H=Stt_rec");
				formBrowse.frmBrw.Title = (StartupBase.M_LAN.Equals("V") ? "Các chứng từ bị lỗi khi post" : "The post data errors");
				formBrowse.frmBrw.LanguageID = "SasIeCt_15";
				formBrowse.ShowDialog();
			}
			else if (info.Xy_ly.Trim() == "1")
			{
				int num3 = 0;
				StartUp.waiting.pgValue = tbDuLieuTrung.Rows.Count;
				foreach (DataRow row3 in tbDuLieuTrung.Rows)
				{
					sqlCommand.CommandText = string.Format("EXEC [dbo].[DeleteVoucher] '{0}', '{1}'", info.Ma_ct.Trim(), row3["stt_rec"].ToString().Trim());
					sqlCommand.ExecuteNonQuery();
					StartUp.waiting.Set_Delete(num3++);
				}
				StartUp.waiting.Set_Delete(StartUp.waiting.pgValue);
			}
			sqlConnection.Close();
			return flag;
		}

		public static bool Post_All(ImportInfo info)
		{
			int dem = 0;
			int count = tbStt_rec.Rows.Count;
			StartUp.waiting.pgValue = count;
			StartUp.waiting.Set(0.0);
			int num = 10;
			tb_Post_Ok = tbStt_rec.Clone();
			tb_Post_Error = tbStt_rec.Clone();
			tb_Post_Error.Columns.Add("ly_do_loi", typeof(string));
			bool _flag_post = true;
			int num2 = count / num;
			int num3 = count % num;
			ManualResetEvent[] resetEvents;
			for (int i = 0; i <= num2; i++)
			{
				int num4 = num;
				if (i == num2 && num3 == 0)
				{
					break;
				}
				if (i == num2 && num3 != 0)
				{
					num4 = num3;
				}
				resetEvents = new ManualResetEvent[num4];
				for (int j = 0; j < num4; j++)
				{
					int num5 = i * num + j;
					resetEvents[j] = new ManualResetEvent(initialState: false);
					ThreadPool.QueueUserWorkItem(delegate (object data)
					{
						int index = int.Parse(((string)data).Split(';')[0].Trim());
						int num8 = int.Parse(((string)data).Split(';')[1].Trim());
						SqlConnection sqlConnection3 = new SqlConnection(GetConnectionString(50000));
						sqlConnection3.Open();
						try
						{
							string arg = tbStt_rec.Rows[index]["stt_rec"].ToString();
							SqlCommand sqlCommand3 = new SqlCommand();
							sqlCommand3.Connection = sqlConnection3;
							sqlCommand3.CommandText = $"EXEC {info.PostProc.Split(',')[1].Trim()} '{arg}'";
							sqlCommand3.ExecuteNonQuery();
							DataRow dataRow3 = tb_Post_Ok.NewRow();
							dataRow3.ItemArray = tbStt_rec.Rows[index].ItemArray;
							tb_Post_Ok.Rows.Add(dataRow3);
						}
						catch (Exception ex)
						{
							DataRow dataRow4 = tb_Post_Error.NewRow();
							dataRow4.ItemArray = tbStt_rec.Rows[index].ItemArray;
							dataRow4["ly_do_loi"] = ex.Message;
							tb_Post_Error.Rows.Add(dataRow4);
							_flag_post = false;
						}
						finally
						{
							sqlConnection3.Close();
							dem++;
							resetEvents[num8].Set();
						}
					}, num5.ToString() + ";" + j.ToString());
				}
				ManualResetEvent[] array = resetEvents;
				foreach (ManualResetEvent manualResetEvent in array)
				{
					manualResetEvent.WaitOne();
					StartUp.waiting.Set(dem);
				}
			}
			StartUp.waiting.Set(StartUp.waiting.pgValue);
			if (!_flag_post)
			{
				SqlConnection sqlConnection = new SqlConnection(GetConnectionString(500000));
				SqlCommand sqlCommand = new SqlCommand();
				sqlCommand.Connection = sqlConnection;
				sqlConnection.Open();
				int num6 = count;
				foreach (DataRow row in tbStt_rec.Rows)
				{
					num6--;
					sqlCommand.CommandText = string.Format("EXEC [dbo].[DeleteVoucher] '{0}', '{1}'", info.Ma_ct.Trim(), row["stt_rec"].ToString().Trim());
					sqlCommand.ExecuteNonQuery();
					StartUp.waiting.Set_Post_Error(num6);
				}
				sqlConnection.Close();
				if (StartUp.waiting != null)
				{
					StartUp.waiting.Close();
				}
				FormBrowse formBrowse = new FormBrowse(StartupBase.SasObj, tb_Post_Error.DefaultView, "ly_do_loi:H=Lý do lỗi:400;ma_qs:H=Mã quyển c.từ  :130;so_ct:H=Số c.từ:130");
				formBrowse.frmBrw.Title = (StartupBase.M_LAN.Equals("V") ? "Các chứng từ bị lỗi khi post" : "The post data errors");
				formBrowse.frmBrw.LanguageID = "SasIeCt_15";
				formBrowse.ShowDialog();
			}
			else if (info.Xy_ly.Trim() == "1")
			{
				StartUp.waiting.pgValue = tbDuLieuTrung.Rows.Count;
				SqlConnection sqlConnection2 = new SqlConnection(GetConnectionString(500000));
				SqlCommand sqlCommand2 = new SqlCommand();
				sqlCommand2.Connection = sqlConnection2;
				sqlConnection2.Open();
				int num7 = 0;
				foreach (DataRow row2 in tbDuLieuTrung.Rows)
				{
					num7++;
					sqlCommand2.CommandText = string.Format("EXEC [dbo].[DeleteVoucher] '{0}', '{1}'", info.Ma_ct.Trim(), row2["stt_rec"].ToString().Trim());
					sqlCommand2.ExecuteNonQuery();
					StartUp.waiting.Set_Delete(num7);
				}
				sqlConnection2.Close();
				StartUp.waiting.Set_Delete(StartUp.waiting.pgValue);
			}
			return _flag_post;
		}

		private static bool ShowError(DataSet dsError, string maImex)
		{
			switch (maImex.Trim())
			{
				case "PK1":
					return ShowError_PK1(dsError);
				case "HDA":
					return ShowError_HDA(dsError);
				case "HD1":
					return ShowError_HD1(dsError);
				case "PND":
					return ShowError_PND(dsError);
				case "PXD":
					return ShowError_PXD(dsError);
				case "PN1":
					return ShowError_PN1(dsError);
				case "PN2":
					return ShowError_PN2(dsError);
				case "PNA":
					return ShowError_PNA(dsError);
				case "PC1":
					return ShowError_PC1(dsError);
				case "PT1":
					return ShowError_PT1(dsError);
				case "PNB":
					return ShowError_PNB(dsError);
				case "BC1":
					return ShowError_BC1(dsError);
				case "BN1":
					return ShowError_BN1(dsError);
				case "PXE":
					return ShowError_PXE(dsError);
				case "PXV":
					return ShowError_PXV(dsError);
				case "PKK":
					return ShowError_PKK(dsError);
				case "PXF":
					return ShowError_PXF(dsError);
                case "KSX":
                    return ShowError_KSX(dsError);
                case "KSW":
                    return ShowError_KSW(dsError);
                case "KSF":
                    return ShowError_KSF(dsError);
                case "KSS":
                    return ShowError_KSS(dsError);
                case "KSK":
                    return ShowError_KSK(dsError);
                case "HDM":
                    return ShowError_HDM(dsError);
                default:
					return false;
			}
		}

		private static bool ShowError_PK1(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct_thieu_nh_dk":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct_thieu_nh_dk:H=Số chứng từ", "Danh sách chứng từ chưa khai báo nhóm định khoảng");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct_thieu_nh_dk:H=Voucher no.", "List of vouchers without group arisen on account");
							}
							result = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_i:H=Tài khoản;ten_tk:H=Tên tài khoản", "Danh sách tài khoản không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_i:H=Account;ten_tk2:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of Unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "nh_dk":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày c.từ:D;so_ct:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ;nh_dk:H=Nhóm định khoản;t_ps_no:H=Tổng ps nợ:N0;t_ps_co:H=Tổng ps có:N0", "Tổng phát sinh nợ khác tổng phát sinh có trong 1 nhóm định khoản");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D;so_ct:H=Voucher no.;ma_qs:H=Book no.;nh_dk:H=The record(s);t_ps_no:H=Total Debit arising:N0;t_ps_co:H=Total Credit arising:N0", "Total other liabilities arising out of a total group arisen on account");
							}
							result = false;
							break;
						case "tk_cn":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_cn:H=Tài khoản;ma_kh_i:H=Mã khách hàng", "Danh sách tài khoản công nợ chưa vào mã khách hàng");
							}
							else
							{
								BrowseError(table, "tk_cn:H=Account;ma_kh_i:H=Customer ID", "List of financial liabilities not in the client code");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục khoản mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "Free List of code is not in the list of cost items");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							}
							else
							{
								BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_PT1(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher code", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "tk":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk:H=Account credit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_i:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_i:H=Account debit;ten_tk:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "so_ct_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_i:H=Số chứng từ", "Danh sách số chứng từ chưa gắn mã khách hàng khi thu nhiều khách hàng");
							}
							else
							{
								BrowseError(table, "so_ct_i:H=Voucher code", "List of vouchers without customer codes when collecting many customers");
							}
							result = false;
							break;
						case "ma_kh_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_i:H=Mã khách hàng chi tiết", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_i:H=Customer ID detail", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							}
							else
							{
								BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ma_ku_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ku_i:H=Mã khế ước", "Danh sách mã khế ước không có trong danh mục khế ước");
							}
							else
							{
								BrowseError(table, "ma_ku_i:H=Loan contract code", "List code of loan contract doesn't exist in the list of loan contract");
							}
							result = false;
							break;
						case "ma_hdm_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_hdm_i:H=Số đh mua", "Danh sách số đh mua không có trong danh mục đh mua");
							}
							else
							{
								BrowseError(table, "ma_hdm_i:H=PO no.", "List of PO no. doesn't exist in the list of PO");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_PC1(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_thue_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thue_i:H=Mã t.suất", "Danh sách mã t.suất không có trong danh mục thuế suất ");
							}
							else
							{
								BrowseError(table, "ma_thue_i:H=Tax rate code", "List of tax rate code is not in the list of tax");
							}
							result = false;
							break;
						case "ma_kh_t":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_t:H=Mã khách thuế", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_t:H=Customer ID Tax", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher code", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "tk":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk:H=Account credit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_i:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_i:H=Account debit;ten_tk:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "so_ct_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_i:H=Số chứng từ", "Danh sách số chứng từ chưa gắn mã khách hàng khi chi nhiều khách hàng");
							}
							else
							{
								BrowseError(table, "so_ct_i:H=Voucher code", "List of vouchers without customer code when spending many customers");
							}
							result = false;
							break;
						case "ma_kh_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_i:H=Mã khách hàng chi tiết", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_i:H=Customer ID detail", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							}
							else
							{
								BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ma_ku_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ku_i:H=Mã khế ước", "Danh sách mã khế ước không có trong danh mục khế ước");
							}
							else
							{
								BrowseError(table, "ma_ku_i:H=Loan contract code", "List code of loan contract doesn't exist in the list of loan contract");
							}
							result = false;
							break;
						case "ma_hdm_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_hdm_i:H=Số đh mua", "Danh sách số đh mua không có trong danh mục đh mua");
							}
							else
							{
								BrowseError(table, "ma_hdm_i:H=PO no.", "List of PO no. doesn't exist in the list of PO");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_BC1(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							flag = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							flag = false;
							break;
						case "tk":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_no:H=Account debit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							flag = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_co:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_co:H=Account credit;ten_tk:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							flag = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							flag = false;
							break;
						case "ma_kh_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_i:H=Mã khách hàng chi tiết", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_i:H=Customer ID detail", "List the client code is not in the list of customers");
							}
							flag = false;
							break;
						case "so_ct_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_i:H=Số chứng từ", "Danh sách số chứng từ chưa gắn mã khách hàng khi chi nhiều khách hàng");
							}
							else
							{
								BrowseError(table, "so_ct_i:H=Voucher code", "List of vouchers without customer code when spending many customers");
							}
							flag = false;
							break;
						case "ma_ku_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Mã khế ước", "Danh sách mã khế ước không có trong danh mục khế ước");
							else
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Loan contract code", "List code of loan contract doesn't exist in the list of loan contract");
							flag = false;
							break;
						case "ma_hd_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=Số đh bán", "Danh sách số đh bán không có trong danh mục đh bán");
							else
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=SO no.", "List of SO no. doesn't exist in the list of SO");
							flag = false;
							break;
						case "ma_hdm_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=Số đh mua", "Danh sách số đh mua không có trong danh mục đh mua");
							else
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=PO no.", "List of PO no. doesn't exist in the list of PO");
							flag = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							else
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							flag = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bpht", "Danh sách mã bpht không có trong danh mục bpht");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							flag = false;
							break;
						case "ma_td4_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã c.tiêu", "Danh sách mã c.tiêu không có trong danh mục c.tiêu");
							else
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Criteria code", "List code of criteria doesn't exist in the list of criteria");
							flag = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							else
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							flag = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=Số lsx", "Danh sách số lsx không có trong danh mục lsx");
							else
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=MO no.", "List of MO no. doesn't exist in the list of MO");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã td 1", "Danh sách mã td 1 không có trong danh mục mã tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H= Free code 1", "List code of free 1 doesn't exist in free field list 1");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã td 2", "Danh sách mã td 2 không có trong danh mục mã tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H= Free code 2", "List code of free 2 doesn't exist in free field list 2");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã td 3", "Danh sách mã td 3 không có trong danh mục mã tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H= Free code 3", "List code of free 3 doesn't exist in free field list 3");
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}
		private static bool ShowError_BN1(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_thue_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thue_i:H=Mã t.suất", "Danh sách mã t.suất không có trong danh mục thuế suất ");
							}
							else
							{
								BrowseError(table, "ma_thue_i:H=Tax rate code", "List of tax rate code is not in the list of tax");
							}
							flag = false;
							break;
						case "ma_kh_t":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_t:H=Mã khách thuế", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_t:H=Customer ID Tax", "List the client code is not in the list of customers");
							}
							flag = false;
							break;
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã ncc", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							flag = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							flag = false;
							break;
						case "tk":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk:H=Account credit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							flag = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_i:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_i:H=Account debit;ten_tk:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							flag = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							flag = false;
							break;
						case "ma_kh_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_i:H=Mã khách hàng chi tiết", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh_i:H=Customer ID detail", "List the client code is not in the list of customers");
							}
							flag = false;
							break;
						case "so_ct_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_i:H=Số chứng từ", "Danh sách số chứng từ chưa gắn mã khách hàng khi chi nhiều khách hàng");
							}
							else
							{
								BrowseError(table, "so_ct_i:H=Voucher code", "List of vouchers without customer code when spending many customers");
							}
							flag = false;
							break;
						case "ma_ku_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Mã khế ước", "Danh sách mã khế ước không có trong danh mục khế ước");
							else
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Loan contract code", "List code of loan contract doesn't exist in the list of loan contract");
							flag = false;
							break;
						case "ma_hd_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=Số đh bán", "Danh sách số đh bán không có trong danh mục đh bán");
							else
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=SO no.", "List of SO no. doesn't exist in the list of SO");
							flag = false;
							break;
						case "ma_hdm_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=Số đh mua", "Danh sách số đh mua không có trong danh mục đh mua");
							else
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=PO no.", "List of PO no. doesn't exist in the list of PO");
							flag = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							else
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							flag = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bpht", "Danh sách mã bpht không có trong danh mục bpht");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							flag = false;
							break;
						case "ma_td4_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã c.tiêu", "Danh sách mã c.tiêu không có trong danh mục c.tiêu");
							else
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Criteria code", "List code of criteria doesn't exist in the list of criteria");
							flag = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							else
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							flag = false;
							continue;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=Số lsx", "Danh sách số lsx không có trong danh mục lsx");
							else
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=MO no.", "List of MO no. doesn't exist in the list of MO");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã td 1", "Danh sách mã td 1 không có trong danh mục mã tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H= Free code 1", "List code of free 1 doesn't exist in free field list 1");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã td 2", "Danh sách mã td 2 không có trong danh mục mã tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H= Free code 2", "List code of free 2 doesn't exist in free field list 2");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã td 3", "Danh sách mã td 3 không có trong danh mục mã tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H= Free code 3", "List code of free 3 doesn't exist in free field list 3");
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}
		private static bool ShowError_HDA(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh_hddt_tt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Mã khách hàng", "Trạng thái khách hàng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Customer ID", "Invoice customer status is invalid");
							}
							result = false;
							break;
						case "sd_hddt_yn":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn:H=Sử dụng hóa đơn điện tử", "Chưa khai báo sử dụng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn:H=Use electronic invoices", "Not declaring the use of electronic invoices");
							}
							result = false;
							break;
						case "sd_hddt_yn_2":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn_2:H=Sử dụng hóa đơn điện tử", "Giá trị khai báo sử dụng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn_2:H=Use electronic invoices", "The declared value uses an invalid electronic invoice");
							}
							result = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kho:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							}
							else
							{
								BrowseError(table, "ma_kho:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the unit code repository of documents");
							}
							result = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H=Mã nhập xuất (Tk nợ);ten_tk:H=Tên nhập xuất", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct_hddt:H=Số hđ", "Danh sách những chứng từ đã phát hành hóa đơn điện tử không sao chép được");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Voucher book;so_ct_hddt:H=Voucher no.", "List invoice no. released no copy");
							}
							result = false;
							break;
						case "tk_dt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_dt:H=Tk dt;ten_tk:H=Tên tk", "Danh sách tài khoản doanh thu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_dt:H=Revenue acct.;ten_tk:H=Acct. name", "List revenue account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_vt:H=Tk vt;ten_tk:H=Tên tk", "Danh sách tài khoản vật tư không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_vt:H=Item account;ten_tk:H=Account name", "List of materials no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_gv":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_gv:H=Tk gv;ten_tk:H=Tên tk", "Danh sách tài khoản giá vốn không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_gv:H=COGS account;ten_tk:H=Account name", "List COGS account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_ck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_ck:H=Tk ck;ten_tk:H=Tên tk", "Danh sách tài khoản chiết khấu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_ck:H=Discount acct.;ten_tk:H=Acct. name", "List of discount account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "tk_thue_co":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_thue_co:H=Tk thuế;ten_tk:H=Tên tk", "Danh sách tài khoản thuế không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_thue_co:H=Tax account;ten_tk:H=Acct. name", "List of tax account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_km_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_km_i:H=Tài khoản cp km;ten_tk:H=Tên tài khoản", "Danh sách tài khoản cp km không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_km_i:H=Account promotion expenses;ten_tk:H=Account name", "List promotion expenses accounts not in list accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
        private static bool ShowError_HDM(DataSet dsError)
        {
            bool result = true;
            foreach (DataTable table in dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            }
                            else
                            {
                                BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
                            }
                            result = false;
                            break;
                      
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            }
                            else
                            {
                                BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
                            }
                            result = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
                            }
                            result = false;
                            break;
                  
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            result = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
                            }
                            else
                            {
                                BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
                            }
                            result = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            }
                            else
                            {
                                BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
                            }
                            result = false;
                            break;
                        case "so_ct_hddt":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct_hddt:H=Số hđ", "Danh sách những chứng từ đã phát hành hóa đơn điện tử không sao chép được");
                            }
                            else
                            {
                                BrowseError(table, "ma_qs:H=Voucher book;so_ct_hddt:H=Voucher no.", "List invoice no. released no copy");
                            }
                            result = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
                            }
                            else
                            {
                                BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
                            }
                            result = false;
                            break;
                    }
                }
            }
            return result;
        }
        private static bool ShowError_HD1(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh1:H=Mã khách hàng;ma_dvcs:H=Mã ĐVCS", "Danh sách mã khách hàng không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh1:H=Customer ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							}
							result = false;
							break;
						case "ma_kh_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh_hddt_tt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Mã khách hàng", "Trạng thái khách hàng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Customer ID", "Invoice customer status is invalid");
							}
							result = false;
							break;
						case "sd_hddt_yn":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn:H=Sử dụng hóa đơn điện tử", "Chưa khai báo sử dụng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn:H=Use electronic invoices", "Not declaring the use of electronic invoices");
							}
							result = false;
							break;
						case "sd_hddt_yn_2":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn_2:H=Sử dụng hóa đơn điện tử", "Giá trị khai báo sử dụng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn_2:H=Use electronic invoices", "The declared value uses an invalid electronic invoice");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Debit account;ten_tk:H=Debit account name", "List of debit account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_thck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thck:H=Mã điều khoản thanh toán", "Danh sách mã điều khoản thanh toán không có trong danh mục điểu khoản thanh toán");
							}
							else
							{
								BrowseError(table, "ma_thck:H=Payment terms code", "List of payment terms code not in payment terms list");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct_hddt:H=Số hđ", "Danh sách những chứng từ đã phát hành hóa đơn điện tử không sao chép được");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Voucher book;so_ct_hddt:H=Voucher no.", "List invoice no. released no copy");
							}
							result = false;
							break;
						case "tk_dt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_dt:H=Tk dt;ten_tk:H=Tên tk", "Danh sách tài khoản doanh thu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_dt:H=Revenue acct.;ten_tk:H=Acct. name", "List revenue account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_ck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_ck:H=Tk ck;ten_tk:H=Tên tk", "Danh sách tài khoản chiết khấu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_ck:H=Discount acct.;ten_tk:H=Acct. name", "List of discount account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "tk_thue_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_thue_i:H=Tk thuế;ten_tk:H=Tên tk", "Danh sách tài khoản thuế không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_thue_i:H=Tax account;ten_tk:H=Acct. name", "List of tax account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã td 1", "Danh sách mã td 1 không có trong danh mục mã tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H= Free code 1", "List of free code 1 is not valid");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã td 2", "Danh sách mã td 2 không có trong danh mục mã tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H= Free code 2", "List of free code 2 is not valid");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã td 3", "Danh sách mã td 3 không có trong danh mục mã tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H= Free code 3", "List of free code 3 is not valid");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bpht", "Danh sách mã bpht không có trong danh mục bpht");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=Dept. ID", "List of Dept. ID is not valid");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_PNA(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã ncc", "Danh sách mã ncc không có trong danh mục ncc");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Supplier ID", "List the client code is not in the list of Suppliers");
							}
							result = false;
							break;
						case "ma_kh1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh1:H=Mã ncc;ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh1:H=Supplier ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							}
							result = false;
							break;
						case "ma_kh_thue":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_thue:H=Mã ncc(Hđ thuế);ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không có trong danh mục ncc hoặc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh_thue:H=Supplier ID (tax);ma_dvcs:H=Unit code", "List the client code is not in the list of Suppliers or not in Unit");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H= Mã nx(Tk có);ten_tk:H=Tên mã nx (tk có)", "Danh sách mã nx(tk có) không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_ms":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ms:H=Nhóm hđ", "Danh sách mã nhóm hóa đơn không có trong danh mục nhóm hđ");
							}
							else
							{
								BrowseError(table, "ma_ms:H=Invoice group code", "List of invoice group code is not valid");
							}
							result = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kho:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							}
							else
							{
								BrowseError(table, "ma_kho:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the unit code repository of documents");
							}
							result = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							}
							result = false;
							break;
						case "ma_nt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nt:H=Mã n.tệ", "Danh sách mã n.tệ không có trong danh mục tiền tệ");
							}
							else
							{
								BrowseError(table, "ma_nt:H=Currency code", "List of currency code is not in the list of currency");
							}
							result = false;
							break;
						case "ma_thue":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thue:H=Mã t.suất", "Danh sách mã t.suất không có trong danh mục thuế suất ");
							}
							else
							{
								BrowseError(table, "ma_thue:H=Tax rate code", "List of tax rate code is not in the list of tax");
							}
							result = false;
							break;
						case "loai_pb":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "loai_pb:H=Kiểu phân bổ", "Danh sách kiểu phân bổ không hợp lệ");
							}
							else
							{
								BrowseError(table, "loai_pb:H=Allocate type", "List of allocate type is not valid");
							}
							result = false;
							break;
						case "ma_thck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thck:H=Mã đ/k tt", "Danh sách mã điều khoản thanh toán không có trong danh mục điểu khoản thanh toán");
							}
							else
							{
								BrowseError(table, "ma_thck:H=Payment terms code", "List of payment terms code not in payment terms list");
							}
							result = false;
							break;
						case "tk_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_vt:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_vt:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H= Số c.từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct0":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct0:H=Số hđ", "Danh sách số hóa đơn trùng số");
							}
							else
							{
								BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "stt_hd_thue":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct0:H=Số hđ", "Danh sách số hóa đơn trùng số");
							}
							else
							{
								BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_PN1(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã ncc", "Danh sách mã ncc không có trong danh mục ncc");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Supplier ID", "List the client code is not in the list of Suppliers");
							}
							result = false;
							break;
						case "ma_kh1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh1:H=Mã ncc;ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh1:H=Supplier ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							}
							result = false;
							break;
						case "ma_kh_thue":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_thue:H=Mã ncc(HĐ Thuế);ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không có trong danh mục ncc hoặc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh_thue:H=Supplier ID(Tax);ma_dvcs:H=Unit code", "List the client code is not in the list of Suppliers or not in Unit");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H=Tk có;ten_tk:H=Tên tk có ", "Danh sách tk có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_ms":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ms:H=Nhóm hđ", "Danh sách nhóm hđ không có trong danh mục nhóm hóa đơn");
							}
							else
							{
								BrowseError(table, "ma_ms:H=Invoice group code", "List of invoice group code is not valid");
							}
							result = false;
							break;
						case "tk_thue_no":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_thue_no:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_thue_no:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_thck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_thck:H=Mã điều khoản thanh toán", "Danh sách mã điều khoản thanh toán không có trong danh mục điểu khoản thanh toán");
							}
							else
							{
								BrowseError(table, "ma_thck:H=Payment terms code", "List of payment terms code not in payment terms list");
							}
							result = false;
							break;
						case "tk_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_vt:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_vt:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct0":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct0:H=Số hóa đơn", "Danh sách số hóa đơn trùng số");
							}
							else
							{
								BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}

		private static bool ShowError_PN2(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã ncc", "Danh sách mã ncc không có trong danh mục ncc");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Supplier ID", "List the client code is not in the list of Suppliers");
							}
							result = false;
							break;
						case "ma_kh1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh1:H=Mã ncc;ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh1:H=Supplier ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							}
							result = false;
							break;
						case "ma_kh_thue":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_thue:H=Mã ncc(HĐ Thuế);ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không có trong danh mục ncc hoặc không thuộc ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_kh_thue:H=Supplier ID(Tax);ma_dvcs:H=Unit code", "List the client code is not in the list of Suppliers or not in Unit");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H=Tk có;ten_tk:H=Tên tk có ", "Danh sách tk có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_ms":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ms:H=Nhóm hđ", "Danh sách nhóm hđ không có trong danh mục nhóm hóa đơn");
							}
							else
							{
								BrowseError(table, "ma_ms:H=Invoice group code", "List of invoice group code is not valid");
							}
							result = false;
							break;
						case "tk_thue_no":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_thue_no:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_thue_no:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_vt:H=Tk nợ;ten_tk:H=Tên tk nợ ", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_vt:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct0":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct0:H=Số hóa đơn", "Danh sách số hóa đơn trùng số");
							}
							else
							{
								BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}

		private static bool ShowError_PNB(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã ncc", "Danh sách mã ncc không có trong danh mục ncc");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Supplier ID", "List the client code is not in the list of Suppliers");
							flag = false;
							break;
						case "ma_kh_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh_i:H=Mã ncc chi phí", "Danh sách mã ncc chi phí không có trong danh mục ncc");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh_i:H=Supplier ID", "List the client code is not in the list of Suppliers");
							flag = false;
							break;
						case "ma_kh_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh_thue:H=Mã ncc(Hđ thuế)", "Danh sách mã ncc không có trong danh mục ncc");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh_thue:H=Supplier ID (tax)", "List the client code is not in the list of Suppliers");
							flag = false;
							break;
						case "ma_kh1":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh1:H=Mã ncc;ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh1:H=Supplier ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							flag = false;
							break;
						case "ma_kh_i1":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh_i1:H=Mã ncc;ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh_i1:H=Supplier ID;ma_dvcs:H=Unit code", "List the client code is invalid");
							flag = false;
							break;
						case "ma_kh_thue1":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh_thue1:H=Mã ncc(Hđ thuế);ma_dvcs:H=Mã ĐVCS", "Danh sách mã ncc không thuộc ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh_thue1:H=Supplier ID (Tax);ma_dvcs:H=Unit code", "List the client code is not in list Unit");
							flag = false;
							break;
						case "tk_thue_nk":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_thue_nk:H=Tk thuế NK;ten_tk:H=Tên tk thuế NK", "Danh sách tk thuế NK không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_thue_nk:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "tk_thue_db":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_thue_db:H=Tk thuế TTĐB;ten_tk:H=Tên tk thuế TTĐB", "Danh sách tk thuế TTĐB không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_thue_db:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_nx:H= Mã nx(Tk có);ten_tk:H=Tên mã nx (tk có)", "Danh sách mã nx(tk có) không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "ma_nx:H=Dr./Cr. account (Account credit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "tk_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_i:H=Tk có(Hđ thuế);ten_tk:H=Tên tk có", "Danh sách tk có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_i:H=Credit account(Tax);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "tk_thue_no_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_thue_no_thue:H=Tk thuế(Hđ thuế);ten_tk:H=Tên tk thuế", "Danh sách tk thuế không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_thue_no_thue:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "tk_du_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_du_thue:H=Tk đ.ứng(Hđ thuế);ten_tk:H=Tên tk thuế ", "Danh sách tk đ.ứng thuế không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_du_thue:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							flag = false;
							break;
						case "ma_ms":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_ms:H=Nhóm hđ", "Danh sách mã nhóm hóa đơn không có trong danh mục nhóm hđ");
							else
								C_ImportVoucher.BrowseError(table, "ma_ms:H=Invoice group code", "List of invoice group code is not valid");
							flag = false;
							break;
						case "ma_kho_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kho_i:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_kho_i:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the unit code repository of documents");
							flag = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							flag = false;
							break;
						case "ma_vt_th1":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt_th1:H=Mã vật tư;tk_vt:H=Tk nợ(Tk vật tư);tk_vt_dmvt:H=Tk kho trong danh mục vật tư:250", "Tk nợ(Tk vật tư) không đúng với tk kho trong danh mục vật tư");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt_th1:H=Item code;tk_vt:H=Dr./Cr. account (Account debit);tk_vt_dmvt:H=Tk kho trong danh mục vật tư:250", "List of supplies code is not in the list of materials");
							flag = false;
							break;
						case "ma_vt_th2":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt_th2:H=Mã vật tư;tk_vt:H=Tk nợ(Tk vật tư)", "Tk nợ(Tk vật tư) không có trong danh mục tài khoản");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt_th2:H=Item code;tk_vt:H=Dr./Cr. account (Account debit)", "List of supplies code is not in the list of materials");
							flag = false;
							break;
						case "ma_nt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_nt:H=Mã n.tệ", "Danh sách mã n.tệ không có trong danh mục tiền tệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_nt:H=Currency code", "List of currency code is not in the list of currency");
							flag = false;
							break;
						case "ma_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_thue:H=Mã t.suất", "Danh sách mã t.suất không có trong danh mục thuế suất ");
							else
								C_ImportVoucher.BrowseError(table, "ma_thue:H=Tax rate code", "List of tax rate code is not in the list of tax");
							flag = false;
							break;
						case "ma_thue_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_thue_thue:H=Mã t.suất (Hđ thuế)", "Danh sách mã t.suất không có trong danh mục thuế suất");
							else
								C_ImportVoucher.BrowseError(table, "ma_thue_thue:H=Tax rate code (Tax invoices)", "List of tax rate code is not in the list of tax");
							flag = false;
							break;
						case "loai_pb":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "loai_pb:H=Kiểu phân bổ", "Danh sách kiểu phân bổ không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "loai_pb:H=Allocate type", "List of allocate type is not valid");
							flag = false;
							break;
						case "ma_thck":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_thck:H=Mã đ/k tt", "Danh sách mã điều khoản thanh toán không có trong danh mục điểu khoản thanh toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_thck:H=Payment terms code", "List of payment terms code not in payment terms list");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H= Số c.từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							flag = false;
							break;
						case "so_ct0":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct0:H=Số hđ", "Danh sách số hóa đơn trùng số");
							else
								C_ImportVoucher.BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							flag = false;
							break;
						case "stt_hd_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct0:H=Số hđ", "Danh sách số hóa đơn trùng số");
							else
								C_ImportVoucher.BrowseError(table, "so_ct0:H=Voucher no.", "List of vouchers with some");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							flag = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							flag = false;
							break;
						case "ma_vv":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv:H=Project code", "List of project code is not in the list of projects");
							flag = false;
							break;
						case "ma_vv_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_thue:H=Mã dự án (Hđ thuế)", "Danh sách mã dự án (Hđ thuế) không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_thue:H=Project code", "List of project code is not in the list of projects");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							flag = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							else
								C_ImportVoucher.BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							flag = false;
							break;
						case "ma_phi":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_phi:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							else
								C_ImportVoucher.BrowseError(table, "ma_phi:H=Fee code", "List of fee code is not in the list of fee");
							flag = false;
							break;
						case "ma_phi_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_phi_thue:H=Mã phí (Hđ thuế)", "Danh sách mã phí (Hđ thuế) không có trong danh mục phí");
							else
								C_ImportVoucher.BrowseError(table, "ma_phi_thue:H=Fee code", "List of fee code is not in the list of fee");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							flag = false;
							break;
						case "dvt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư;dvt:H=Đvt", "Danh sách mã vật tư có đvt không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Item code;dvt:H=UOM", "List of UOM is not in the list of materials");
							flag = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							else
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							flag = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							flag = false;
							break;
						case "ma_bpht":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							flag = false;
							break;
						case "ma_bpht_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_thue:H=Mã bộ phận hạch toán (Hđ thuế)", "Danh sách mã bộ phận hạch toán (Hđ thuế) không có trong danh mục bộ phận hạch toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_thue:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							flag = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							else
								C_ImportVoucher.BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							flag = false;
							break;
						case "so_lsx":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_lsx:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							else
								C_ImportVoucher.BrowseError(table, "so_lsx:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							flag = false;
							break;
						case "so_lsx_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_lsx_thue:H=Số LSX (Hđ thuế)", "Danh sách mã lệnh sản xuất (Hđ thuế) không có trong danh mục lệnh sản xuất");
							else
								C_ImportVoucher.BrowseError(table, "so_lsx_thue:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							flag = false;
							break;
						case "ma_ku_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Mã khế ước", "Danh sách mã khế ước không có trong danh mục khế ước");
							else
								C_ImportVoucher.BrowseError(table, "ma_ku_i:H=Loan contract code", "List code of loan contract doesn't exist in the list of loan contract");
							flag = false;
							break;
						case "ma_hd_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=Số đh bán", "Danh sách số đh bán không có trong danh mục đh bán");
							else
								C_ImportVoucher.BrowseError(table, "ma_hd_i:H=SO no.", "List of SO no. doesn't exist in the list of SO");
							flag = false;
							break;
						case "ma_hdm_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=Số đh mua", "Danh sách số đh mua không có trong danh mục đh mua");
							else
								C_ImportVoucher.BrowseError(table, "ma_hdm_i:H=PO no.", "List of PO no. doesn't exist in the list of PO");
							flag = false;
							break;
						case "ma_td4_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã c.tiêu", "Danh sách mã c.tiêu không có trong danh mục c.tiêu");
							else
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Criteria code", "List code of criteria doesn't exist in the list of criteria");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							flag = false;
							break;
						case "ma_td":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							flag = false;
							break;
						case "ma_td_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_thue:H=Mã tự do 1 (Hđ thuế)", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_thue:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							flag = false;
							break;
						case "ma_td2":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							flag = false;
							break;
						case "ma_td2_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_thue:H=Mã tự do 2 (Hđ thuế)", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_thue:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							flag = false;
							break;
						case "ma_td3":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							flag = false;
							break;
						case "ma_td3_thue":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_thue:H=Mã tự do 3 (Hđ thuế)", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_thue:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							flag = false;
							break;
						case "so_ct_khac_cp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_cp:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có chi phí phân bổ khác tổng chi phí");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_cp:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							flag = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}
		private static bool ShowError_PND(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kho:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							}
							else
							{
								BrowseError(table, "ma_kho:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the code DVCS repository of documents");
							}
							result = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "tk_no":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_no:H=Account debit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_co":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_co:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_co:H=Account credit;ten_tk:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							}
							else
							{
								BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}
		private static bool ShowError_PXD(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kho:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							}
							else
							{
								BrowseError(table, "ma_kho:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the code DVCS repository of documents");
							}
							result = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ", "Danh sách quyển c.từ  không có trong danh mục quyển c.từ  hoặc chứng từ không thuộc quyển c.từ  này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher code", "List of vouchers with some");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "tk_no":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_no:H=Account debit;ten_tk:H=Account name", "List creditors account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_co":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_co:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_co:H=Account credit;ten_tk2:H=Account name", "List of no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							}
							else
							{
								BrowseError(table, "ma_sp:H=Product code", "List product code is not in the list of products");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=Dept ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							result = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}

		private static bool ShowError_PXE(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
							flag = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							flag = false;
							break;
						case "ma_khon":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_khon:H=Mã kho nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho nhập không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_khon:H=Mã kho nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho nhập không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							flag = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
							flag = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
							flag = false;
							break;
						case "tk_no":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							flag = false;
							break;
						case "tk_co":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_co:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_co:H=Tài khoản có e;ten_tk2:H=Tên tài khoản e", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp e");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							flag = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							flag = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							else
								C_ImportVoucher.BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							flag = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							else
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							flag = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							flag = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							flag = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}

		private static bool ShowError_PXV(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
							flag = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho tồn kho theo dự án hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho tồn kho theo dự án hoặc không thuộc mã ĐVCS của chứng từ");
							flag = false;
							break;
						case "ma_khon":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_khon:H=Mã kho nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho nhập không có trong danh mục kho tồn kho theo dự án hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_khon:H=Mã kho nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho nhập không có trong danh mục kho tồn kho theo dự án hoặc không thuộc mã ĐVCS của chứng từ");
							flag = false;
							break;
						case "ma_vv":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv:H=Mã dự án xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã dự án xuất không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv:H=Mã dự án xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã dự án xuất không có trong danh mục dự án e");
							flag = false;
							break;
						case "ma_vv_n":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vv_n:H=Mã dự án nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã dự án nhập không có trong danh mục dự án");
							else
								C_ImportVoucher.BrowseError(table, "ma_vv_n:H=Mã dự án nhập;ma_dvcs:H=Mã ĐVCS", "Danh sách mã dự án nhập không có trong danh mục dự án e");
							flag = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
							flag = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
							flag = false;
							break;
						case "tk_no":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_no:H=Tài khoản nợ;ten_tk:H=Tên tài khoản", "Danh sách tài khoản nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp e");
							flag = false;
							break;
						case "tk_co":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "tk_co:H=Tài khoản có;ten_tk:H=Tên tài khoản", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							else
								C_ImportVoucher.BrowseError(table, "tk_co:H=Tài khoản có e;ten_tk2:H=Tên tài khoản e", "Danh sách tài khoản có không có trong danh mục tài khoản hoặc là tài khoản tổng hợp e");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
							flag = false;
							break;
						case "ma_td4_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
							else
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							flag = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							else
								C_ImportVoucher.BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							flag = false;
							break;
						case "ma_sp":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							else
								C_ImportVoucher.BrowseError(table, "ma_sp:H=Mã sản phẩm", "Danh sách mã sản phẩm không có trong danh mục sản phẩm");
							flag = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							else
								C_ImportVoucher.BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							flag = false;
							break;
						case "so_lsx_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_lsx_i:H=Số LSX", "Danh sách mã lệnh sản xuất không có trong danh mục lệnh sản xuất");
							}
							else
							{
								BrowseError(table, "so_lsx_i:H=Manufacturing order", "List of manufacturing order is not in the list of manufacturing order entry");
							}
							flag = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}

		private static bool ShowError_PKK(DataSet dsError)
		{
			bool flag = true;
			foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							else
								C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
							flag = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							else
								C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							flag = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							else
								C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
							flag = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
							flag = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							else
								C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
							flag = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							else
								C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
							flag = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							else
								C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
							flag = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							else
								C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
							flag = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							else
								C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
							flag = false;
							break;
						case "ma_td4_i":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
							else
								C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
							flag = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							else
								C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
							flag = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							else
								C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							flag = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							else
								C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
							flag = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							flag = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							flag = false;
							break;
						default:
							break;
					}
				}
			}
			return flag;
		}
        private static bool ShowError_KSX(DataSet dsError)
        {
            bool flag = true;
            foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
                            flag = false;
                            break;
                        case "ma_sp_sp":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_kho":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            flag = false;
                            break;
                        case "ma_vt":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
                            flag = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
                            flag = false;
                            break;
                        case "ma_gd":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
                            flag = false;
                            break;
                        case "ma_td_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
                            flag = false;
                            break;
                        case "ma_td2_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
                            flag = false;
                            break;
                        case "ma_td3_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
                            flag = false;
                            break;
                        case "ma_td4_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
                            flag = false;
                            break;
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
                            flag = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            else
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            flag = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            else
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            flag = false;
                            break;
                        case "ngay_ct_ks":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
                            }
                            flag = false;
                            break;
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            flag = false;
                            break;
                        default:
                            break;
                    }
                }
            }
            return flag;
        }
        private static bool ShowError_KSS(DataSet dsError)
        {
            bool flag = true;
            foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
                            flag = false;
                            break;
                        case "ma_sp_sp":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_kho":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            flag = false;
                            break;
                        case "ma_vt":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
                            flag = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
                            flag = false;
                            break;
                        case "ma_gd":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
                            flag = false;
                            break;
                        case "ma_td_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
                            flag = false;
                            break;
                        case "ma_td2_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
                            flag = false;
                            break;
                        case "ma_td3_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
                            flag = false;
                            break;
                        case "ma_td4_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
                            flag = false;
                            break;
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
                            flag = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            else
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            flag = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            else
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            flag = false;
                            break;
                        case "ngay_ct_ks":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
                            }
                            flag = false;
                            break;
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            flag = false;
                            break;
                        default:
                            break;
                    }
                }
            }
            return flag;
        }
        private static bool ShowError_KSF(DataSet dsError)
        {
            bool flag = true;
            foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
                            flag = false;
                            break;
                        case "ma_sp_sp":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_kho":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            flag = false;
                            break;
                        case "ma_vt":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
                            flag = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
                            flag = false;
                            break;
                        case "ma_gd":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
                            flag = false;
                            break;
                        case "ma_td_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
                            flag = false;
                            break;
                        case "ma_td2_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
                            flag = false;
                            break;
                        case "ma_td3_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
                            flag = false;
                            break;
                        case "ma_td4_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
                            flag = false;
                            break;
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
                            flag = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            else
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            flag = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            else
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            flag = false;
                            break;
                        case "ngay_ct_ks":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
                            }
                            flag = false;
                            break;
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            flag = false;
                            break;
                        default:
                            break;
                    }
                }
            }
            return flag;
        }
        private static bool ShowError_KSW(DataSet dsError)
        {
            bool flag = true;
            foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
                            flag = false;
                            break;
                        case "ma_sp_sp":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_kho":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            flag = false;
                            break;
                        case "ma_vt":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
                            flag = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
                            flag = false;
                            break;
                        case "ma_gd":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
                            flag = false;
                            break;
                        case "ma_td_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
                            flag = false;
                            break;
                        case "ma_td2_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
                            flag = false;
                            break;
                        case "ma_td3_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
                            flag = false;
                            break;
                        case "ma_td4_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
                            flag = false;
                            break;
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
                            flag = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            else
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            flag = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            else
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            flag = false;
                            break;
                        case "ngay_ct_ks":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
                            }
                            flag = false;
                            break;
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            flag = false;
                            break;
                        default:
                            break;
                    }
                }
            }
            return flag;
        }
        private static bool ShowError_KSK(DataSet dsError)
        {
            bool flag = true;
            foreach (DataTable table in (InternalDataCollectionBase)dsError.Tables)
            {
                if (table.Rows.Count != 0)
                {
                    switch (table.Columns[0].ColumnName)
                    {
                        case "ma_kh":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng e");
                            flag = false;
                            break;
                        case "ma_kho":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_kho:H=Mã kho xuất;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho xuất không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
                            flag = false;
                            break;
                        case "ma_sp_sp":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_vt":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_vt:H=Mã vật tư e", "Danh sách mã vật tư không có trong danh mục vật tư e");
                            flag = false;
                            break;
                        case "ma_qs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e", "Danh sách quyển sổ không có trong danh mục quyển sổ hoặc chứng từ không thuộc quyển sổ này e");
                            flag = false;
                            break;
                        case "so_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_qs:H=Quyển sổ e;so_ct:H=Số chứng từ e", "Danh sách số chứng từ trùng số e");
                            flag = false;
                            break;
                        case "ma_gd":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_gd:H=Mã giao dịch e", "Danh sách mã giao dịch không hợp lệ e");
                            flag = false;
                            break;
                        case "ma_td_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1 e");
                            flag = false;
                            break;
                        case "ma_td2_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2 e");
                            flag = false;
                            break;
                        case "ma_td3_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3 e");
                            flag = false;
                            break;
                        case "ma_td4_i":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_td4_i:H=Mã tự do 4", "Danh sách mã tự do 4 không có trong danh mục tự do 4 e");
                            flag = false;
                            break;
                        case "ma_dvcs":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
                            else
                                C_ImportVoucher.BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS e");
                            flag = false;
                            break;
                        case "ngay_ct":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            else
                                C_ImportVoucher.BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
                            flag = false;
                            break;
                        case "so_ct_khac_ngay":
                            if (StartupBase.M_LAN == "V")
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            else
                                C_ImportVoucher.BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển sổ", "Danh sách c.từ có ngày c.từ khác nhau");
                            flag = false;
                            break;
                        case "ngay_ct_ks":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
                            }
                            else
                            {
                                BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
                            }
                            flag = false;
                            break;
                        case "dvt1":
                            if (StartupBase.M_LAN == "V")
                            {
                                BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
                            }
                            else
                            {
                                BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
                            }
                            flag = false;
                            break;
                        default:
                            break;
                    }
                }
            }
            return flag;
        }
        private static bool ShowError_PXF(DataSet dsError)
		{
			bool result = true;
			foreach (DataTable table in dsError.Tables)
			{
				if (table.Rows.Count != 0)
				{
					switch (table.Columns[0].ColumnName)
					{
						case "ma_kh":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng");
							}
							else
							{
								BrowseError(table, "ma_kh:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt:H=Mã khách hàng", "Danh sách mã khách hàng không có trong danh mục khách hàng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt:H=Customer ID", "List the client code is not in the list of customers");
							}
							result = false;
							break;
						case "ma_kh_hddt_tt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Mã khách hàng", "Trạng thái khách hàng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_kh_hddt_tt:H=Customer ID", "Invoice customer status is invalid");
							}
							result = false;
							break;
						case "sd_hddt_yn":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn:H=Sử dụng hóa đơn điện tử", "Chưa khai báo sử dụng hóa đơn điện tử");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn:H=Use electronic invoices", "Not declaring the use of electronic invoices");
							}
							result = false;
							break;
						case "sd_hddt_yn_2":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "sd_hddt_yn_2:H=Sử dụng hóa đơn điện tử", "Giá trị khai báo sử dụng hóa đơn điện tử không hợp lệ");
							}
							else
							{
								BrowseError(table, "sd_hddt_yn_2:H=Use electronic invoices", "The declared value uses an invalid electronic invoice");
							}
							result = false;
							break;
						case "ma_kho":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_kho:H=Mã kho;ma_dvcs:H=Mã ĐVCS", "Danh sách mã kho không có trong danh mục kho hoặc không thuộc mã ĐVCS của chứng từ");
							}
							else
							{
								BrowseError(table, "ma_kho:H=Site code;ma_dvcs:H=Unit code", "List of code repositories is not in the list or not on the unit code repository of documents");
							}
							result = false;
							break;
						case "ma_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư", "Danh sách mã vật tư không có trong danh mục vật tư");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code", "List of supplies code is not in the list of materials");
							}
							result = false;
							break;
						case "ma_nx":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_nx:H=Mã nhập xuất (Tk nợ);ten_tk:H=Tên nhập xuất", "Danh sách tk nợ không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "ma_nx:H=Dr./Cr. account (Account debit);ten_tk:H=Dr./Cr. Name", "List of import export code (account debt) is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "ma_qs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ", "Danh sách quyển c.từ không có trong danh mục quyển c.từ hoặc chứng từ không thuộc quyển c.từ này");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.", "Book list is not in the list or book vouchers of this book");
							}
							result = false;
							break;
						case "so_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct:H=Số chứng từ", "Danh sách số chứng từ trùng số");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Book no.;so_ct:H=Voucher no.", "List of vouchers with some");
							}
							result = false;
							break;
						case "so_ct_hddt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_qs:H=Quyển c.từ  ;so_ct_hddt:H=Số hđ", "Danh sách những chứng từ đã phát hành hóa đơn điện tử không sao chép được");
							}
							else
							{
								BrowseError(table, "ma_qs:H=Voucher book;so_ct_hddt:H=Voucher no.", "List invoice no. released no copy");
							}
							result = false;
							break;
						case "tk_dt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_dt:H=Tk dt;ten_tk:H=Tên tk", "Danh sách tài khoản doanh thu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_dt:H=Revenue acct.;ten_tk:H=Acct. name", "List revenue account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_vt":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_vt:H=Tk vt;ten_tk:H=Tên tk", "Danh sách tài khoản vật tư không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_vt:H=Item account;ten_tk:H=Account name", "List of materials no account in the chart of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "tk_gv":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_gv:H=Tk gv;ten_tk:H=Tên tk", "Danh sách tài khoản giá vốn không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_gv:H=COGS account;ten_tk:H=Account name", "List COGS account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_ck":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_ck:H=Tk ck;ten_tk:H=Tên tk", "Danh sách tài khoản chiết khấu không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_ck:H=Discount acct.;ten_tk:H=Acct. name", "List of discount account is not in the list of accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_gd":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_gd:H=Mã giao dịch", "Danh sách mã giao dịch không hợp lệ");
							}
							else
							{
								BrowseError(table, "ma_gd:H=Transaction code", "List an invalid transaction code");
							}
							result = false;
							break;
						case "ma_dvcs":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_dvcs:H=Mã ĐVCS", "Danh sách mã ĐVCS không có trong danh mục ĐVCS");
							}
							else
							{
								BrowseError(table, "ma_dvcs:H=Unit code", "Unit code list is not in the list of unit code");
							}
							result = false;
							break;
						case "ngay_ct":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày mở sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct:H=Voucher date:D", "List of documents smaller day open day window");
							}
							result = false;
							break;
						case "tk_thue_co":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_thue_co:H=Tk thuế;ten_tk:H=Tên tk", "Danh sách tài khoản thuế không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_thue_co:H=Tax account;ten_tk:H=Acct. name", "List of tax account is not in the list of accounts or the consolidated accounts");
							}
							result = false;
							break;
						case "tk_km_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "tk_km_i:H=Tài khoản cp km;ten_tk:H=Tên tài khoản", "Danh sách tài khoản cp km không có trong danh mục tài khoản hoặc là tài khoản tổng hợp");
							}
							else
							{
								BrowseError(table, "tk_km_i:H=Account promotion expenses;ten_tk:H=Account name", "List promotion expenses accounts not in list accounts or consolidated accounts");
							}
							result = false;
							break;
						case "ma_vv_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vv_i:H=Mã dự án", "Danh sách mã dự án không có trong danh mục dự án");
							}
							else
							{
								BrowseError(table, "ma_vv_i:H=Project code", "List of project code is not in the list of projects");
							}
							result = false;
							break;
						case "so_ct_khac_ngay":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "so_ct_khac_ngay:H=Số c.từ;ma_qs:H=Mã quyển c.từ  ", "Danh sách c.từ có ngày c.từ khác nhau");
							}
							else
							{
								BrowseError(table, "so_ct_khac_ngay:H=Voucher no.;ma_qs:H=Book no.", "List of documents with different dates vouchers");
							}
							result = false;
							break;
						case "ma_phi_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_phi_i:H=Mã phí", "Danh sách mã phí không có trong danh mục phí");
							}
							else
							{
								BrowseError(table, "ma_phi_i:H=Fee code", "List of fee code is not in the list of fee");
							}
							result = false;
							break;
						case "ngay_ct_ks":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ngay_ct_ks:H=Ngày chứng từ:D", "Danh sách ngày chứng từ nhỏ hơn ngày khóa sổ");
							}
							else
							{
								BrowseError(table, "ngay_ct_ks:H=Voucher date:D", "List of documents smaller day close day window");
							}
							result = false;
							break;
						case "ma_px_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_px_i:H=Mã phân xưởng", "Danh sách mã phân xưởng không có trong danh mục phân xưởng");
							}
							else
							{
								BrowseError(table, "ma_px_i:H=Workshop code", "List of workshops code is not in the list of workshops");
							}
							result = false;
							break;
						case "ma_bpht_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_bpht_i:H=Mã bộ phận hạch toán", "Danh sách mã bộ phận hạch toán không có trong danh mục bộ phận hạch toán");
							}
							else
							{
								BrowseError(table, "ma_bpht_i:H=DEPT. ID", "List the accounting department code is not in the list of the accounting department");
							}
							result = false;
							break;
						case "dvt1":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_vt:H=Mã vật tư; dvt1:H=Đơn vị tính", "Danh sách đơn vị tính không có trong danh mục đơn vị tính quy đổi");
							}
							else
							{
								BrowseError(table, "ma_vt:H=Item code; dvt1:H=Manufacturing order", "List of units not included in the list of units of conversion");
							}
							result = false;
							break;
						case "ma_td_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td_i:H=Mã tự do 1", "Danh sách mã tự do 1 không có trong danh mục tự do 1");
							}
							else
							{
								BrowseError(table, "ma_td_i:H=Free code 1", "List of free codes 1 are not included in the list of free codes 1");
							}
							result = false;
							break;
						case "ma_td2_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td2_i:H=Mã tự do 2", "Danh sách mã tự do 2 không có trong danh mục tự do 2");
							}
							else
							{
								BrowseError(table, "ma_td2_i:H=Free code 2", "List of free codes 2 are not included in the list of free codes 2");
							}
							result = false;
							break;
						case "ma_td3_i":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_td3_i:H=Mã tự do 3", "Danh sách mã tự do 3 không có trong danh mục tự do 3");
							}
							else
							{
								BrowseError(table, "ma_td3_i:H=Free code 3", "List of free codes 3 are not included in the list of free codes 3");
							}
							result = false;
							break;
						case "ma_ms":
							if (StartupBase.M_LAN == "V")
							{
								BrowseError(table, "ma_ms:H=Nhóm hóa đơn", "Danh sách nhóm hóa đơn không có trong danh mục phân nhóm hóa đơn");
							}
							else
							{
								BrowseError(table, "ma_ms:H=Invoice group", "The invoice group list is not in the invoice grouping list");
							}
							result = false;
							break;
					}
				}
			}
			return result;
		}

		public static void BrowseError(DataTable data, string fields, string title)
		{
			if (StartUp.waiting != null)
			{
				StartUp.waiting.Close();
			}
			FormBrowse oBrowse = new FormBrowse(StartupBase.SasObj, data.DefaultView, fields);
			oBrowse.frmBrw.Title = SysFunc.Cat_Dau(title);
			oBrowse.frmBrw.LanguageID = "SasIeCt_4";
			oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)delegate
			{
				oBrowse.frmBrw.ChangeLanguage(StartupBase.M_LAN);
			});
			oBrowse.ShowDialog();
		}

	}   
}
