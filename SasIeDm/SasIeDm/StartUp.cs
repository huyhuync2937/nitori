using Infragistics.Windows.DataPresenter;
using SasControls;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;

namespace SasIeDm
{
    public class StartUp : StartupBase
	{
		public static FrmWaiting waiting; 

		public static DataSet DataImport;

		public static string M_CHR_ERR = "";

		private DataSet dsDmimexdm;

		private DataRow drCommandInfo;

		private static FormBrowse oBrowse;

		private string[] sFieldArrays;

		public static int _User_id = 1;

		public static string _paths = "";

		public static string ma_imex_truoc = "";

		public static string M_CHECK_LONG = "0";

		public static DataTable dtImexMessage;

		public override void Run()
		{
			StartupBase.Namespace = "SasIeDm";
			Show(StartupBase.Menu_Id);
		}

		private void Show(string id)
		{
			M_CHECK_LONG = StartupBase.SasObj.GetOption("M_CHECK_LONG").ToString().Trim();
			M_CHR_ERR = StartupBase.SasObj.GetSysvar("M_CHR_ERR").ToString().Trim()
				.Replace(" ", "");
			drCommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, id);
			_User_id = (int)StartupBase.SasObj.UserInfo.Rows[0][0];
			if (drCommandInfo == null || drCommandInfo.ItemArray.Length == 0)
			{
				if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
				{
					Application.Current.Shutdown();
				}
				return;
			}
			SqlCommand sqlCommand = new SqlCommand(drCommandInfo["store_proc"].ToString());
			sqlCommand.CommandType = CommandType.StoredProcedure;
			sqlCommand.Parameters.Add("@ma_phan_he", SqlDbType.VarChar).Value = drCommandInfo["parameter"].ToString().Trim() + "%";
			dsDmimexdm = StartupBase.SasObj.ExcuteReader(sqlCommand);
			if (dsDmimexdm == null || dsDmimexdm.Tables.Count == 0)
			{
				if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
				{
					Application.Current.Shutdown();
				}
				return;
			}
			sqlCommand = new SqlCommand("Select * From ImexMessage order by ma_imex, stt");
			dtImexMessage = StartupBase.SasObj.ExcuteReader(sqlCommand).Tables[0].Copy();
			if (StartupBase.M_LAN == "V")
			{
				sFieldArrays = drCommandInfo["Vbrowse1"].ToString().Split("|".ToCharArray());
			}
			else
			{
				sFieldArrays = drCommandInfo["Ebrowse1"].ToString().Split("|".ToCharArray());
			}
			string strBrowse = sFieldArrays[0];
			oBrowse = new FormBrowse(StartupBase.SasObj, dsDmimexdm.Tables[0].DefaultView, strBrowse);
			oBrowse.CTRL_R += oBrowse_CTRL_R;
			oBrowse.F4 += oBrowse_F4;
			oBrowse.F3 += oBrowse_F3;
			oBrowse.F5 += oBrowse_F5;
			SysFunc.LoadIcon(oBrowse.frmBrw);
			object obj = oBrowse.frmBrw.ToolBar.FindName("tbReport");
			if (obj != null)
			{
				ToolBar toolBar = obj as ToolBar;
				(toolBar.Items[1] as ToolBarButton).Text = "Lấy tệp mẫu";
				(toolBar.Items[1] as ToolBarButton).ToolTip = "F3";
				(toolBar.Items[1] as ToolBarButton).IsEnabled = true;
				(toolBar.Items[2] as ToolBarButton).Text = "Lấy dữ liệu";
				(toolBar.Items[2] as ToolBarButton).ToolTip = "F4";
				(toolBar.Items[2] as ToolBarButton).IsEnabled = true;
				(toolBar.Items[2] as ToolBarButton).Click += StartUp_Click;
			}
			oBrowse.frmBrw.DisplayLanguage = StartupBase.M_LAN;
			oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? drCommandInfo["bar"].ToString().Trim() : drCommandInfo["bar2"].ToString().Trim());
			oBrowse.frmBrw.LanguageID = "SasIeDm_2";
			oBrowse.ShowDialog();
			if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
			{
				Application.Current.Shutdown();
			}
		}

		private void oBrowse_F5(object sender, EventArgs e)
		{
		}

		private void oBrowse_F3(object sender, EventArgs e)
		{
			if (oBrowse.DataGrid.ActiveRecord != null && oBrowse.DataGrid.ActiveRecord.RecordType == RecordType.DataRecord)
			{
				DataRowView dataRowView = (oBrowse.DataGrid.ActiveRecord as DataRecord).DataItem as DataRowView;
				StartupBase.SasObj.SynchroFile(".\\Excel-Mau", dataRowView["dbf_mau"].ToString().Trim() + ".Xls");
				string sourceFileName = StartupBase.SasObj.M_StartUp_Path + "Excel-Mau\\" + dataRowView["dbf_mau"].ToString().Trim() + ".Xls";
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "Excel 2003 (.xls)|*.xls";
				string text = "";
				if (saveFileDialog.ShowDialog() == true)
				{
					text = saveFileDialog.FileName;
					try
					{
						File.Copy(sourceFileName, text, overwrite: true);
					}
					catch (Exception ex)
					{
						ExMessageBox.Show(610, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
						return;
					}
				}
				if (File.Exists(text))
				{
					Process.Start(text);
				}
			}
		}

		private void StartUp_Click(object sender, RoutedEventArgs e)
		{
			oBrowse_F4(null, null);
		}

		private void StartUp_Click2(object sender, RoutedEventArgs e)
		{
			oBrowse_F3(null, null);
		}

		private void oBrowse_F4(object sender, EventArgs e)
		{
			if (oBrowse.DataGrid.ActiveRecord == null || oBrowse.DataGrid.ActiveRecord.RecordType != 0)
			{
				return;
			}
			SasIeDmF4 smImexDmF = new SasIeDmF4();
			smImexDmF.DataContext = (oBrowse.DataGrid.ActiveRecord as DataRecord).DataItem;
			smImexDmF.Title = oBrowse.frmBrw.Title;
			SysFunc.LoadIcon(smImexDmF);
			if (!smImexDmF.ShowDialog())
			{
				return;
			}
			waiting = new FrmWaiting(5.0);
			waiting.Show();
			switch (ThucThi(smImexDmF))
			{
				case 1:
					if (waiting != null)
					{
						waiting.Close();
					}
					ExMessageBox.Show(600, StartupBase.SasObj, "Chương trình đã thực hiện xong!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
					break;
				case 0:
					if (waiting != null)
					{
						waiting.Close();
					}
					ExMessageBox.Show(605, StartupBase.SasObj, "Chương trình huỷ bỏ số liệu đưa vào!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
					break;
				default:
					if (waiting != null)
					{
						waiting.Close();
					}
					break;
			}
		}

		private static int ThucThi(SasIeDmF4 imexCtF4)
		{
			DataImport = C_GetDataExcel.GetData(imexCtF4.Info, imexCtF4.txtBANGMA.Text);
			if (DataImport == null)
			{
				return 0;
			}
			if (DataImport.Tables.Count == 0 || DataImport.Tables[0].Rows.Count == 0)
			{
				return 2;
			}
			waiting.Set(1.0);
			if (!CheckKeyNull(DataImport.Tables["DataExcel"], imexCtF4.Info.Khoa))
			{
				return 0;
			}
			waiting.Set(2.0);
			if (!KiemTraKyTuDacBiet(DataImport.Tables["DataExcel"], imexCtF4.Info.Ma_Imex))
			{
				return 0;
			}
			waiting.Set(3.0);
			if (!C_ImportVoucher.UploadTable(imexCtF4.Info.TableTemplate, DataImport, C_GetDataExcel.StrBrowse))
			{
				return 0;
			}
			waiting.Set(4.0);
			if (!C_ImportVoucher.Post(imexCtF4.Info))
			{
				return 0;
			}
			waiting.Set(5.0);
			return 1;
		}

		private static bool CheckKeyNull(DataTable tb_excel, string list_key)
		{
			string[] array = list_key.Split(';');
			DataTable dataTable = tb_excel.Clone();
			for (int i = 0; i < tb_excel.Rows.Count; i++)
			{
				for (int j = 0; j < array.Length; j++)
				{
					if (tb_excel.Columns.Contains(array[j].Trim()) && tb_excel.Rows[i][array[j].Trim()].ToString().Trim() == "")
					{
						dataTable.ImportRow(tb_excel.Rows[i]);
						break;
					}
				}
			}
			if (dataTable.Rows.Count > 0)
			{
				C_ImportVoucher.ShowError(dataTable, C_GetDataExcel.StrBrowseKhoa, "Danh sách các cột bị bỏ trống");
				return false;
			}
			return true;
		}

		private static bool KiemTraKyTuDacBiet(DataTable tb_excel, string ma_iemx)
		{
			string text = C_GetDataExcel.MaKiemTraKyTuDacBiet(ma_iemx);
			string[] array = text.Split(';');
			DataTable dataTable = tb_excel.Clone();
			for (int i = 0; i < tb_excel.Rows.Count; i++)
			{
				for (int j = 0; j < array.Length; j++)
				{
					if (!tb_excel.Columns.Contains(array[j].Trim()))
					{
						continue;
					}
					for (int k = 0; k < M_CHR_ERR.Length; k++)
					{
						if (tb_excel.Rows[i][array[j].Trim()].ToString().Trim().IndexOf(M_CHR_ERR[k].ToString().Trim()) >= 0 || tb_excel.Rows[i][array[j].Trim()].ToString().Trim().IndexOf(" ") > 0)
						{
							dataTable.ImportRow(tb_excel.Rows[i]);
							break;
						}
					}
				}
			}
			if (dataTable.Rows.Count > 0)
			{
				C_ImportVoucher.ShowError(dataTable, C_GetDataExcel.StrBrowseKhoa_dac_biet, "Mã không được phép chứa khoảng trắng hoặc các ký tự đặc biệt " + M_CHR_ERR + "!");
				return false;
			}
			return true;
		}

		private void oBrowse_CTRL_R(object sender, EventArgs e)
		{
			int index = 0;
			if (oBrowse.ActiveRecord != null)
			{
				index = oBrowse.ActiveRecord.Index;
			}
			SqlCommand sqlCommand = new SqlCommand(drCommandInfo["store_proc"].ToString());
			sqlCommand.CommandType = CommandType.StoredProcedure;
			sqlCommand.Parameters.Add("@ma_phan_he", SqlDbType.VarChar).Value = drCommandInfo["parameter"].ToString();
			dsDmimexdm = StartupBase.SasObj.ExcuteReader(sqlCommand);
			oBrowse.DataGrid.DataSource = dsDmimexdm.Tables[0].DefaultView;
			if (oBrowse.DataGrid.Records.Count != 0)
			{
				if (index >= oBrowse.DataGrid.Records.Count)
				{
					index = oBrowse.DataGrid.Records.Count - 1;
				}
				oBrowse.frmBrw.Dispatcher.BeginInvoke((Action)delegate
				{
					oBrowse.DataGrid.ActiveRecord = oBrowse.DataGrid.Records[index];
				}, DispatcherPriority.Background);
			}
		}
	}

}
