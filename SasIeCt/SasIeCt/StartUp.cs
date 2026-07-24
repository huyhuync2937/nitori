using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Microsoft.Win32;

namespace SasIeCt
{
	public class StartUp: StartupBase
{
	private DataSet dsDmimexct;

	public static DataRow drCommandInfo;

	private FormBrowse oBrowse;

	private SqlCommand cmd_gia2 = new SqlCommand();

	public static string M_UPDATE_GIA2 = ""; 

	public static string Ma_nt0 = "VND";

	public static int _User_id = 1;

	public static string _paths = ""; 

	public static string ma_imex_truoc = "";

	public static FrmWaiting waiting;

	public static DataSet DataImport;

	public static string Ws_Id
	{
		get;
		set;
	}

	public override void Run()
	{
		StartupBase.Namespace = "SasIeCt";
		Show(StartupBase.Menu_Id);
	}

	private void Show(string id)
	{
		drCommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, id);
		Ma_nt0 = (string)StartupBase.SasObj.GetOption("M_MA_NT0");
		M_UPDATE_GIA2 = StartupBase.SasObj.GetOption("M_UPDATE_GIA2").ToString().Trim();
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
		sqlCommand.Parameters.Add("@ma_phan_he", SqlDbType.Char).Value = (string.IsNullOrEmpty(drCommandInfo["parameter"].ToString()) ? "%%" : drCommandInfo["parameter"].ToString().Trim() + "%");

		string sql = "exec " + drCommandInfo["store_proc"].ToString() + " @ma_phan_he='" + (string.IsNullOrEmpty(drCommandInfo["parameter"].ToString()) ? "%%" : drCommandInfo["parameter"].ToString().Trim() + "%") + "'";

            string s = (string.IsNullOrEmpty(drCommandInfo["parameter"].ToString()) ? "%%" : drCommandInfo["parameter"].ToString().Trim() + "%");
            dsDmimexct = StartupBase.SasObj.ExcuteReader(sqlCommand);
		if (dsDmimexct == null || dsDmimexct.Tables.Count == 0)
		{
			if (!Process.GetCurrentProcess().ProcessName.Equals("SasProcess"))
			{
				Application.Current.Shutdown();
			}
			return;
		}
		string[] array = (!(StartupBase.M_LAN == "V")) ? drCommandInfo["Ebrowse1"].ToString().Split("|".ToCharArray()) : drCommandInfo["Vbrowse1"].ToString().Split("|".ToCharArray());
		string strBrowse = array[0];
		oBrowse = new FormBrowse(StartupBase.SasObj, dsDmimexct.Tables[0].DefaultView, strBrowse);
		oBrowse.CTRL_R += oBrowse_CTRL_R;
		oBrowse.F3 += oBrowse_F3;
		oBrowse.F4 += oBrowse_F4;
		oBrowse.F5 += oBrowse_F5;
		oBrowse.frmBrw.Closed += frmBrw_Closed;
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
		oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? drCommandInfo["bar"].ToString() : drCommandInfo["bar2"].ToString());
		oBrowse.frmBrw.LanguageID = "SasIeCt_3";
		oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)delegate
		{
			oBrowse.frmBrw.ChangeLanguage();
		});
		oBrowse.ShowDialog();
		if (!Process.GetCurrentProcess().ProcessName.Equals("SaProcess"))
		{
			Application.Current.Shutdown();
		}
	}

	private void StartUp_Click2(object sender, RoutedEventArgs e)
	{
		oBrowse_F3(null, null);
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

	private void frmBrw_Closed(object sender, EventArgs e)
	{
		cmd_gia2.CommandText = "update options set value = '" + M_UPDATE_GIA2.Trim() + "' where name = 'M_UPDATE_GIA2'";
		StartupBase.SasObj.ExcuteNonQuery(cmd_gia2);
		Environment.Exit(Environment.ExitCode);
	}

	private void StartUp_Click(object sender, RoutedEventArgs e)
	{
		oBrowse_F4(null, null);
	}

	private void oBrowse_F5(object sender, EventArgs e)
	{
	}

	private void oBrowse_F4(object sender, EventArgs e)
	{
		if (oBrowse.DataGrid.ActiveRecord == null || oBrowse.DataGrid.ActiveRecord.RecordType != 0)
		{
			return;
		}
		SasIeCtF4 SasIeCtF = new SasIeCtF4(); 
		SasIeCtF.DataContext = (oBrowse.DataGrid.ActiveRecord as DataRecord).DataItem;
		SasIeCtF.Title = SysFunc.Cat_Dau(SasIeCtF.Title);
		SasIeCtF.BindingSasObj = StartupBase.SasObj;
		SysFunc.LoadIcon(SasIeCtF);
		if (!SasIeCtF.ShowDialog())
		{
			return;
		}
		waiting = new SasIeCt.FrmWaiting(2.0);
		waiting.Show();
		switch (ThucThi(SasIeCtF))
		{
			case 1:
				if (waiting != null)
				{
					waiting.Close();
				}
				ExMessageBox.Show(170, StartupBase.SasObj, "Chương trình đã thực hiện xong!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				break;
			case 0:
				if (waiting != null)
				{
					waiting.Close();
				}
				ExMessageBox.Show(175, StartupBase.SasObj, "Chương trình huỷ bỏ số liệu đưa vào!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				break;
			default:
				if (waiting != null)
				{
					waiting.Close();
				}
				break;
		}
	}

	private int ThucThi(SasIeCtF4 imexCtF4)
	{
		try
		{
			DataImport = C_GetDataExcel.GetData(imexCtF4.Info, imexCtF4.txtBangMa.Text);
			if (DataImport == null)
			{
				return 0;
			}
			if (DataImport.Tables.Count == 0 || DataImport.Tables[0].Rows.Count == 0)
			{
				return 2;
			}
			if (!CheckKeyNull(DataImport.Tables["DataExcel"], imexCtF4.Info.FieldNotNull))
			{
				return 0;
			}
			if (!C_ImportVoucher.UploadTable(imexCtF4.Info.TableTemplate, DataImport, C_GetDataExcel.StrBrowse))
			{
				return 0;
			}
			if (!C_ImportVoucher.Check_Data(imexCtF4.Info)) 
			{
				return 0;
			}
			cmd_gia2.CommandText = "update options set value = '0' where name = 'M_UPDATE_GIA2'";
			StartupBase.SasObj.ExcuteNonQuery(cmd_gia2);
			if (!C_ImportVoucher.Post(imexCtF4.Info))
			{
				return 0;
			}
		}
		catch
		{
			return 0;
		}
		finally
		{
			cmd_gia2.CommandText = "update options set value = '" + M_UPDATE_GIA2.Trim() + "' where name = 'M_UPDATE_GIA2'";
			StartupBase.SasObj.ExcuteNonQuery(cmd_gia2);
		}
		return 1;
	}

	private static bool CheckKeyNull(DataTable tb_excel, string list_key)
	{
		if (list_key.Trim() == "")
		{
			return true;
		}
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
			FormBrowse formBrowse = new FormBrowse(StartupBase.SasObj, dataTable.DefaultView, C_GetDataExcel.StrBrowseFieldNull);
			formBrowse.frmBrw.Title = SysFunc.Cat_Dau("Danh sách các cột bị bỏ trống");
			formBrowse.frmBrw.LanguageID = "SasIeCt_41";
			formBrowse.ShowDialog();
			return false;
		}
		return true;
	}

	private void oBrowse_CTRL_R(object sender, EventArgs e)
	{
		SqlCommand sqlCommand = new SqlCommand(drCommandInfo["store_proc"].ToString());
		sqlCommand.CommandType = CommandType.StoredProcedure;
		sqlCommand.Parameters.Add("@ma_phan_he", SqlDbType.Char).Value = (string.IsNullOrEmpty(drCommandInfo["parameter"].ToString()) ? "%%" : drCommandInfo["parameter"].ToString());
		dsDmimexct = StartupBase.SasObj.ExcuteReader(sqlCommand);
		oBrowse.DataGrid.DataSource = dsDmimexct.Tables[0].DefaultView;
		if (oBrowse.DataGrid.Records.Count > 0)
		{
			oBrowse.frmBrw.Dispatcher.BeginInvoke((Action)delegate
			{
				oBrowse.DataGrid.ActiveRecord = oBrowse.DataGrid.Records[0];
			}, DispatcherPriority.Background);
		}
	}
}


}
