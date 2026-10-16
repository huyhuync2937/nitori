using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Inhsqd
{
	public class StartUp : StartupBase
    {
		public static string sqlTableName = "dmhsqd";

		public static string sqlTableView = "v_dmhsqd";

		public static DataRow CommandInfo;

		public static string SqlTableKey = "ma_vt;dvt1";

		public static string SqlTableObjectName = "ten_vt";

		public static string TableName = "hệ số quy đổi";

		public static string TableName_CatDau = "he so quy doi";

		public static string TableName_CatDau2 = "coefficient for conversion";

		public static FrameBrowse oBrowse;

		public static ActionTask currActionTask = ActionTask.None;

		public static string currSqlTableKey = string.Empty;

		public static string titleWindow = string.Empty;

		public static bool IsLoadFromLookUp = false;

		public static DataTable LastEditTable = null;

		public static int M_IP_SL = 0;

		public override void Run()
		{
			StartupBase.Namespace = "Inhsqd";
			try
			{
				StartupBase.M_LAN = StartupBase.SasObj.GetOption("M_LAN").ToString();
				M_IP_SL = Convert.ToInt16(StartupBase.SasObj.GetSysvar("M_ROUND_SL"));
				CommandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
				oBrowse = new FrameBrowse(StartupBase.SasObj, sqlTableName);
				oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? CommandInfo["bar"].ToString() : CommandInfo["bar2"].ToString());
				SysFunc.LoadIcon(oBrowse.frmBrw);
				oBrowse.F2 += oBrowse_F2;
				oBrowse.F3 += oBrowse_F3;
				oBrowse.F4 += oBrowse_F4;
				oBrowse.Ctrl_F4 += oBrowse_Ctrl_F4;
				oBrowse.F7 += oBrowse_F7;
				oBrowse.F8 += oBrowse_F8;
				oBrowse.frmBrw.ToolBar.Cm_Moi += ToolBar_Cm_Moi;
				oBrowse.frmBrw.ToolBar.Cm_Sua += ToolBar_Cm_Sua;
				oBrowse.frmBrw.ToolBar.Cm_Xoa += ToolBar_Cm_Xoa;
				oBrowse.frmBrw.ToolBar.Cm_In += ToolBar_Cm_In;
				oBrowse.frmBrw.ToolBar.Cm_Copy += ToolBar_Cm_Copy;
				oBrowse.frmBrw.ShowInTaskbar = true;
				oBrowse.frmBrw.LanguageID = "Inhsqd_1";
				oBrowse.ShowDialog();
			}
			catch (Exception ex)
			{
				ErrorLog.CatchMessage(ex);
			}
		}

		private void oBrowse_Ctrl_F4(object sender, EventArgs e)
		{
			Copy();
		}

		private void ToolBar_Cm_DoiMa()
		{
			DoiMa();
		}

		private void ToolBar_Cm_Copy()
		{
			Copy();
		}

		private void ToolBar_Cm_In()
		{
			In();
		}

		private void ToolBar_Cm_Xoa()
		{
			Xoa();
		}

		private void ToolBar_Cm_Sua()
		{
			Sua();
		}

		private void ToolBar_Cm_Moi()
		{
			Them();
		}

		private void oBrowse_F2(object sender, EventArgs e)
		{
			if (currActionTask == ActionTask.None && oBrowse.ActiveRecord != null)
			{
				currActionTask = ActionTask.View;
				currSqlTableKey = strKeyValue();
				titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Xem thông tin " + TableName) : ("View " + TableName_CatDau2 + " information"));
				showWindow();
			}
		}

		private void oBrowse_F3(object sender, EventArgs e)
		{
			Sua();
		}

		private void Sua()
		{
			if (currActionTask == ActionTask.None && oBrowse.ActiveRecord != null)
			{
				currActionTask = ActionTask.Edit;
				currSqlTableKey = strKeyValue();
				titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Sửa thông tin " + TableName) : ("Edit " + TableName_CatDau2 + " information"));
				showWindow();
			}
		}

		private string strKeyValue()
		{
			string[] array = SqlTableKey.Split(';');
			string text = string.Empty;
			string empty = string.Empty;
			for (int i = 0; i < array.Count(); i++)
			{
				empty = array[i];
				text = text + oBrowse.ActiveRecord.Cells[empty].Value.ToString() + ";";
			}
			return text;
		}

		private void oBrowse_F4(object sender, EventArgs e)
		{
			Them();
		}

		private void Them()
		{
			if (currActionTask == ActionTask.None)
			{
				currActionTask = ActionTask.Add;
				if (oBrowse.ActiveRecord != null)
				{
					currSqlTableKey = strKeyValue();
				}
				titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Thêm " + TableName) : ("Add " + TableName_CatDau2));
				showWindow();
			}
		}

		private void Copy()
		{
			if (currActionTask == ActionTask.None && oBrowse.ActiveRecord != null)
			{
				currActionTask = ActionTask.Copy;
				currSqlTableKey = strKeyValue();
				titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Thêm " + TableName) : ("Add " + TableName_CatDau2));
				showWindow();
			}
		}

		private void oBrowse_F6(object sender, EventArgs e)
		{
			DoiMa();
		}

		private void DoiMa()
		{
			if (currActionTask == ActionTask.None && oBrowse.ActiveRecord != null)
			{
				SysFunc.ChangeListID(StartupBase.SasObj, sqlTableName, oBrowse, SqlTableKey);
			}
		}

		private void oBrowse_F7(object sender, EventArgs e)
		{
			In();
		}

		private void In()
		{
			SqlCommand sqlcmd = new SqlCommand("select * from " + sqlTableView);
			DataSet localDataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
			ReportManager reportManager = new ReportManager(StartupBase.SasObj, CommandInfo["rep_file"].ToString());
			reportManager.Preview(localDataSet);
		}

		private void oBrowse_F8(object sender, EventArgs e)
		{
			Xoa();
		}

		private void Xoa()
		{
			if (currActionTask == ActionTask.None && oBrowse.ActiveRecord != null)
			{
				currSqlTableKey = strKeyValue();
				currActionTask = ActionTask.Delete;
				if (SysFunc.CheckPermission(StartupBase.SasObj, currActionTask, StartupBase.Menu_Id))
				{
					try
					{
						if (ExMessageBox.Show(2435, StartupBase.SasObj, "Có chắc chắn xóa không?", "Xac nhan nhap lieu", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
						{
							deleteRowByKey(sqlTableName);
						}
					}
					catch (SqlException sqlEx)
					{
						ErrorLog.CatchMessage(sqlEx);
					}
				}
				else
				{
					ExMessageBox.Show(2445, StartupBase.SasObj, "Không có quyền xóa [" + TableName + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				}
				currActionTask = ActionTask.None;
				oBrowse.frmBrw.ReloadData();
			}
		}

		private void deleteRowByKey(string tableName)
		{
			DataTable row = GetRow(tableName);
			if (row.Rows.Count > 0)
			{
				ListFunc.deleteRowInDatabaseByKey(tableName, SqlTableKey, row.Rows[0], StartupBase.SasObj);
			}
		}

		public DataTable Extend_oBrowse_Command(SasObject SasObj, ActionTask Mode, string Current_Ma_kh)
		{
			currActionTask = Mode;
			if (SysFunc.CheckPermission(SasObj, currActionTask, StartupBase.Menu_Id))
			{
				IsLoadFromLookUp = true;
				switch (Mode)
				{
					case ActionTask.Add:
						try
						{
							StartupBase.SasObj = SasObj;
							titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Thêm " + TableName) : ("Add " + TableName_CatDau2));
							showWindow();
						}
						catch (Exception ex3)
						{
							ErrorLog.CatchMessage(ex3);
						}
						break;
					case ActionTask.Edit:
						try
						{
							StartupBase.SasObj = SasObj;
							currSqlTableKey = Current_Ma_kh;
							titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Sửa thông tin " + TableName) : ("Edit " + TableName_CatDau2 + " information"));
							showWindow();
						}
						catch (Exception ex2)
						{
							ErrorLog.CatchMessage(ex2);
						}
						break;
					case ActionTask.View:
						try
						{
							StartupBase.SasObj = SasObj;
							currSqlTableKey = Current_Ma_kh;
							titleWindow = (StartupBase.M_LAN.Equals("V") ? ("Xem thông tin " + TableName) : ("View " + TableName_CatDau2 + " information"));
							showWindow();
						}
						catch (Exception ex)
						{
							ErrorLog.CatchMessage(ex);
						}
						break;
				}
			}
			else
			{
				ExMessageBox.Show(2450, SasObj, "Không có quyền [" + titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				currActionTask = ActionTask.None;
			}
			return LastEditTable;
		}

		private void showWindow()
		{
			if (SysFunc.CheckPermission(StartupBase.SasObj, currActionTask, StartupBase.Menu_Id))
			{
				newForm newForm = new newForm();
				switch (currActionTask)
				{
					case ActionTask.View:
						newForm.EnableEditMode(isEditMode: false);
						break;
					case ActionTask.Edit:
						newForm.EnableEditMode(isEditMode: true);
						break;
					case ActionTask.Copy:
						newForm.EnableEditMode(isEditMode: true);
						break;
					case ActionTask.Add:
						newForm.EnableEditMode(isEditMode: true);
						break;
				}
				newForm.ShowDialog();
				if (currActionTask != ActionTask.View && LastEditTable != null && oBrowse != null)
				{
					oBrowse.frmBrw.ReloadData(LastEditTable.Rows[0]);
				}
			}
			else
			{
				ExMessageBox.Show(2455, StartupBase.SasObj, "Không có quyền [" + titleWindow.ToLower() + "]!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				currActionTask = ActionTask.None;
			}
		}

		public static DataTable GetRow(string sqlTableName)
		{
			DataTable result = null;
			SqlCommand sqlCommand = new SqlCommand();
			try
			{
				string text = "select * from " + sqlTableName + " where ";
				if (currActionTask == ActionTask.Add)
				{
					text += "1=0";
				}
				else
				{
					DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, sqlTableName);
					string[] array = currSqlTableKey.Split(';');
					string[] array2 = SqlTableKey.Split(';');
					for (int i = 0; i < array2.Count(); i++)
					{
						if (i > 0)
						{
							text += " and";
						}
						DataRow[] array3 = sqlTableFieldList.Select("name='" + array2[i] + "'");
						string text2 = text;
						text = text2 + " " + array2[i] + "=@" + array2[i];
						if (i == 0)
						{
							sqlCommand.Parameters.Add("@" + array2[i], ListFunc.GetSqlDBType(array3[0]["datatype"].ToString())).Value = array[i];
						}
						else
						{
							sqlCommand.Parameters.Add("@" + array2[i], SqlDbType.NVarChar).Value = array[i];
						}
					}
				}
				sqlCommand.CommandText = text;
				result = StartupBase.SasObj.ExcuteReader(sqlCommand).Tables[0];
				return result;
			}
			catch (SqlException sqlEx)
			{
				ErrorLog.CatchMessage(sqlEx);
				return result;
			}
		}
	}
}
