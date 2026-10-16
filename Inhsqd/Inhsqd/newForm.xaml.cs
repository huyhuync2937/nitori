using SasControls;
using SasDataLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Inhsqd
{
    public partial class newForm : FormList
    {
		private DataTable newDataTable = new DataTable();

		private DataTable OldRow;

		private string[] OldKeyValue = new string[2];

		private string[] OldKeyCol = new string[2];

		private string strOldNgay_ban = string.Empty;

		private string strOldMa_vt = string.Empty;

		private EditModeBindingObject FormInEditMode;

		private CodeValueBindingObject M_IP_GIA;

		private CodeValueBindingObject M_IP_GIA_NT;

		private CodeValueBindingObject M_IP_SL;
		public newForm()
		{
			InitializeComponent();
			base.BindingSasObj = StartupBase.SasObj;
			base.DisplayLanguage = StartupBase.M_LAN;
			SysFunc.LoadIcon(this);
			FormInEditMode = (EditModeBindingObject)FindResource("IsInEditMode");
			M_IP_GIA = (CodeValueBindingObject)FindResource("M_IP_GIA");
			M_IP_GIA_NT = (CodeValueBindingObject)FindResource("M_IP_GIA_NT");
			M_IP_SL = (CodeValueBindingObject)FindResource("M_IP_SL");
			if (StartUp.currActionTask == ActionTask.View)
			{
				ConfirmGV.ButtonType = 1;
			}
		}

		private void LoadForm()
		{
			DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(StartupBase.SasObj, StartUp.sqlTableName);
			txtDvt1.MaxLength = ListFunc.GetLengthColumn(sqlTableFieldList, "dvt1");
			OldKeyCol = StartUp.SqlTableKey.Split(';');
			OldKeyValue = StartUp.currSqlTableKey.Split(';');
			base.Title = SysFunc.Cat_Dau(StartUp.titleWindow);
		}

		private void newForm_Loaded(object sender, RoutedEventArgs e)
		{
			M_IP_GIA.Text = base.BindingSasObj.GetOption("M_IP_GIA").ToString();
			M_IP_GIA_NT.Text = base.BindingSasObj.GetOption("M_IP_GIA_NT").ToString();
			M_IP_SL.Text = base.BindingSasObj.GetOption("M_IP_SL").ToString();
			LoadForm();
			TextBox textBox = SysFunc.FindChild<TextBox>(this, "txtMa_vt");
			if (textBox != null)
			{
				textBox.SelectAll();
				textBox.Focus();
			}
			switch (StartUp.currActionTask)
			{
				case ActionTask.Add:
					try
					{
						newDataTable = StartUp.GetRow(StartUp.sqlTableView);
						DataRow dataRow = newDataTable.NewRow();
						dataRow["status"] = "1";
						dataRow["hs_qd"] = 1;
						StartupBase.SasObj.GetOption("M_AUTO_LIST_NUM").ToString().Equals("1");
						newDataTable.Rows.Add(dataRow);
					}
					catch (Exception ex4)
					{
						ErrorLog.CatchMessage(ex4);
					}
					txtMa_vt.IsFocus = true;
					break;
				case ActionTask.Edit:
					try
					{
						newDataTable = StartUp.GetRow(StartUp.sqlTableView);
						if (newDataTable.Rows.Count > 0)
						{
							OldRow = newDataTable.Copy();
						}
						strOldMa_vt = OldKeyValue[0];
						strOldNgay_ban = OldKeyValue[1];
					}
					catch (Exception ex2)
					{
						ErrorLog.CatchMessage(ex2);
					}
					txtMa_vt.IsFocus = true;
					break;
				case ActionTask.View:
					try
					{
						StartUp.LastEditTable = null;
						newDataTable = StartUp.GetRow(StartUp.sqlTableView);
					}
					catch (Exception ex3)
					{
						ErrorLog.CatchMessage(ex3);
					}
					txtDvt.Focusable = false;
					txtDvt1.Focusable = false;
					txtHs_qd.Focusable = false;
					ConfirmGV.pnlButton.btnCancel.Focus();
					break;
				case ActionTask.Copy:
					try
					{
						newDataTable = StartUp.GetRow(StartUp.sqlTableView);
						strOldMa_vt = OldKeyValue[0];
						strOldNgay_ban = OldKeyValue[1];
					}
					catch (Exception ex)
					{
						ErrorLog.CatchMessage(ex);
					}
					newDataTable.Rows[0]["status"] = "1";
					newDataTable.Rows[0]["hs_qd"] = 1;
					txtMa_vt.IsFocus = true;
					break;
			}
			gridMain.DataContext = newDataTable;
		}

		private void saveCustomer()
		{
			StartUp.LastEditTable = null;
			if (StartUp.currActionTask == ActionTask.View)
			{
				return;
			}
			newDataTable.AcceptChanges();
			int num = int.Parse(StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString());
			string value = StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString().Trim();
			if (newDataTable.Columns.Contains("search"))
			{
				SysFunc.SetStrSearch(StartupBase.SasObj, StartUp.sqlTableName, ref newDataTable);
			}
			newDataTable.Rows[0]["hs_qd"] = SysFunc.Round(ParseDecimal(newDataTable.Rows[0]["hs_qd"], 1m), StartUp.M_IP_SL);
			if (StartUp.currActionTask == ActionTask.Edit)
			{
				newDataTable.Rows[0]["date"] = DateTime.Now;
				newDataTable.Rows[0]["time"] = DateTime.Now.ToString("HH:mm:ss");
				newDataTable.Rows[0]["user_id"] = num;
				newDataTable.Rows[0]["user_name"] = value;
				if (OldRow != null)
				{
					ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, OldRow.Rows[0], newDataTable.Rows[0], StartupBase.SasObj);
				}
			}
			else
			{
				newDataTable.Rows[0]["date"] = DateTime.Now;
				newDataTable.Rows[0]["time"] = DateTime.Now.ToString("HH:mm:ss");
				newDataTable.Rows[0]["user_id"] = num;
				newDataTable.Rows[0]["user_name"] = value;
				newDataTable.Rows[0]["date0"] = DateTime.Now;
				newDataTable.Rows[0]["time0"] = DateTime.Now.ToString("HH:mm:ss");
				newDataTable.Rows[0]["user_id0"] = num;
				newDataTable.Rows[0]["user_name0"] = value;
				ListFunc.inserRowInDataBase(StartUp.sqlTableName, newDataTable.Rows[0], StartupBase.SasObj);
			}
			StartUp.LastEditTable = newDataTable;
		}

		public decimal ParseDecimal(object obj, decimal defaultvalue)
		{
			decimal result = defaultvalue;
			decimal.TryParse((obj != null) ? obj.ToString() : defaultvalue.ToString(), out result);
			return result;
		}

		private bool CheckValid()
		{
			bool flag = true;
			if (txtMa_vt.Text.Trim() == string.Empty && flag)
			{
				TabInfor.SelectedIndex = 0;
				ExMessageBox.Show(2370, StartupBase.SasObj, "Chưa vào mã vật tư!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				txtMa_vt.Focus();
				flag = false;
			}
			if (txtDvt1.Text.Trim() == string.Empty && flag)
			{
				ExMessageBox.Show(223, StartupBase.SasObj, "Chưa vào đơn vị tính quy đổi!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				txtDvt1.Focus();
				flag = false;
			}
			if (txtMa_vt.Text.Trim() != "" && flag && flag)
			{
				try
				{
					bool flag2 = true;
					if (StartUp.currActionTask == ActionTask.Edit && flag)
					{
						flag2 = false;
						if (OldRow.Rows[0]["ma_vt"].ToString().Trim() != newDataTable.Rows[0]["ma_vt"].ToString().Trim() || OldRow.Rows[0]["dvt1"].ToString().Trim() != newDataTable.Rows[0]["dvt1"].ToString().Trim())
						{
							flag2 = true;
						}
					}
					SqlCommand sqlCommand = new SqlCommand("select count(1)\tfrom " + StartUp.sqlTableName + " where ma_vt = @ma_vt and dvt1 = @dvt1");
					sqlCommand.Parameters.Add("@ma_vt", SqlDbType.Char).Value = txtMa_vt.Text.Trim();
					sqlCommand.Parameters.Add("@dvt1", SqlDbType.Char).Value = txtDvt1.Text.Trim();
					int num = (int)StartupBase.SasObj.ExcuteScalar(sqlCommand);
					if (num <= 0)
					{
						return flag;
					}
					if (!flag2)
					{
						return flag;
					}
					TabInfor.SelectedIndex = 0;
					ExMessageBox.Show(2380, StartupBase.SasObj, "Mã đã có hoặc lồng nhau!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
					txtMa_vt.IsFocus = true;
					flag = false;
					return flag;
				}
				catch (Exception ex)
				{
					ErrorLog.CatchMessage(ex);
					return flag;
				}
			}
			return flag;
		}

		public void EnableEditMode(bool isEditMode)
		{
			FormInEditMode.IsEditMode = isEditMode;
		}

		private void txtnh_kh1_LostFocus(object sender, RoutedEventArgs e)
		{
		}

		private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
		{
			if (StartUp.currActionTask == ActionTask.View)
			{
				Close();
			}
			else if (CheckValid())
			{
				saveCustomer();
				StartUp.currActionTask = ActionTask.None;
				Close();
			}
		}

		private void txtghi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
		{
			if (Keyboard.IsKeyDown(Key.Return) && Keyboard.Modifiers == ModifierKeys.Alt)
			{
				TextBox textBox = sender as TextBox;
				textBox.SelectedText = Environment.NewLine;
				textBox.SelectionStart++;
				textBox.SelectionLength = 0;
			}
		}

		protected override bool IsEnterToPassObject(object sender)
		{
			if (sender is TextBox && (sender as TextBox).Name == "txtghi_chu" && Keyboard.Modifiers == ModifierKeys.Alt)
			{
				return false;
			}
			return base.IsEnterToPassObject(sender);
		}

		private void FormList_Closed(object sender, EventArgs e)
		{
			StartUp.currActionTask = ActionTask.None;
		}

		private void ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
		{
			StartUp.LastEditTable = null;
			Close();
		}

		private void FormList_PreviewKeyUp(object sender, KeyEventArgs e)
		{
			if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Escape)
			{
				StartUp.LastEditTable = null;
				Close();
			}
		}

		private void txtMa_vt_LostFocus(object sender, RoutedEventArgs e)
		{
			if (txtMa_vt.RowResult == null)
			{
				txtten_vt.Text = "";
				return;
			}
			txtten_vt.Text = (StartupBase.M_LAN.Equals("V") ? txtMa_vt.RowResult["ten_vt"].ToString() : txtMa_vt.RowResult["ten_vt2"].ToString());
			txtDvt.Text = txtMa_vt.RowResult["dvt"].ToString();
		}

		private void txtstatus_LostFocus(object sender, RoutedEventArgs e)
		{
			if (txtstatus.Text == "")
			{
				newDataTable.Rows[0]["status"] = 0;
			}
		}

		private void txtHs_qd_GotFocus(object sender, RoutedEventArgs e)
		{
			txtHs_qd.SelectAll();
		}

	}
}
