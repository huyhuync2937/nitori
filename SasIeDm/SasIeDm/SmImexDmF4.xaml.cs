using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Threading;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Markup;
using System.Data;
using Microsoft.Win32;
using SasControls;
using SasDataLib;
using SasFormReport;
using SasLib;

namespace SasIeDm
{
    public partial class SmImexDmF4 : FormFilter
	{
        private bool bResult = false;

        public static string bangma = "";

        public static string Ten_file = "";
		private string TemplateTitle
		{
			get
			{
				DataRowView dataRowView = base.DataContext as DataRowView;
				return StartupBase.M_LAN.Equals("V") ? dataRowView["ten"].ToString().Trim() : dataRowView["ten2"].ToString().Trim();
			}
		}

		private string TableTemplate
		{
			get
			{
				DataRowView dataRowView = base.DataContext as DataRowView;
				return dataRowView["dbf_mau"].ToString().Trim();
			}
		}

		private string Khoa
		{
			get
			{
				DataRowView dataRowView = base.DataContext as DataRowView;
				return dataRowView["khoa"].ToString().Trim();
			}
		}

		private string Ma_Imex
		{
			get
			{
				DataRowView dataRowView = base.DataContext as DataRowView;
				return dataRowView["ma_imex"].ToString().Trim();
			}
		}
		public ImportInfo Info
		{
			get;
			set;
		}

		public SmImexDmF4()
		{
			InitializeComponent();
			txtFileName.Focus();
		}

		protected override void OnPropertyChanged(DependencyPropertyChangedEventArgs e)
		{
			base.OnPropertyChanged(e);
			if (e.Property == FrameworkElement.DataContextProperty)
			{
				DataRowView dataRowView = base.DataContext as DataRowView;
				string a = dataRowView["ma_imex"].ToString().Trim();
				if (a == StartUp.ma_imex_truoc)
				{
					txtFileName.Text = StartUp._paths;
				}
				else
				{
					txtFileName.Text = "";
				}
			}
		}

		private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
		{
			if (txtFileName.Text.Trim() == "")
			{
				ExMessageBox.Show(615, StartupBase.SasObj, string.Format("Chưa chọn file excel!", txtFileName.Text), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				txtFileName.Focus();
				return;
			}
			if (!File.Exists(txtFileName.Text))
			{
				ExMessageBox.Show(620, StartupBase.SasObj, $"Tập tin [{txtFileName.Text}] không tồn tại!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
				txtFileName.Focus();
				return;
			}
			Directory.CreateDirectory(Environment.GetEnvironmentVariable("Temp") + "\\ExcelImport\\");
			string text = Environment.GetEnvironmentVariable("Temp") + "\\ExcelImport\\" + Ten_file.Trim();
			string text2 = txtFileName.Text;
			if (File.Exists(text))
			{
				FileInfo fileInfo = new FileInfo(text);
				if (fileInfo.IsReadOnly)
				{
					fileInfo.IsReadOnly = false;
				}
			}
			File.Copy(text2, text, overwrite: true);
			bangma = txtBANGMA.Text.Trim();
			DataRowView dataRowView = base.DataContext as DataRowView;
			ImportInfo info = default(ImportInfo);
			info.Name = dataRowView["ten"].ToString().Trim();
			info.FileName = text;
			info.TableTemplate = dataRowView["dbf_mau"].ToString().Trim();
			info.ExcelTemplate = dataRowView["ex_mau"].ToString().Trim();
			info.Ma_Imex = dataRowView["ma_imex"].ToString().Trim();
			info.Khoa = dataRowView["khoa"].ToString().Trim();
			info.vBrowse = dataRowView["vbrowse1"].ToString();
			info.eBrowse = dataRowView["ebrowse1"].ToString();
			info.PostProc = dataRowView["postproc"].ToString().Trim();
			info.CheckTrung = dataRowView["kt_trung"].ToString().Trim();
			info.CheckLong = dataRowView["kt_long"].ToString().Trim();
			info.OverWrite = txtOverWrite.Text.Trim();
			Info = info;
			bResult = true;
			Close();
		}

		public new bool ShowDialog()
		{
			base.ShowDialog();
			return bResult;
		}

		private void btnDbf_Click(object sender, RoutedEventArgs e)
		{
			StartupBase.SasObj.SynchroFile(".\\Excel-Mau", txtdbf_mau.Text.Trim() + ".Xls");
			string text = StartupBase.SasObj.M_StartUp_Path + "Excel-Mau\\" + txtdbf_mau.Text.Trim() + ".Xls";
			if (File.Exists(text))
			{
				Process.Start(text);
			}
		}

		private void btnOpen_Click(object sender, RoutedEventArgs e)
		{
			string currentDirectory = Environment.CurrentDirectory;
			string text = txtFileName.Text;
			string fileName = "";
			if (text.Trim() != "")
			{
				text = Path.GetDirectoryName(txtFileName.Text);
				if (text.Trim() != "" && Directory.Exists(text))
				{
					Environment.CurrentDirectory = text;
				}
				fileName = Path.GetFileName(txtFileName.Text);
			}
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.FileName = fileName;
			openFileDialog.Filter = "Microsoft excel (*.Xls)|*.xls";
			if (openFileDialog.ShowDialog() == true)
			{
				txtFileName.Text = openFileDialog.FileName;
				Ten_file = openFileDialog.SafeFileName;
			}
			Environment.CurrentDirectory = currentDirectory;
		}

		private void btnMau_Click(object sender, RoutedEventArgs e)
		{
			StartupBase.SasObj.SynchroFile(".\\Excel-Mau", txtdbf_mau.Text.Trim() + ".Xls");
			string text = StartupBase.SasObj.M_StartUp_Path + "Excel-Mau\\" + txtdbf_mau.Text.Trim() + ".Xls";
			if (File.Exists(text))
			{
				Process.Start(text);
			}
			else
			{
				Debug.WriteLine($"File {text} does not exist.");
			}
		}

		private void btnXuatMau_Click(object sender, RoutedEventArgs e)
		{
			StartupBase.SasObj.SynchroFile(".\\Excel-Mau", txtdbf_mau.Text.Trim() + ".Xls");
			string sourceFileName = StartupBase.SasObj.M_StartUp_Path + "Excel-Mau\\" + txtdbf_mau.Text.Trim() + ".Xls";
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

		private void FrmSmImexDmF4_Closed(object sender, EventArgs e)
		{
			DataRowView dataRowView = base.DataContext as DataRowView;
			StartUp._paths = txtFileName.Text.Trim();
			StartUp.ma_imex_truoc = dataRowView["ma_imex"].ToString().Trim();
		}

		private void btnReview_Click(object sender, RoutedEventArgs e)
		{
			if (!string.IsNullOrEmpty(txtFileName.Text) && File.Exists(txtFileName.Text))
			{
				Process.Start(txtFileName.Text);
			}
		}


	}
}
