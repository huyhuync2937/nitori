using SasControls;
using SasFormBrowes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace SisTCT
{
	/// <summary>
	/// Interaction logic for FrmMstInfo.xaml
	/// </summary>
	public partial class FrmMstInfo : Form
	{
		private string mst = string.Empty;

		private DataTable tbInfo = null;

		private bool isUpdate = false;

		private Point location;

		public static string M_TOOLTIP_CHECK_MST = string.Empty;

		public static string M_TOOLTIP_CHECK_MST2 = string.Empty;

		private bool isDm = false;

		public FrmMstInfo()
		{
			InitializeComponent();
			SysFunc.LoadIcon(this);
		}

		public FrmMstInfo(Point _localtion, string _mst, DataTable _tbInfo, bool _isUpdate, bool _isDm)
		{
			InitializeComponent();
			SysFunc.LoadIcon(this);
			if (!_isUpdate)
			{
				ConfirmGV.ButtonType = 1;
			}
			M_TOOLTIP_CHECK_MST = StartupBase.SasObj.GetSysvar("M_TOOLTIP_CHECK_MST").ToString().Trim();
			M_TOOLTIP_CHECK_MST2 = StartupBase.SasObj.GetSysvar2("M_TOOLTIP_CHECK_MST").ToString().Trim();
			GrdLayoutContent.RowDefinitions[1].Height = new GridLength(0.0, GridUnitType.Pixel);
			borderTenVietTat.BorderThickness = new Thickness(0.0, 1.0, 0.0, 1.0);
			borderTenVietTat1.BorderThickness = new Thickness(1.0, 1.0, 0.0, 1.0);
			borderTenQuocTe.BorderThickness = new Thickness(0.0, 0.0, 0.0, 1.0);
			borderTenQuocTe1.BorderThickness = new Thickness(1.0, 0.0, 0.0, 1.0);
			mst = _mst;
			tbInfo = _tbInfo;
			isUpdate = _isUpdate;
			location = _localtion;
			isDm = _isDm;
			setLocation();
		}

		public FrmMstInfo(Point _localtion, string _mst, DataTable _tbInfo, bool _isUpdate)
		{
			InitializeComponent();
			SysFunc.LoadIcon(this);
			if (!_isUpdate)
			{
				ConfirmGV.ButtonType = 1;
			}
			M_TOOLTIP_CHECK_MST = StartupBase.SasObj.GetSysvar("M_TOOLTIP_CHECK_MST").ToString().Trim();
			M_TOOLTIP_CHECK_MST2 = StartupBase.SasObj.GetSysvar2("M_TOOLTIP_CHECK_MST").ToString().Trim();
			GrdLayoutContent.RowDefinitions[1].Height = new GridLength(0.0, GridUnitType.Pixel);
			borderTenVietTat.BorderThickness = new Thickness(0.0, 1.0, 0.0, 1.0);
			borderTenVietTat1.BorderThickness = new Thickness(1.0, 1.0, 0.0, 1.0);
			borderTenQuocTe.BorderThickness = new Thickness(0.0, 0.0, 0.0, 1.0);
			borderTenQuocTe1.BorderThickness = new Thickness(1.0, 0.0, 0.0, 1.0);
			mst = _mst;
			tbInfo = _tbInfo;
			isUpdate = _isUpdate;
			location = _localtion;
			setLocation();
		}

		public FrmMstInfo(Point _localtion, string _mst, DataTable _tbInfo)
		{
			InitializeComponent();
			SysFunc.LoadIcon(this);
			if (!isUpdate)
			{
				ConfirmGV.ButtonType = 1;
			}
			M_TOOLTIP_CHECK_MST = StartupBase.SasObj.GetSysvar("M_TOOLTIP_CHECK_MST").ToString().Trim();
			M_TOOLTIP_CHECK_MST2 = StartupBase.SasObj.GetSysvar2("M_TOOLTIP_CHECK_MST").ToString().Trim();
			GrdLayoutContent.RowDefinitions[8].Height = new GridLength(0.0, GridUnitType.Pixel);
			borderTenVietTat.BorderThickness = new Thickness(0.0, 0.0, 0.0, 1.0);
			borderTenVietTat1.BorderThickness = new Thickness(1.0, 0.0, 0.0, 1.0);
			borderTenQuocTe.BorderThickness = new Thickness(0.0);
			borderTenQuocTe1.BorderThickness = new Thickness(1.0, 0.0, 0.0, 0.0);
			mst = _mst;
			tbInfo = _tbInfo;
			location = _localtion;
			setLocation();
		}

		private void setLocation()
		{
			double primaryScreenWidth = SystemParameters.PrimaryScreenWidth;
			double primaryScreenHeight = SystemParameters.PrimaryScreenHeight;
			Point point = location;
			double num = primaryScreenWidth - point.X - base.Width;
			double num2 = primaryScreenHeight - point.Y - base.Height;
			base.Left = point.X + ((num < 0.0) ? num : 0.0);
			base.Top = point.Y + ((num2 < 0.0) ? num2 : 0.0);
		}

		private void FormList_Loaded(object sender, RoutedEventArgs e)
		{
			string[] _M_TOOLTIP_CHECK_MST = M_TOOLTIP_CHECK_MST.Split(';');
			string[] _M_TOOLTIP_CHECK_MST2 = M_TOOLTIP_CHECK_MST2.Split(';');
            this.Dispatcher.BeginInvoke((Action)delegate
            {
                ConfirmGV.pnlButton.btnOk.Content = (StartupBase.M_LAN.Equals("V") ? "Cập nhật" : "Update");
                ConfirmGV.pnlButton.btnCancel.Content = (StartupBase.M_LAN.Equals("V") ? "Đóng" : "Close");
                if (!isUpdate)
                {
                    ConfirmGV.ButtonType = 1;
                }
                if (!isDm)
                {
                    base.Title = (StartupBase.M_LAN.Equals("V") ? SasControls.ControlLib.ControlFunction.Cat_Dau(_M_TOOLTIP_CHECK_MST[0]) : _M_TOOLTIP_CHECK_MST[1]);
                }
                else
                {
                    base.Title = (StartupBase.M_LAN.Equals("V") ? SasControls.ControlLib.ControlFunction.Cat_Dau(_M_TOOLTIP_CHECK_MST2[0]) : _M_TOOLTIP_CHECK_MST2[1]);
                }
            }, DispatcherPriority.Background);

            txtma_so_thue.Text = mst;
			List<DataRow> list = new List<DataRow>();
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "ten_cong_ty"
					select x).ToList();
			if (list.Count > 0)
			{
				txtten_cong_ty.Text = list.FirstOrDefault()["value"].ToString().Trim().ToUpper();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "ten_viet_tat"
					select x).ToList();
			if (list.Count > 0)
			{
				txtten_viet_tat.Text = list.FirstOrDefault()["value"].ToString().Trim().ToUpper();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "dia_chi"
					select x).ToList();
			if (list.Count > 0)
			{
				txtdia_chi.Text = list.FirstOrDefault()["value"].ToString().Trim();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "nguoi_dai_dien"
					select x).ToList();
			if (list.Count > 0)
			{
				txtnguoi_dai_dien.Text = list.FirstOrDefault()["value"].ToString().Trim();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "dien_thoai"
					select x).ToList();
			if (list.Count > 0)
			{
				txtdien_thoai.Text = getSDT(list.FirstOrDefault()["value"].ToString().Trim()).Trim();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "ngay_hoat_dong"
					select x).ToList();
			if (list.Count > 0)
			{
				try
				{
					DateTime dateTime = Convert.ToDateTime(list.FirstOrDefault()["value"].ToString().Trim());
					txtngay_thanh_lap.Text = dateTime.ToString("dd-MM-yyyy");
				}
				catch
				{
					txtngay_thanh_lap.Text = string.Empty;
				}
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "ten_quoc_te"
					select x).ToList();
			if (list.Count > 0)
			{
				txtten_quoc_te.Text = list.FirstOrDefault()["value"].ToString().Trim();
			}
			list = (from x in tbInfo.Select()
					where x["key"].ToString() == "tinh_trang"
					select x).ToList();
			if (list.Count > 0)
			{
				txttinh_trang.Text = list.FirstOrDefault()["value"].ToString().Trim();
				txttinh_trang1.Text = list.FirstOrDefault()["value"].ToString().Trim();
			}
			base.Dispatcher.BeginInvoke((Action)delegate
			{
				txtma_so_thue.Focus();
				Win32.SendKey(ModifierKeys.None, Key.Tab);
			}, DispatcherPriority.Background);
		}

		private string getNumberInString(string s)
		{
			string text = string.Empty;
			for (int i = 0; i < s.Length; i++)
			{
				char c = s[i];
				if (!char.IsDigit(c))
				{
					break;
				}
				text += c;
			}
			return text;
		}

		private string getSDT(string s)
		{
			string text = string.Empty;
			for (int i = 0; i < s.Length; i++)
			{
				char c = s[i];
				if (!c.ToString().Equals(" ") && !c.ToString().Equals("-") && !c.ToString().Equals(".") && !char.IsDigit(c))
				{
					break;
				}
				text += c;
			}
			return text;
		}

		private void FormList_Closed(object sender, EventArgs e)
		{
		}

		private void ConfirmGV_OnOk(object sender, RoutedEventArgs e)
		{
			if (isUpdate)
			{
				base.DialogResult = true;
			}
			Close();
		}

		private void ConfirmGV_OnCancel(object sender, RoutedEventArgs e)
		{
			Close();
		}
	}

}
