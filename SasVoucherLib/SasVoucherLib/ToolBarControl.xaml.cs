using SasControls;
using SasControls.ControlLib;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Threading;

namespace SasVoucherLib
{
    /// <summary>Interaction logic for ToolBarControl.xaml</summary>
    /// <summary>ToolBarControl</summary>
    public partial class ToolBarControl : UserControl
    {
        public static readonly DependencyProperty SasObjProperty = DependencyProperty.Register(nameof(SasObj), typeof(SasObject), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty DisableOpacityProperty = DependencyProperty.Register(nameof(DisableOpacity), typeof(double), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((object)0.3));
        public static readonly DependencyProperty IsQLProperty = DependencyProperty.Register(nameof(IsQL), typeof(bool), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((object)false, new PropertyChangedCallback(ToolBarControl.IsQLChanged)));
        public static readonly DependencyProperty IsNd51Property = DependencyProperty.Register(nameof(IsNd51), typeof(bool), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((object)false, new PropertyChangedCallback(ToolBarControl.OnIsNd51Changed)));
        /// <summary>
        /// Using a DependencyProperty as the backing store for IsInEditMode.  This enables animation, styling, binding, etc...
        /// </summary>
        public static readonly DependencyProperty IsInEditModeProperty = DependencyProperty.Register(nameof(IsInEditMode), typeof(bool), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((object)false, new PropertyChangedCallback(ToolBarControl.OnEditModeChanged)));
        public static readonly DependencyProperty DViewSourceProperty = DependencyProperty.Register(nameof(DViewSource), typeof(DataView), typeof(ToolBarControl), (PropertyMetadata)new UIPropertyMetadata((object)null, new PropertyChangedCallback(ToolBarControl.DViewSourceChanged)));
        public int currentTemplate = -1;
        private string AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        public string[] ListAdd;
        public string[] ListEdit;
        public string[] ListDelete;
        public string[] ListPrint;
        public DataSet dsTemplate;
        public string Menu_id;
        public string LanguageID;
        public string currentStt_rec;
        public int IsAdmin;


        public event ToolBarControl.Execute Cm_ContextMenu;

        public event ToolBarControl.Execute Cm_BookMark;

        public DataRow VoucherRow { get; set; }

        public FormTrans FormParent { get; set; }

        public SasObject SasObj
        {
            get
            {
                return (SasObject)this.GetValue(ToolBarControl.SasObjProperty);
            }
            set
            {
                this.SetValue(ToolBarControl.SasObjProperty, (object)value);
            }
        }

        public ToolBarControl()
        {
            this.InitializeComponent();
            this.OnIsInEditModeChanged();
            if (!StartUpTrans.M_LAN.Equals("V"))
            {
                this.btnAddBookmark.Content = (object)"Add";
                this.btnViewBookmark.Content = (object)"View";
                this.btnDeleteBookmark.Content = (object)"Delete";
                this.GridMain.pnlButton.btnOk.Content = (object)"_Ok";
                this.GridMain.pnlButton.btnCancel.Content = (object)"_Cancel";
            }
            if (StartUpTrans.DsTrans != null)
            {
                if (StartupBase.SasObj.GetSysvar("M_CHECK_VOUCHER").ToString().Trim().Equals("0") || !StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("hd_thue"))
                    this.cb_hd_thue.Visibility = Visibility.Collapsed;
            }
            else
                this.cb_hd_thue.Visibility = Visibility.Collapsed;
            object selectedItem;
            this.GrdBookmark.SelectionChanged += (SelectionChangedEventHandler)((s, a) => selectedItem = this.GrdBookmark.SelectedItem);
        }

        /// <summary>Add tool bar button click handle.</summary>
        /// <param name="handler"></param>
        public void AddButtonClick(Delegate handler)
        {
            this.AddHandler(ButtonBase.ClickEvent, handler);
        }

        /// <summary>Remove tool bar button click handle.</summary>
        /// <param name="handler"></param>
        public void RemoveButtonClick(RoutedEventHandler handler)
        {
            this.RemoveHandler(ButtonBase.ClickEvent, (Delegate)handler);
        }

        /// <summary>
        /// Button disable Opacity. This is a dependency property.
        /// </summary>
        public double DisableOpacity
        {
            get
            {
                return (double)this.GetValue(ToolBarControl.DisableOpacityProperty);
            }
            set
            {
                this.SetValue(ToolBarControl.DisableOpacityProperty, (object)value);
            }
        }

        public bool IsQL
        {
            get
            {
                return (bool)this.GetValue(ToolBarControl.IsQLProperty);
            }
            set
            {
                this.SetValue(ToolBarControl.IsQLProperty, (object)value);
            }
        }

        public static void IsQLChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            ToolBarControl toolbar = sender as ToolBarControl;
            toolbar.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => toolbar.OnIsInEditModeChanged()));
        }

        public bool IsNd51
        {
            get
            {
                return (bool)this.GetValue(ToolBarControl.IsNd51Property);
            }
            set
            {
                this.SetValue(ToolBarControl.IsNd51Property, (object)value);
            }
        }

        public static void OnIsNd51Changed(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {
            ToolBarControl toolBarControl = sender as ToolBarControl;
            if (toolBarControl.IsNd51)
                toolBarControl.vc.mnuInInfo.Visibility = Visibility.Visible;
            else
                toolBarControl.vc.mnuInInfo.Visibility = Visibility.Collapsed;
            toolBarControl.OnIsInEditModeChanged();
        }

        /// <summary>
        /// Get/set tool bar mode. View or Edit moding. This is a dependency property.
        /// </summary>
        public bool IsInEditMode
        {
            get
            {
                return (bool)this.GetValue(ToolBarControl.IsInEditModeProperty);
            }
            set
            {
                this.SetValue(ToolBarControl.IsInEditModeProperty, (object)value);
                this.OnIsInEditModeChanged();
            }
        }

        public DataView DViewSource
        {
            get
            {
                return (DataView)this.GetValue(ToolBarControl.DViewSourceProperty);
            }
            set
            {
                this.SetValue(ToolBarControl.DViewSourceProperty, (object)value);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void OnEditModeChanged(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {
            ToolBarControl toolbar = sender as ToolBarControl;
            toolbar.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => toolbar.OnIsInEditModeChanged()));
        }

        public static void DViewSourceChanged(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {
            ToolBarControl toolbar = sender as ToolBarControl;
            toolbar.OnIsInEditModeChanged();
            toolbar.DViewSource.ListChanged += (ListChangedEventHandler)((s, a) => toolbar.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => toolbar.OnIsInEditModeChanged())));
        }

        /// <summary>
        /// Setup button setting when IsInEditMode property changed.
        /// </summary>
        /// <param name="d"></param>
        /// <param name="e"></param>
        public void OnIsInEditModeChanged()
        {
            if (this.IsAdmin == 0)
                this.btnOptions.Visibility = Visibility.Collapsed;
            else
                this.btnOptions.Visibility = Visibility.Visible;
            if (this.IsQL)
            {
                this.btnNew.IsEnabled = false;
                this.btnNew.Opacity = this.DisableOpacity;
                this.btnEdit.IsEnabled = false;
                this.btnEdit.Opacity = this.DisableOpacity;
                this.btnCopy.IsEnabled = false;
                this.btnCopy.Opacity = this.DisableOpacity;
                this.btnBottom.IsEnabled = false;
                this.btnBottom.Opacity = this.DisableOpacity;
                this.btnCancel.IsEnabled = false;
                this.btnCancel.Opacity = this.DisableOpacity;
                this.btnNext.IsEnabled = false;
                this.btnNext.Opacity = this.DisableOpacity;
                this.btnPrevious.IsEnabled = false;
                this.btnPrevious.Opacity = this.DisableOpacity;
                this.btnReportFromVoucher.IsEnabled = false;
                this.btnReportFromVoucher.Opacity = this.DisableOpacity;
                this.btnSave.IsEnabled = false;
                this.btnSave.Opacity = this.DisableOpacity;
                this.btnSearch.IsEnabled = false;
                this.btnSearch.Opacity = this.DisableOpacity;
                this.btnTop.IsEnabled = false;
                this.btnTop.Opacity = this.DisableOpacity;
                this.btnView.IsEnabled = false;
                this.btnView.Opacity = this.DisableOpacity;
                this.btnVoucherConextMenu.IsEnabled = false;
                this.btnVoucherConextMenu.Opacity = this.DisableOpacity;
                this.btnVoucherBookmarkMenu.IsEnabled = false;
                this.btnVoucherBookmarkMenu.Opacity = this.DisableOpacity;
                this.btnDelete.IsEnabled = false;
                this.btnDelete.Opacity = this.DisableOpacity;
                this.btnOptions.IsEnabled = false;
                this.btnOptions.Opacity = this.DisableOpacity;
            }
            else if (this.DViewSource != null)
            {
                bool flag = this.DViewSource.Table.Rows.Count > 1;
                this.btnSave.IsEnabled = this.IsInEditMode;
                this.btnCancel.IsEnabled = this.IsInEditMode;
                this.btnNew.IsEnabled = !this.IsInEditMode;
                if (!this.IsNd51)
                {
                    this.btnEdit.IsEnabled = !this.IsInEditMode && flag;
                    this.btnDelete.IsEnabled = !this.IsInEditMode && flag;
                    this.btnPrint.IsEnabled = !this.IsInEditMode && flag;
                }
                this.btnCopy.IsEnabled = !this.IsInEditMode && flag;
                this.btnView.IsEnabled = !this.IsInEditMode && flag;
                this.btnSearch.IsEnabled = !this.IsInEditMode;
                this.btnTop.IsEnabled = !this.IsInEditMode && flag;
                this.btnPrevious.IsEnabled = !this.IsInEditMode && flag;
                this.btnNext.IsEnabled = !this.IsInEditMode && flag;
                this.btnBottom.IsEnabled = !this.IsInEditMode && flag;
                this.btnVoucherBookmarkMenu.IsEnabled = !this.IsInEditMode && flag;
                this.btnReportFromVoucher.IsEnabled = !this.IsInEditMode;
                this.btnOptions.IsEnabled = !this.IsInEditMode;
                this.btnSave.Opacity = this.btnSave.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnCancel.Opacity = this.btnCancel.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnNew.Opacity = this.btnNew.IsEnabled ? 1.0 : this.DisableOpacity;
                if (!this.IsNd51)
                {
                    this.btnEdit.Opacity = this.btnEdit.IsEnabled ? 1.0 : this.DisableOpacity;
                    this.btnDelete.Opacity = this.btnDelete.IsEnabled ? 1.0 : this.DisableOpacity;
                    this.btnPrint.Opacity = this.btnPrint.IsEnabled ? 1.0 : this.DisableOpacity;
                }
                this.btnEdit.Opacity = this.btnEdit.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnDelete.Opacity = this.btnDelete.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnCopy.Opacity = this.btnCopy.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnView.Opacity = this.btnView.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnSearch.Opacity = this.btnSearch.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnTop.Opacity = this.btnTop.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnPrevious.Opacity = this.btnPrevious.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnNext.Opacity = this.btnNext.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnBottom.Opacity = this.btnBottom.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnVoucherConextMenu.Opacity = this.btnVoucherConextMenu.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnVoucherBookmarkMenu.Opacity = this.btnVoucherBookmarkMenu.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnReportFromVoucher.Opacity = this.btnReportFromVoucher.IsEnabled ? 1.0 : this.DisableOpacity;
                this.btnOptions.Opacity = this.btnOptions.IsEnabled ? 1.0 : this.DisableOpacity;
                if (this.IsAdmin != 1)
                {
                    if (!((IEnumerable<string>)this.ListAdd).Contains<string>(this.Menu_id))
                    {
                        this.btnNew.IsEnabled = false;
                        this.btnNew.Opacity = this.DisableOpacity;
                        this.btnCopy.IsEnabled = false;
                        this.btnCopy.Opacity = this.DisableOpacity;
                    }
                    if (!((IEnumerable<string>)this.ListEdit).Contains<string>(this.Menu_id))
                    {
                        this.btnEdit.IsEnabled = false;
                        this.btnEdit.Opacity = this.DisableOpacity;
                    }
                    if (!((IEnumerable<string>)this.ListDelete).Contains<string>(this.Menu_id))
                    {
                        this.btnDelete.IsEnabled = false;
                        this.btnDelete.Opacity = this.DisableOpacity;
                    }
                    if (!((IEnumerable<string>)this.ListPrint).Contains<string>(this.Menu_id))
                    {
                        this.btnPrint.IsEnabled = false;
                        this.btnPrint.Opacity = this.DisableOpacity;
                    }
                }
            }
            if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt") && (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("9") || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("4")))
            {
                this.btnPrint.IsEnabled = false;
                this.btnPrint.Opacity = 0.3;
                this.btnEdit.IsEnabled = false;
                this.btnEdit.Opacity = 0.3;
                this.btnDelete.IsEnabled = false;
                this.btnDelete.Opacity = 0.3;
            }
            if (StartupBase.SasObj.GetSysvar("M_CHECK_VOUCHER_WITH_MODE").ToString().Trim().Equals("1"))
            {
                this.cb_hd_thue.IsEnabled = true;
            }
            else
            {
                if (!StartupBase.SasObj.GetSysvar("M_CHECK_VOUCHER_WITH_MODE").ToString().Trim().Equals("2"))
                    return;
                this.cb_hd_thue.IsEnabled = this.IsInEditMode;
            }

        }
        private void btnVoucherConextMenu_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (this.Cm_ContextMenu == null)
                return;
            this.Cm_ContextMenu();
        }

        private void btnVoucherConextMenu_Click(object sender, RoutedEventArgs e)
        {
            (sender as Button).ContextMenu.IsOpen = true;
        }

        private void ToolBarVoucher_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Right)
            {
                if (this.btnReportFromVoucher.IsFocused)
                {
                    if (this.btnNew.IsEnabled)
                        this.btnNew.Focus();
                    else
                        this.btnSave.Focus();
                }
                else
                    SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.Tab);
                e.Handled = true;
            }
            else
            {
                if (e.Key != Key.Left)
                    return;
                if (this.btnNew.IsFocused)
                {
                    this.btnReportFromVoucher.Focus();
                    e.Handled = true;
                }
                else if (this.btnSave.IsFocused)
                {
                    this.btnReportFromVoucher.Focus();
                    e.Handled = true;
                }
                else if (this.btnCancel.IsFocused)
                {
                    this.btnSave.Focus();
                    e.Handled = true;
                }
                else
                {
                    if (!this.btnReportFromVoucher.IsFocused || !this.btnCancel.IsEnabled)
                        return;
                    this.btnCancel.Focus();
                    e.Handled = true;
                }
            }
        }

        private void btnVoucherBookmarkMenu_Click(object sender, RoutedEventArgs e)
        {
            if (this.Cm_BookMark != null)
                this.Cm_BookMark();
            this.LoadBookmark();
            this.popBookmark.IsOpen = true;
        }

        private void mnuAdd_Click(object sender, RoutedEventArgs e)
        {
            string str1 = "";
            if (this.VoucherRow.Table.Columns.Contains("ngay_ct"))
                str1 = str1 + "[" + string.Format("{0:dd/MM/yyyy}", (object)Convert.ToDateTime(this.VoucherRow["ngay_ct"])) + "]";
            string str2 = str1 + " - ";
            if (this.VoucherRow.Table.Columns.Contains("so_ct"))
                str2 = str2 + "[" + this.VoucherRow["so_ct"].ToString().Trim() + "]";
            string str3 = str2 + " - ";
            if (this.VoucherRow.Table.Columns.Contains("dien_giai"))
                str3 += this.VoucherRow["dien_giai"].ToString().Trim();
            this.txtGhi_chu.Text = str3;
            this.popAddBookmark.IsOpen = true;
        }

        private void LoadBookmark()
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("");
                sqlcmd.CommandText = "SELECT CAST(0 AS bit) AS active, * FROM bookmarks WHERE user_id = @user_id AND ma_ct LIKE @ma_ct ORDER BY id DESC";
                sqlcmd.Parameters.Add(new SqlParameter("@ma_ct", SqlDbType.Char)).Value = (object)StartUpTrans.Ma_ct;
                sqlcmd.Parameters.Add(new SqlParameter("@user_id", SqlDbType.Decimal)).Value = (object)Convert.ToInt16(this.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                this.GrdBookmark.ItemsSource = (IEnumerable)this.SasObj.ExcuteReader(sqlcmd).Tables[0].DefaultView;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void btnVoucherBookmarkMenu_LostFocus(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement != null && (ControlFunction.IsRelated((DependencyObject)this.btnVoucherBookmarkMenu, Keyboard.FocusedElement as DependencyObject) || ControlFunction.IsRelated((DependencyObject)this.GrdPopup, Keyboard.FocusedElement as DependencyObject)))
                return;
            this.popBookmark.IsOpen = false;
        }

        private void btnDeleteBookmark_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRow[] dataRowArray = (this.GrdBookmark.ItemsSource as DataView).Table.Select("active = 1");
                if (dataRowArray.Length <= 0)
                    return;
                string str = "";
                for (int index = 0; index < dataRowArray.Length; ++index)
                {
                    if (!string.IsNullOrEmpty(str))
                        str += ",";
                    str += dataRowArray[index]["id"].ToString();
                }
                SqlCommand sqlcmd = new SqlCommand("");
                sqlcmd.CommandText = "DELETE FROM bookmarks WHERE user_id = @user_id AND ma_ct LIKE @ma_ct AND id IN (" + str + ")";
                sqlcmd.Parameters.Add(new SqlParameter("@ma_ct", SqlDbType.Char)).Value = (object)this.VoucherRow["ma_ct"].ToString();
                sqlcmd.Parameters.Add(new SqlParameter("@user_id", SqlDbType.Decimal)).Value = (object)Convert.ToInt16(this.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                if (this.SasObj.ExcuteNonQuery(sqlcmd) <= 0)
                    return;
                this.LoadBookmark();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void btnViewBookmark_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataRow[] dataRowArray = (this.GrdBookmark.ItemsSource as DataView).Table.Select("active = 1");
                if (dataRowArray.Length <= 0)
                    return;
                for (int index = 0; index < dataRowArray.Length; ++index)
                {
                    DataRow dataRow = dataRowArray[index];
                    string Ma_ct = dataRow["ma_ct"].ToString();
                    string Stt_rec = dataRow["stt_rec"].ToString();
                    if (!string.IsNullOrEmpty(Stt_rec.Trim()))
                        SysFunc.EditVoucherFromBrowse(this.SasObj, Ma_ct, Stt_rec, this.AppPath, this.SasObj.M_ProcessName);
                }
                this.popBookmark.IsOpen = false;
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void GridMain_OnOk(object sender, RoutedEventArgs e)
        {
            this.AddBookmark(this.txtGhi_chu.Text.Trim());
        }

        public void AddBookmark(string ghi_chu)
        {
            try
            {
                SqlCommand sqlcmd = new SqlCommand("");
                sqlcmd.CommandText = "INSERT INTO bookmarks(stt_rec,ma_ct,user_id,dien_giai) VALUES(@stt_rec,@ma_ct,@user_id,@dien_giai)";
                sqlcmd.Parameters.Add(new SqlParameter("@stt_rec", SqlDbType.Char)).Value = (object)this.VoucherRow["stt_rec"].ToString();
                sqlcmd.Parameters.Add(new SqlParameter("@ma_ct", SqlDbType.Char)).Value = (object)this.VoucherRow["ma_ct"].ToString();
                sqlcmd.Parameters.Add(new SqlParameter("@user_id", SqlDbType.Decimal)).Value = (object)Convert.ToInt16(this.SasObj.UserInfo.Rows[0]["user_id"].ToString());
                sqlcmd.Parameters.Add(new SqlParameter("@dien_giai", SqlDbType.NVarChar)).Value = (object)ghi_chu;
                if (this.SasObj.ExcuteNonQuery(sqlcmd) == 0)
                {
                    int num = (int)ExMessageBox.Show(-2680, StartupBase.SasObj, "Không đánh dấu được chứng từ hiện thời!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
                else
                {
                    this.LoadBookmark();
                    this.popAddBookmark.IsOpen = false;
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void txtGhi_chu_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.Return) || !Keyboard.IsKeyDown(Key.LeftAlt) && !Keyboard.IsKeyDown(Key.RightAlt))
                return;
            TextBox textBox = sender as TextBox;
            textBox.SelectedText = Environment.NewLine;
            ++textBox.SelectionStart;
            textBox.SelectionLength = 1;
            e.Handled = true;
        }

        private void GridMain_OnCancel(object sender, RoutedEventArgs e)
        {
            this.popAddBookmark.IsOpen = false;
        }

        private void btnTemplate_Click(object sender, RoutedEventArgs e)
        {
            this.LoadTemplate();
            this.popTemplate.IsOpen = true;
        }

        private void LoadTemplate()
        {
            if (this.dsTemplate == null || this.dsTemplate.Tables.Count <= 0)
                return;
            //if (this.dsTemplate.Tables[0].Rows.Count == 0 || !this.dsTemplate.Tables[0].Rows[this.dsTemplate.Tables[0].Rows.Count - 1]["file_name"].Equals("Ngầm định1234598765"))
            //{
            //    DataRow row = this.dsTemplate.Tables[0].NewRow();
            //    row["stt_mau"] = "-1";
            //    row["file_name"] = "Ngầm định1234598765";
            //    row["dien_giai"] = ("Mẫu " + (this.dsTemplate.Tables[0].Rows.Count + 1).ToString() + ": đầy đủ");
            //    row["dien_giai2"] = "Default";
            //    this.dsTemplate.Tables[0].Rows.Add(row);
            //}
            DataRow[] rowmau = this.dsTemplate.Tables[0].Select("stt_mau = -1 and file_name = 'NULL'");
            if(rowmau != null && rowmau.Length > 0)
            {
                rowmau[0]["stt_mau"] = "-1";
                rowmau[0]["file_name"] = "Ngầm định1234598765";
            }
            this.LstTemplate.ItemsSource = (IEnumerable)this.dsTemplate.Tables[0].DefaultView;
            if (this.dsTemplate.Tables[1].Rows.Count == 1)
            {
                DataRow[] dataRowArray = this.dsTemplate.Tables[0].Select("stt_mau = " + this.dsTemplate.Tables[1].Rows[0]["stt_mau"].ToString());
                if (dataRowArray.Length == 1)
                {
                    int num = this.dsTemplate.Tables[0].Rows.IndexOf(dataRowArray[0]);
                    if (num >= 0)
                        this.LstTemplate.SelectedIndex = num;
                }
                else
                    this.LstTemplate.SelectedIndex = 0;//this.dsTemplate.Tables[0].Rows.Count - 1;
            }
            else
            {
                int num = 0;
                try
                {
                    num = this.dsTemplate.Tables[0].Rows.Count - 1;
                    if (num > 0)
                    {
                        for (int index = 0; index < this.dsTemplate.Tables[0].Rows.Count; ++index)
                        {
                            if (this.dsTemplate.Tables[0].Rows[index]["default"].ToString().Trim() == "1")
                            {
                                num = index;
                                break;
                            }
                        }
                    }
                }
                catch
                {
                }
                this.LstTemplate.SelectedIndex = num;
            }
            this.currentTemplate = this.LstTemplate.SelectedIndex;
        }

        private void btnTemplate_LostFocus(object sender, RoutedEventArgs e)
        {
            if (ControlFunction.IsRelated((DependencyObject)this.btnTemplate, Keyboard.FocusedElement as DependencyObject) || ControlFunction.IsRelated((DependencyObject)this.GrdTemplate, Keyboard.FocusedElement as DependencyObject))
                return;
            this.popTemplate.IsOpen = false;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            this.LstTemplate.SelectedIndex = this.LstTemplate.ItemContainerGenerator.IndexFromContainer((DependencyObject)(ControlFunction.GetParentControl((sender as CheckBox).Parent, typeof(ListBoxItem)) as ListBoxItem));
            if (this.SasObj == null || this.currentTemplate == this.LstTemplate.SelectedIndex || (this.LstTemplate.SelectedIndex == -1 || string.IsNullOrEmpty(this.LanguageID)))
                return;
            //if (this.LstTemplate.SelectedIndex != this.dsTemplate.Tables[0].Rows.Count)
            //{
                DataRow row = this.dsTemplate.Tables[0].Rows[this.LstTemplate.SelectedIndex];
                SqlCommand sqlcmd = new SqlCommand();
                sqlcmd.CommandText = "DELETE FROM template_user WHERE user_name LIKE @user_nameDelete AND languageid LIKE @languageidDelete ; INSERT INTO template_user(user_name,languageid,stt_mau) VALUES(@user_name,@languageid,@stt_mau)";
                sqlcmd.Parameters.Add(new SqlParameter("@user_nameDelete", SqlDbType.Char)).Value = this.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                sqlcmd.Parameters.Add(new SqlParameter("@languageidDelete", SqlDbType.NVarChar)).Value = this.LanguageID;
                sqlcmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.Char)).Value = this.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                sqlcmd.Parameters.Add(new SqlParameter("@languageid", SqlDbType.NVarChar)).Value = this.LanguageID;
                sqlcmd.Parameters.Add(new SqlParameter("@stt_mau", SqlDbType.Int)).Value = row["stt_mau"];
                this.SasObj.ExcuteNonQuery(sqlcmd);
                DependencyObject parentControl = ControlFunction.GetParentControl((DependencyObject)this, typeof(FormTrans));
                if (parentControl == null)
                    return;
                string[] parameters = new string[1]
                {
                    StartupBase.Menu_Id
                };
                SysFunc.CallModule(StartupBase.Namespace + ".exe", parameters, this.AppPath, this.SasObj.M_ProcessName, this.SasObj);
                (parentControl as Window).Close();
          //  }
          //  else
          //  {
          //      DataRow row = this.dsTemplate.Tables[0].Rows[this.LstTemplate.SelectedIndex];
          //      SqlCommand sqlcmd = new SqlCommand();
          //      sqlcmd.CommandText = "DELETE FROM template_user WHERE user_name LIKE @user_nameDelete AND languageid LIKE @languageidDelete";
          //      sqlcmd.Parameters.Add(new SqlParameter("@user_nameDelete", SqlDbType.Char)).Value = this.SasObj.UserInfo.Rows[0]["user_name"].ToString();
          //      sqlcmd.Parameters.Add(new SqlParameter("@languageidDelete", SqlDbType.NVarChar)).Value = this.LanguageID;
          //      this.SasObj.ExcuteNonQuery(sqlcmd);
          //      DependencyObject parentControl = ControlFunction.GetParentControl((DependencyObject)this, typeof(FormTrans));
          //      if (parentControl == null)
          //          return;
          //      string[] parameters = new string[1]
          //      {
          //StartupBase.Menu_Id
          //      };
          //      SysFunc.CallModule(StartupBase.Namespace + ".exe", parameters, this.AppPath, this.SasObj.M_ProcessName, this.SasObj);
          //      (parentControl as Window).Close();
          //  }
        }

        private void LstTemplate_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Mouse.LeftButton == MouseButtonState.Pressed)
            {
                CheckBox child = VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.LstTemplate.ItemContainerGenerator.ContainerFromIndex(this.LstTemplate.SelectedIndex), 0), 0), 0), 0) as CheckBox;
                bool? isChecked = child.IsChecked;
                if ((isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) == 0)
                    return;
                child.IsChecked = new bool?(true);
            }
            else
            {
                if (Mouse.RightButton != MouseButtonState.Pressed)
                    return;
                if (VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.LstTemplate.ItemContainerGenerator.ContainerFromIndex(this.LstTemplate.SelectedIndex), 0), 0), 0), 1) is Popup child)
                    child.IsOpen = true;
                for (int index = 0; index < this.LstTemplate.Items.Count; ++index)
                {
                    if (index != this.LstTemplate.SelectedIndex && VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(VisualTreeHelper.GetChild(this.LstTemplate.ItemContainerGenerator.ContainerFromIndex(index), 0), 0), 0), 1) is Popup child1)
                        child1.IsOpen = false;
                }
            }
        }

        private void ExButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsAdmin == 0)
            {
                int num = (int)ExMessageBox.Show(-9680, StartupBase.SasObj, "Bạn không phải là người quản trị!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            if (ExMessageBox.Show(-2685, this.SasObj, "Bạn có muốn xóa mẫu hiện thời không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes || this.SasObj == null || (this.LstTemplate.SelectedIndex == -1 || string.IsNullOrEmpty(this.LanguageID)) || this.LstTemplate.SelectedIndex == 0) //this.LstTemplate.Items.Count - 1
                return;
            DataRowView dataRowView = this.LstTemplate.Items[this.LstTemplate.SelectedIndex] as DataRowView;
            try
            {
                SqlCommand sqlcmd1 = new SqlCommand();
                sqlcmd1.CommandText = "SELECT COUNT(1) FROM template_user WHERE user_name LIKE @user_name AND languageid LIKE @languageid AND stt_mau = @stt_mau";
                sqlcmd1.Parameters.Add(new SqlParameter("@user_name", SqlDbType.Char)).Value = this.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                sqlcmd1.Parameters.Add(new SqlParameter("@languageid", SqlDbType.NVarChar)).Value = this.LanguageID;
                sqlcmd1.Parameters.Add(new SqlParameter("@stt_mau", SqlDbType.Int)).Value = dataRowView["stt_mau"];
                if (Convert.ToInt16(this.SasObj.ExcuteScalar(sqlcmd1)) != 0)
                    return;
                SqlCommand sqlcmd2 = new SqlCommand();
                sqlcmd2.CommandText = "DELETE FROM template_user WHERE user_name LIKE @user_name AND languageid LIKE @languageid0 AND stt_mau = @stt_mau0; DELETE FROM dmtemplate WHERE languageid LIKE @languageid1 AND stt_mau = @stt_mau1; DELETE FROM dmfile WHERE id =@id";
                sqlcmd2.Parameters.Add(new SqlParameter("@user_name", SqlDbType.Char)).Value = this.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                sqlcmd2.Parameters.Add(new SqlParameter("@languageid0", SqlDbType.NVarChar)).Value = this.LanguageID;
                sqlcmd2.Parameters.Add(new SqlParameter("@stt_mau0", SqlDbType.Int)).Value = dataRowView["stt_mau"];
                sqlcmd2.Parameters.Add(new SqlParameter("@languageid1", SqlDbType.NVarChar)).Value = this.LanguageID;
                sqlcmd2.Parameters.Add(new SqlParameter("@stt_mau1", SqlDbType.Int)).Value = dataRowView["stt_mau"];
                sqlcmd2.Parameters.Add(new SqlParameter("@id", SqlDbType.Int)).Value = dataRowView["id"];
                this.SasObj.ExcuteNonQuery(sqlcmd2);
                this.dsTemplate.Tables[0].Rows.Remove(dataRowView.Row);
                this.dsTemplate.Tables[0].AcceptChanges();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void cb_hd_thue_Click(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.DsTrans == null)
                return;
            string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString().Trim();
            string str2 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim();
            SqlCommand sqlcmd = new SqlCommand("Exec [dbo].[UpdateCheckVoucher] @stt_rec, @ma_ct, @ischeck");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = (object)str1;
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)str2;
            bool? isChecked = this.cb_hd_thue.IsChecked;
            if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
            {
                StartUpTrans.hd_thue = "1";
                sqlcmd.Parameters.Add("@ischeck", SqlDbType.Int).Value = (object)1;
                if (StartUpTrans.DsTrans != null)
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["hd_thue"] = (object)"1";
            }
            else
            {
                StartUpTrans.hd_thue = "0";
                sqlcmd.Parameters.Add("@ischeck", SqlDbType.Int).Value = (object)0;
                if (StartUpTrans.DsTrans != null)
                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["hd_thue"] = (object)"0";
            }
            StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
        }

        private void popTemplate_Opened(object sender, EventArgs e)
        {
            if (this.FormParent == null)
                return;
            this.FormParent.EscToClose = false;
        }

        private void popTemplate_Closed(object sender, EventArgs e)
        {
            if (this.FormParent == null)
                return;
            this.FormParent.EscToClose = true;
        }

        public void CloseTemplate()
        {
            if (!this.popTemplate.IsOpen)
                return;
            this.popTemplate.IsOpen = false;
        }

        public delegate void Execute();

        private void ExPhanQuyen_Click(object sender, RoutedEventArgs e)
        {
            if (this.IsAdmin == 0)
            {
                int num = (int)ExMessageBox.Show(-9680, StartupBase.SasObj, "Bạn không phải là người quản trị!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }    
            DataRowView dataRowView = this.LstTemplate.Items[this.LstTemplate.SelectedIndex] as DataRowView;
            try
            {
                string user_right = dataRowView["user_right"].ToString().Trim();
                DanhSachUsers danhSachUsers = new DanhSachUsers();
                danhSachUsers.struserid = user_right.Trim();
                danhSachUsers.Title = StartUpTrans.M_LAN.Equals("V") ? ("Phan quyen: " + SysFunc.Cat_Dau(dataRowView["dien_giai"].ToString().Trim())) : ("Decentralization: " + SysFunc.Cat_Dau(dataRowView["dien_giai2"].ToString().Trim()));
                bool? nullable = danhSachUsers.ShowDialog();
                if ((!nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                    return;

                user_right = danhSachUsers.struserid.Trim();                

                SqlCommand sqlcmd1 = new SqlCommand();
                sqlcmd1.CommandText = "Update dmtemplate SET user_right = @user_right WHERE stt_mau = @stt_mau";
                sqlcmd1.Parameters.Add(new SqlParameter("@user_right", SqlDbType.Char)).Value = user_right.ToString().Trim();
                sqlcmd1.Parameters.Add(new SqlParameter("@stt_mau", SqlDbType.Int)).Value = dataRowView["stt_mau"];
                if(this.SasObj.ExcuteNonQuery(sqlcmd1).ToString().Trim() != "1")
                 return;
                dataRowView["user_right"] = user_right;
                this.dsTemplate.Tables[0].AcceptChanges();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }
    }
}
