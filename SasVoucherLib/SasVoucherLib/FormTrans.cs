using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Infragistics.Windows.Editors;
using Microsoft.Win32;
using SasControls;
using SasControls.ControlLib;
using SasDefine;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace SasVoucherLib
{
    public class FormTrans : Form
    {
        private static readonly string M_MA_CT_KHTMP = "D01;D02;D03;D04;D05;D06;D07;D08;D09;D11;D12;D13;D14;D15;D15;D16;D17;D18;D19;D20;KH1;KH2;TH1;TH2;KD1;D21;D22;D23;D24;D25;D26;D27;D28;D29;D30;D31;D32;D33;D34;D35;D36;D37;D38;D39;D40;D41;D42;D43;D44"; //Khong kiem tra don vi co so thuoc don vi hien thoi cho thực hiện sửa, xóa
        public static readonly DependencyProperty isCurrentFormActiveProperty = DependencyProperty.Register(nameof(isCurrentFormActive), typeof(bool), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)false));
        public static readonly DependencyProperty IsEditModeProperty = DependencyProperty.Register(nameof(IsEditMode), typeof(bool), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)false));
        public static readonly DependencyProperty IsNd51Property = DependencyProperty.Register(nameof(IsNd51), typeof(bool), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)false));
        public static readonly DependencyProperty GridCtProperty = DependencyProperty.Register(nameof(GridCt), typeof(List<BasicGridView>), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty GridCtgtProperty = DependencyProperty.Register(nameof(GridCtgt), typeof(List<BasicGridView>), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty IsMau_tu_inProperty = DependencyProperty.Register(nameof(IsMau_tu_in), typeof(bool), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)false));
        /// <summary>SasObject dùng để thao tác dự liệu cho phiếu.</summary>
        public static SasObject SasO = (SasObject)null;
        public static readonly DependencyProperty ToolbarProperty = DependencyProperty.Register(nameof(Toolbar), typeof(ToolBarControl), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty VCContextMenuProperty = DependencyProperty.Register(nameof(VCContextMenu), typeof(VoucherContextMenuMouseRightClick), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty DK_MA_DVCSProperty = DependencyProperty.Register(nameof(DK_MA_DVCS), typeof(string), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)""));
        public static readonly DependencyProperty M_LANProperty = DependencyProperty.Register(nameof(M_LAN), typeof(string), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)string.Empty));
        public static readonly DependencyProperty StrFilterQSProperty = DependencyProperty.Register(nameof(StrFilterQS), typeof(string), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)string.Empty, new PropertyChangedCallback(FormTrans.StrFilterQSPropertyChanged)));
        public static readonly DependencyProperty auQSProperty = DependencyProperty.Register("auQS", typeof(AutoCompleteTextBox), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty tSo_ctProperty = DependencyProperty.Register("tSo_ct", typeof(TextBox), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty dNgayHTProperty = DependencyProperty.Register("dNgayHT", typeof(DateTextBox), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty C_Ma_ntProperty = DependencyProperty.Register(nameof(C_Ma_nt), typeof(AutoCompleteTextBox), typeof(FormTrans), (PropertyMetadata)new UIPropertyMetadata((object)null, new PropertyChangedCallback(FormTrans.OnC_Ma_ntChanged)));
        public static ActionTask currActionTask = ActionTask.View;
        private bool canclose = true;
        private bool isMainWindow = true;
        public int stt_mau_temlate = -1;
        private int _containma_td4 = -1;
        private int _containma_vv = -1;
        private string _filterAuto;
        private string _filterPerQS;
        private string _filterNgayQS;
        private ContextMenu cm;
        private IEnumerable<SasDefine.Command> ListCommand;
        public bool IsClosing;
        protected bool IsSequenceSave;
        private DataRow _drTemplate;
        private DataSet DsBackup;

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            this.Dispatcher.BeginInvoke((Delegate)new Action(() => this.canclose = string.IsNullOrEmpty(StartUpTrans.Editing_Stt_Rec)), DispatcherPriority.Background);
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            this.canclose = false;
        }

        public bool isCurrentFormActive
        {
            get
            {
                return (bool)this.GetValue(FormTrans.isCurrentFormActiveProperty);
            }
            set
            {
                this.SetValue(FormTrans.isCurrentFormActiveProperty, (object)value);
            }
        }

        /// <summary>Cờ trạng thái editmode hay không của phiếu.</summary>
        public bool IsEditMode
        {
            get
            {
                return (bool)this.GetValue(FormTrans.IsEditModeProperty);
            }
            set
            {
                this.SetValue(FormTrans.IsEditModeProperty, (object)value);
            }
        }

        /// <summary>Cờ trạng thái hd nd51.</summary>
        public bool IsNd51
        {
            get
            {
                return (bool)this.GetValue(FormTrans.IsNd51Property);
            }
            set
            {
                this.SetValue(FormTrans.IsNd51Property, (object)value);
            }
        }

        public List<BasicGridView> GridCt
        {
            get
            {
                return (List<BasicGridView>)this.GetValue(FormTrans.GridCtProperty);
            }
            set
            {
                this.SetValue(FormTrans.GridCtProperty, (object)value);
            }
        }

        public List<BasicGridView> GridCtgt
        {
            get
            {
                return (List<BasicGridView>)this.GetValue(FormTrans.GridCtgtProperty);
            }
            set
            {
                this.SetValue(FormTrans.GridCtgtProperty, (object)value);
            }
        }

        public bool IsMau_tu_in
        {
            get
            {
                return (bool)this.GetValue(FormTrans.IsMau_tu_inProperty);
            }
            set
            {
                this.SetValue(FormTrans.IsMau_tu_inProperty, (object)value);
            }
        }

        /// <summary>Toolbar của phiếu.</summary>
        public ToolBarControl Toolbar
        {
            get
            {
                return (ToolBarControl)this.GetValue(FormTrans.ToolbarProperty);
            }
            set
            {
                this.SetValue(FormTrans.ToolbarProperty, (object)value);
            }
        }

        public VoucherContextMenuMouseRightClick VCContextMenu
        {
            get
            {
                return (VoucherContextMenuMouseRightClick)this.GetValue(FormTrans.VCContextMenuProperty);
            }
            set
            {
                this.SetValue(FormTrans.VCContextMenuProperty, (object)value);
            }
        }

        /// <summary>
        /// DK lọc theo ma_dvcs hiện hành cho các phieu nhap lieu có mã kho
        /// </summary>
        public string DK_MA_DVCS
        {
            get
            {
                return (string)this.GetValue(FormTrans.DK_MA_DVCSProperty);
            }
            set
            {
                this.SetValue(FormTrans.DK_MA_DVCSProperty, (object)value);
            }
        }

        /// <summary>Ngôn ngữ đang dùng hiện tại của hệ thống.</summary>
        public string M_LAN
        {
            get
            {
                return (string)this.GetValue(FormTrans.M_LANProperty);
            }
            set
            {
                this.SetValue(FormTrans.M_LANProperty, (object)value);
            }
        }

        public string StrFilterQS
        {
            get
            {
                return (string)this.GetValue(FormTrans.StrFilterQSProperty);
            }
            set
            {
                this.SetValue(FormTrans.StrFilterQSProperty, (object)value);
            }
        }

        private static void StrFilterQSPropertyChanged(
          DependencyObject d,
          DependencyPropertyChangedEventArgs e)
        {
            if (!(d is FormTrans formTrans) || formTrans.C_QS == null)
                return;
            formTrans.C_QS.Filter = formTrans._filterAuto.Trim() + (formTrans._filterAuto.Trim().Equals("") ? "(" : " AND (") + formTrans.StrFilterQS + ")";
        }

        private string FilterPerQS
        {
            get
            {
                return this._filterPerQS;
            }
            set
            {
                this._filterPerQS = value;
                this.StrFilterQS = this._filterPerQS + (string.IsNullOrEmpty(this._filterNgayQS) ? "" : " AND ") + this._filterNgayQS;
            }
        }

        private string FilterNgayQS
        {
            get
            {
                return this._filterNgayQS;
            }
            set
            {
                this._filterNgayQS = value;
                this.StrFilterQS = this._filterPerQS + (string.IsNullOrEmpty(this._filterNgayQS) ? "" : " AND ") + this._filterNgayQS;
            }
        }

        /// <summary>Button qs để xét quyền và ngày sử dụng qs.</summary>
        public AutoCompleteTextBox C_QS
        {
            get
            {
                return (AutoCompleteTextBox)this.GetValue(FormTrans.auQSProperty);
            }
            set
            {
                this.SetValue(FormTrans.auQSProperty, (object)value);
                if (value == null)
                    return;
                this._filterAuto = this.C_QS.Filter;
                StartUpTrans.M_trung_so = "0";
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (this.C_QS.RowResult == null)
                       this.C_QS.SearchInit();
                   if (this.C_QS.RowResult == null)
                       return;
                   StartUpTrans.M_trung_so = this.C_QS.RowResult["chkso_ct"].ToString();
               }));
                this.C_QS.TextChanged += (AutoCompleteTextBox.Text_Changed)((s, e) =>
               {
                   this.CheckND51();
                   this.C_QS_Text_Changed();
                   StartUpTrans.M_trung_so = "0";
                   if (this.C_QS.RowResult == null)
                       this.C_QS.SearchInit();
                   if (this.C_QS.RowResult == null)
                       return;
                   StartUpTrans.M_trung_so = this.C_QS.RowResult["chkso_ct"].ToString();
                   if (!StartUpTrans.DsTrans.Tables[0].Columns.Contains("kh_mau_hd") || !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["kh_mau_hd"].ToString().Trim()))
                       return;
                   if (string.IsNullOrEmpty(this.C_QS.RowResult["kh_mau_hd"].ToString().Trim()))
                       StartUpTrans.DsTrans.Tables[0].DefaultView[0]["kh_mau_hd"] = (object)this.C_QS.RowResult["mau_hd"].ToString();
                   else
                       StartUpTrans.DsTrans.Tables[0].DefaultView[0]["kh_mau_hd"] = (object)this.C_QS.RowResult["kh_mau_hd"].ToString();
               });
            }
        }

        private void CheckND51()
        {
            this.IsMau_tu_in = false;
                return;
        }

        private void C_QS_Text_Changed()
        {
            if (this.C_NgayHT == null || this.C_QS == null || this.C_So_ct == null)
                return;
            if (!this.IsNd51)
            {
                if (this.Toolbar.IsInEditMode && (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy || FormTrans.currActionTask == ActionTask.Edit))
                {
                    this.C_NgayHT.IsReadOnly = false;
                    this.C_NgayHT.IsTabStop = true;
                    this.C_QS.IsReadOnly = false;
                    this.C_QS.IsTabStop = true;
                    this.C_So_ct.IsReadOnly = false;
                    this.C_So_ct.IsTabStop = true;
                    if (!Keyboard.IsKeyDown(Key.Return) || !this.C_QS.IsFocus)
                        return;
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_So_ct.Focus()));
                }
                else
                {
                    this.C_NgayHT.IsReadOnly = true;
                    this.C_NgayHT.IsTabStop = false;
                    this.C_QS.IsReadOnly = true;
                    this.C_QS.IsTabStop = false;
                    this.C_So_ct.IsReadOnly = true;
                    this.C_So_ct.IsTabStop = false;
                }
            }
            else
            {
                if (this.C_NgayHT == null || this.C_QS == null || this.C_So_ct == null)
                    return;
                if (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy)
                {
                    this.C_NgayHT.IsReadOnly = false;
                    this.C_NgayHT.IsTabStop = true;
                    this.C_QS.IsReadOnly = false;
                    this.C_QS.IsTabStop = true;
                }
                else
                {
                    this.C_NgayHT.IsReadOnly = true;
                    this.C_NgayHT.IsTabStop = false;
                    this.C_QS.IsReadOnly = true;
                    this.C_QS.IsTabStop = false;
                }
                this.C_So_ct.IsReadOnly = true;
                this.C_So_ct.IsTabStop = false;
            }
        }

        public TextBox C_So_ct
        {
            get
            {
                return (TextBox)this.GetValue(FormTrans.tSo_ctProperty);
            }
            set
            {
                this.SetValue(FormTrans.tSo_ctProperty, (object)value);
            }
        }

        /// <summary>Button qs để xét quyền và ngày sử dụng qs.</summary>
        public DateTextBox C_NgayHT
        {
            get
            {
                return (DateTextBox)this.GetValue(FormTrans.dNgayHTProperty);
            }
            set
            {
                this.SetValue(FormTrans.dNgayHTProperty, (object)value);
                if (value == null)
                    return;
                this.C_NgayHT.ValueChanged += (RoutedPropertyChangedEventHandler<object>)((s, e) =>
               {
                   if (!(this.C_NgayHT.dValue != new DateTime()))
                       return;
                   this.FilterNgayQS = " ngay_qs1 <= '" + this.C_NgayHT.dValue.ToString("yyyyMMdd") + "'";
               });
            }
        }

        public AutoCompleteTextBox C_Ma_nt
        {
            get
            {
                return (AutoCompleteTextBox)this.GetValue(FormTrans.C_Ma_ntProperty);
            }
            set
            {
                this.SetValue(FormTrans.C_Ma_ntProperty, (object)value);
            }
        }

        private static void OnC_Ma_ntChanged(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {
            if (!(sender is FormTrans formTrans))
                return;
            formTrans.SetRound();
        }

        public StartUpTrans StartUpMain { get; set; }

        /// <summary>Event thêm mới phiếu.</summary>
        public event FormTrans.Execute Cm_Moi;

        /// <summary>Event sửa phiếu.</summary>
        public event FormTrans.Execute Cm_Sua;

        /// <summary>Event đồng ý các thay đổi của phiếu.</summary>
        public event FormTrans.Execute Cm_Nhan;

        /// <summary>Event xóa phiếu.</summary>
        public event FormTrans.Execute Cm_Xoa;

        /// <summary>Event hủy các thay đổi phiếu.</summary>
        public event FormTrans.Execute Cm_Huy;

        /// <summary>Event hủy hóa đơn.</summary>
        public event FormTrans.Execute Cm_HuyHD;

        /// <summary>Event xem danh sách phiếu.</summary>
        public event FormTrans.Execute Cm_Xem;

        /// <summary>Event tìm kiếm phiếu.</summary>
        public event FormTrans.Execute Cm_Tim;

        /// <summary>Event in phiếu.</summary>
        public event FormTrans.Execute Cm_In;

        /// <summary>Event copy phiếu.</summary>
        public event FormTrans.Execute Cm_Copy;

        /// <summary>Event nhảy tới phiếu kế tiếp.</summary>
        public event FormTrans.Execute Cm_Truoc;

        /// <summary>Event quay lui phiếu sau.</summary>
        public event FormTrans.Execute Cm_Sau;

        /// <summary>Event nhảy tới phiếu đầu tiên trong danh sách.</summary>
        public event FormTrans.Execute Cm_Dau;

        /// <summary>Event quay lui phiếu sau cùng trong danh sách.</summary>
        public event FormTrans.Execute Cm_Cuoi;

        /// <summary>Event xử lý khi xử lý toolbar</summary>
        public event FormTrans.OnEditMode EditModeEnded;

        public FormTrans()
        {
            this.Loaded += new RoutedEventHandler(this.FormTrans_Loaded);
            this.Unloaded += new RoutedEventHandler(this.FormTrans_Unloaded);
            this.GridCt = new List<BasicGridView>();
            this.GridCtgt = new List<BasicGridView>();
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.C_QS == null)
                   return;
               if (StartupBase.SasObj.GetOption("M_QL_HD").ToString().Equals("0"))
               {
                   if (!StartUpTrans.DmctInfo["m_dmqs"].ToString().Equals("0"))
                       return;
                   this.C_QS.dmfix = 1;
               }
               else
               {
                   if (!StartupBase.SasObj.GetOption("M_QL_HD").ToString().Equals("1"))
                       return;
                   SqlCommand sqlcmd = new SqlCommand("select ma_ct_qs from dmloaihdthue where ISNULL(ma_ct_qs,'') not like ''");
                   if (StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Select("ma_ct_qs LIKE '%" + StartUpTrans.Ma_ct + "%'").Length != 0 || !StartUpTrans.DmctInfo["m_dmqs"].ToString().Equals("0"))
                       return;
                   this.C_QS.dmfix = 1;
               }
           }));
            if (StartupBase.SasObj == null || !StartupBase.SasObj.GetOption("M_FTRAN_MAXIMIZE").ToString().Trim().Equals("1"))
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.WindowState = WindowState.Maximized));
            double dWidth = this.Width;
            double dHeight = this.Height;
            this.Width = SystemParameters.WorkArea.Width;
            this.Height = SystemParameters.WorkArea.Height;
            this.StateChanged += (EventHandler)((s, e) =>
           {
               this.Width = dWidth;
               this.Height = dHeight;
           });
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            this.IsClosing = true;
            if (this.IsEditMode)
            {
                if (this.Cm_Huy != null)
                {
                    if (Keyboard.IsKeyDown(Key.Escape))
                        e.Cancel = true;
                    this.isCurrentFormActive = true;
                    if (ExMessageBox.Show(-930, FormTrans.SasO, "Hủy bỏ các thay đổi?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                    {
                        if (FormTrans.currActionTask == ActionTask.Edit)
                        {
                            DataProvider.UpdateDataTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", this.DsBackup.Tables[0], "stt_rec;row_id");
                            if (this.DsBackup.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()))
                                DataProvider.UpdateCtTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), this.DsBackup.Tables[1], StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                            if (this.DsBackup.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                DataProvider.UpdateCtTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_ctgtdbf"].ToString(), this.DsBackup.Tables[2], StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                        }
                        this.Cm_Huy();
                        this.C_QS_Text_Changed();
                        if (!this.IsEditMode)
                        {
                            ToolBarButton btnMoi = this.Toolbar.FindName("btnNew") as ToolBarButton;
                            btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => btnMoi.Focus()));
                        }
                        this.SetInvisible();
                        if (this.EditModeEnded != null)
                            this.EditModeEnded((object)this, "btnCancel", (RoutedEventArgs)null);
                    }
                    else
                        e.Cancel = true;
                    this.isCurrentFormActive = false;
                }
            }
            else if (Keyboard.IsKeyDown(Key.Escape) && string.IsNullOrEmpty(StartUpTrans.Editing_Stt_Rec.Trim()) && ExMessageBox.Show(-935, FormTrans.SasO, "Kết thúc cập nhật chứng từ?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                e.Cancel = true;
            if (!string.IsNullOrEmpty(StartUpTrans.Editing_Stt_Rec.Trim()))
                return;
            base.OnClosing(e);
        }

        private void FormTrans_Unloaded(object sender, RoutedEventArgs e)
        {
            this.Toolbar.RemoveHandler(ButtonBase.ClickEvent, (Delegate)new RoutedEventHandler(this.ToolBarButton_Click));
        }

        private GridLayout GetGridMain()
        {
            DependencyObject reference = (DependencyObject)this;
            try
            {
                for (int index = 0; index < 4; ++index)
                    reference = VisualTreeHelper.GetChild(reference, 0);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return reference as GridLayout;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private void FormTrans_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.EscToClose = true;
                this.IsClosing = false;
                if (this.StartUpMain != null && !string.IsNullOrEmpty(StartUpTrans.Editing_Stt_Rec))
                    this.isMainWindow = false;
                FormTrans.SasO = this.BindingSasObj;
                this.SetDK_MA_DVCS();
                GridLayout GridMain = this.GetGridMain();
                this.Toolbar = new ToolBarControl();
                this.Toolbar.Margin = new Thickness(0, 0, 0, 5);
                this.Toolbar.HorizontalAlignment = HorizontalAlignment.Center;
                this.Toolbar.VerticalAlignment = VerticalAlignment.Top;
                //this.Toolbar.Margin = new Thickness(0.0);
                this.Toolbar.Height = 35.0;
                this.Toolbar.vc.SasObj = FormTrans.SasO;
                this.Toolbar.SasObj = FormTrans.SasO;
                this.Toolbar.FormParent = this;
                this.Toolbar.SetBinding(FrameworkElement.WidthProperty, (BindingBase)new Binding("ActualWidth")
                {
                    Source = (object)GridMain,
                    Mode = BindingMode.OneWay
                });
                this.Toolbar.SetBinding(ToolBarControl.IsInEditModeProperty, (BindingBase)new Binding("IsEditMode")
                {
                    Source = (object)this,
                    Mode = BindingMode.OneWay
                });
                this.Toolbar.SetBinding(ToolBarControl.IsNd51Property, (BindingBase)new Binding("IsNd51")
                {
                    Source = (object)this,
                    Mode = BindingMode.OneWay
                });
                this.Toolbar.ListAdd = FormTrans.SasO.UserInfo.Rows[0]["r_add"].ToString().Split('/');
                this.Toolbar.ListEdit = FormTrans.SasO.UserInfo.Rows[0]["r_edit"].ToString().Split('/');
                this.Toolbar.ListDelete = FormTrans.SasO.UserInfo.Rows[0]["r_del"].ToString().Split('/');
                this.Toolbar.ListPrint = FormTrans.SasO.UserInfo.Rows[0]["r_print"].ToString().Split('/');
                this.Toolbar.Menu_id = StartupBase.Menu_Id;
                this.Toolbar.LanguageID = this.LanguageID;
                int.TryParse(FormTrans.SasO.UserInfo.Rows[0]["is_admin"] == DBNull.Value ? "0" : FormTrans.SasO.UserInfo.Rows[0]["is_admin"].ToString(), out this.Toolbar.IsAdmin);
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (GridMain.DataContext == null || !(GridMain.DataContext is DataView)) 
                       return;
                   this.Toolbar.SetBinding(ToolBarControl.DViewSourceProperty, (BindingBase)new Binding("DataContext")
                   {
                       Source = (object)GridMain,
                       Mode = BindingMode.OneWay
                   });
               }));

                int countGridMain = GridMain.RowDefinitions.Count;

                if (GridMain.RowDefinitions[0].Height == new GridLength(35.0))
                    GridMain.RowDefinitions[0].Height = new GridLength(0.0);
               
                GridMain.RowDefinitions.Add(new RowDefinition());
                GridMain.RowDefinitions[countGridMain].Height = new GridLength(40.0);
                GridMain.Children.Add(this.Toolbar);
                Grid.SetRow(this.Toolbar, countGridMain);
                Grid.SetColumn(this.Toolbar, 0);
                
                
                this.VCContextMenu = new VoucherContextMenuMouseRightClick(this.Toolbar);
                this.MouseRightButtonDown += new MouseButtonEventHandler(this.FormTrans_MouseRightButtonDown);
                this.Toolbar.AddHandler(ButtonBase.ClickEvent, (Delegate)new RoutedEventHandler(this.ToolBarButton_Click));
                this.PreviewKeyDown += new KeyEventHandler(this.FormTrans_PreviewKeyDown);
                this.KeyUp += new KeyEventHandler(this.FormTrans_KeyUp);
                this.FilterPerQS = "((EXISTS (SELECT * FROM dmuserqs WHERE [user_id] = " + FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString() + ") AND EXISTS (SELECT * FROM dmuserqs WHERE v_dmqs.ma_qs = dmuserqs.ma_qs AND [user_id] = " + FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString() + "))OR NOT EXISTS (SELECT * FROM dmuserqs WHERE [user_id] =  " + FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString() + ") OR NOT EXISTS(SELECT 1 FROM dmuserqs WHERE v_dmqs.ma_qs = ma_qs))";
                this.Toolbar.Cm_ContextMenu += (ToolBarControl.Execute)(() =>
               {
                   if (this.StartUpMain == null || StartUpTrans.DsTrans == null || (StartUpTrans.DsTrans.Tables.Count <= 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                       return;
                   DataTable table = FormTrans.SasO.ExcuteReader(new SqlCommand("SELECT * from " + StartUpTrans.DmctInfo["m_phdbf"] + " WHERE stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'")).Tables[0];
                   if (table.Rows.Count != 1)
                       return;
                   this.Toolbar.vc.VoucherRow = table.Rows[0];
                   this.Toolbar.vc.DmctRow = StartUpTrans.DmctInfo;
               });
                this.Toolbar.Cm_BookMark += (ToolBarControl.Execute)(() =>
               {
                   if (this.StartUpMain == null || StartUpTrans.DsTrans == null || (StartUpTrans.DsTrans.Tables.Count <= 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                       return;
                   DataTable table = FormTrans.SasO.ExcuteReader(new SqlCommand("SELECT * from " + StartUpTrans.DmctInfo["m_phdbf"] + " WHERE stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'")).Tables[0];
                   if (table.Rows.Count != 1)
                       return;
                   this.Toolbar.VoucherRow = table.Rows[0];
               });
                if (this.BindingSasObj != null)
                {
                    SqlCommand sqlcmd = new SqlCommand();
                    sqlcmd.CommandText = "IF NOT EXISTS(SELECT TOP 1 * FROM dmtemplate WHERE stt_mau = -1 AND languageid LIKE @language_id)";
                    sqlcmd.CommandText += "\n \t INSERT INTO dmtemplate(stt_mau,languageid,dien_giai,dien_giai2,[path],[file_name],[default],user_right) values('-1',@language_id,N'Mẫu : đầy đủ',N'Default','',N'Ngầm định1234598765','0','')";
                    sqlcmd.CommandText += "\n SELECT a.stt_mau, a.dien_giai, a.dien_giai2,a.[default], a.[user_right], b.* FROM dmtemplate a INNER JOIN dmfile b ON a.path = b.path AND a.file_name = b.file_name";
                    if(this.Toolbar.IsAdmin == 1)
                    {
                        sqlcmd.CommandText += " WHERE a.languageid LIKE @languageid ";
                        sqlcmd.CommandText += " union all SELECT a.stt_mau, a.dien_giai, a.dien_giai2,a.[default], a.[user_right], b.* FROM dmtemplate a LEFT JOIN dmfile b ON a.path = b.path AND a.file_name = b.file_name";
                        sqlcmd.CommandText += " WHERE a.languageid LIKE @languageid AND a.stt_mau = -1";
                    }    
                    else
                    {
                        sqlcmd.CommandText += " WHERE a.languageid LIKE @languageid AND " + string.Format("(ISNULL(a.user_right, '') = '' OR dbo.InList2(LTRIM(RTRIM(str({0}))) + ';',a.user_right,';')>0)", this.BindingSasObj.UserInfo.Rows[0]["user_id"].ToString());
                        sqlcmd.CommandText += " union all SELECT a.stt_mau, a.dien_giai, a.dien_giai2,a.[default], a.[user_right], b.* FROM dmtemplate a LEFT JOIN dmfile b ON a.path = b.path AND a.file_name = b.file_name";
                        sqlcmd.CommandText += " WHERE a.languageid LIKE @languageid AND a.stt_mau = -1 AND " + string.Format("(ISNULL(a.user_right, '') = '' OR dbo.InList2(LTRIM(RTRIM(str({0}))) + ';',a.user_right,';')>0)", this.BindingSasObj.UserInfo.Rows[0]["user_id"].ToString());
                    }

                    sqlcmd.CommandText += " ORDER BY stt_mau;";

                    sqlcmd.CommandText += "SELECT * FROM template_user WHERE languageid LIKE @language_id AND user_name LIKE @user_name";
                    sqlcmd.Parameters.Add(new SqlParameter("@languageid", SqlDbType.NVarChar)).Value = this.LanguageID;
                    sqlcmd.Parameters.Add(new SqlParameter("@language_id", SqlDbType.NVarChar)).Value = this.LanguageID;
                    sqlcmd.Parameters.Add(new SqlParameter("@user_name", SqlDbType.Char)).Value = this.BindingSasObj.UserInfo.Rows[0]["user_name"].ToString();
                    DataSet dataSet = this.BindingSasObj.ExcuteReader(sqlcmd);
                    this.Toolbar.dsTemplate = dataSet;
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        DataRow[] dataRowArray = dataSet.Tables[1].Rows.Count != 1 ? dataSet.Tables[0].Select("[default] = 1") : dataSet.Tables[0].Select("stt_mau = " + dataSet.Tables[1].Rows[0]["stt_mau"].ToString());
                        if (dataRowArray.Length == 1)
                        {
                            this.stt_mau_temlate = Convert.ToInt32(dataRowArray[0]["stt_mau"].ToString());
                            this._drTemplate = dataRowArray[0];
                        }
                    }
                }
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   object name = this.FindName("txtStatus");
                   if (name != null)
                       (name as XamComboEditor).GetBindingExpression(XamComboEditor.SelectedIndexProperty).UpdateTarget();
                   if (this.stt_mau_temlate > 0)
                   {
                       this.BindingSasObj.SynchroFile(this._drTemplate["path"].ToString(), this._drTemplate["file_name"].ToString());
                       this.SetTemplate(this._drTemplate, (Form)this);
                   }
                   if (this.GridCt == null || this.GridCt.Count <= 0)
                       return;
                   for (int index = 0; index < this.GridCt.Count; ++index)
                       this.GridCt[index].CellActivating += new EventHandler<CellActivatingEventArgs>(this.GridCt_CellActivating);
                   if (this._containma_td4 != -1)
                       return;
                   this._containma_td4 = 0;
                   this._containma_vv = 0;
                   for (int index = 0; index < this.GridCt[0].FieldLayouts[0].Fields.Count; ++index)
                   {
                       if (this.GridCt[0].FieldLayouts[0].Fields[index].Name.Equals("ma_td4_i"))
                           this._containma_td4 = 1;
                       if (this.GridCt[0].FieldLayouts[0].Fields[index].Name.Equals("ma_vv_i"))
                           this._containma_vv = 1;
                   }
               }));
                if (this.C_Ma_nt != null)
                {
                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.SetRound()));
                    this.C_Ma_nt.PreviewLostFocus += (AutoCompleteTextBox.Lost_Focus)((s, a) =>
                   {
                       this.SetRound();
                       if (this._drTemplate == null || !this.C_Ma_nt.IsDataChanged)
                           return;
                       this.SetTemplate(this._drTemplate, (Form)this);
                   });
                }
                if (FormTrans.SasO.GetRegInfo() != null && !FormTrans.SasO.GetRegInfo().Rows[18]["content"].ToString().Trim().Equals("FK"))
                    this.Title += SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? " - Đơn vị: " + FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString() : " - Unit code: " + FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString());
                this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
               {
                   if (FormTrans.SasO.GetOption("M_VIEW_CK").ToString().Trim() == "1" && this.Cm_Xem != null && (StartUpTrans.DsTrans.Tables.Count > 0 && StartUpTrans.DsTrans.Tables[0].Rows.Count > 1) && string.IsNullOrEmpty(StartUpTrans.Editing_Stt_Rec))
                   {
                       this.Cm_Xem();
                       this.CheckND51();
                       this.C_QS_Text_Changed();
                       this.SetInvisible();
                       if (this.EditModeEnded != null)
                           this.EditModeEnded((object)this, "btnView", e);
                   }
                   this.CheckBoxVoucher();
                   this.ChangeLanguage(this.M_LAN, (object)this.Toolbar);
               }), DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void FormTrans_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (this.IsEditMode)
                return;
            this.VCContextMenu.IsOpen = true;
        }

        private void GridCt_CellActivating(object sender, CellActivatingEventArgs e)
        {
            switch (e.Cell.Field.Name)
            {
                case "ma_td4_i":
                    this.UpdateFilterMa_td4(e.Cell);
                    break;
                case "ma_vv_i":
                    this.UpdateFilterMa_vv(e.Cell);
                    break;
            }
        }

        private void UpdateFilterMa_td4(Cell _cellma_td4_i)
        {
            if (this._containma_td4 != 1 || this._containma_vv != 1)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell(_cellma_td4_i);
               if (cellValuePresenter == null)
                   return;
               DataRecord record = _cellma_td4_i.Record;
               AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(cellValuePresenter.Editor as ControlHostEditor);
               if (autoCompleteControl == null)
                   return;
               autoCompleteControl.Filter = " ISNULL(ma_dvcs,'" + this.BindingSasObj.M_ma_dvcs.Trim() + "') LIKE '" + this.BindingSasObj.M_ma_dvcs.Trim() + "' AND ISNULL(cach_tinh,'') LIKE '' AND ma_vv LIKE '" + record.Cells["ma_vv_i"].Value.ToString().Trim() + "'";
               if (autoCompleteControl.CheckLostFocus())
                   return;
               autoCompleteControl.Text = "";
           }));
        }

        private void UpdateFilterMa_vv(Cell _cellma_vv_i)
        {
            if (this._containma_vv != 1)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               CellValuePresenter cellValuePresenter = CellValuePresenter.FromCell(_cellma_vv_i);
               if (cellValuePresenter == null)
                   return;
               DataRecord record = cellValuePresenter.Record;
               AutoCompleteTextBox autoCompleteControl = ControlFunction.GetAutoCompleteControl(cellValuePresenter.Editor as ControlHostEditor);
               if (autoCompleteControl == null)
                   return;
               autoCompleteControl.Filter = !this.BindingSasObj.GetOption("M_IN_MA_VV_ME").ToString().Equals("0") ? "ISNULL(su_dung,1) != 0" : "NOT EXISTS (SELECT 1 FROM dmvv WHERE " + autoCompleteControl.Listinfo["table_view"].ToString().Trim() + ".ma_vv LIKE dmvv.ma_vv_me) AND ISNULL(su_dung,1) != 0";
               if (autoCompleteControl.CheckLostFocus())
                   return;
               autoCompleteControl.Text = "";
           }));
        }
        public void LoadStatus()
        {            
            if (StartupBase.SasObj.UserInfo.Rows[0]["is_admin"].ToString().Trim() == "1")
                StartUpTrans.tbStatus = DataProvider.FillCommand(StartupBase.SasObj, new SqlCommand("Select * from dmPost where ma_ct like '%" + StartUpTrans.Ma_ct + "%'")).Tables[0];
            else
                StartUpTrans.tbStatus = DataProvider.FillCommand(StartupBase.SasObj, new SqlCommand("Select * from dmPost where ma_ct like '%" + StartUpTrans.Ma_ct + "%' AND (CHARINDEX(';'+'" + StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString().Trim() + "'+';',';' + LTRIM(user_right)+';')>0 OR user_right is null OR user_right = '')")).Tables[0];
        }

        public string GetStatus()
        {
            string newStatus = "";
            if (StartUpTrans.tbStatus.Rows.Count > 0)
            {
                var staquery = from DataRowView rowView in StartUpTrans.tbStatus.DefaultView
                               where rowView.Row.Field<string>("ma_post") == StartUpTrans.DmctInfo["ma_post"].ToString()
                               select rowView;
                if (staquery.Count() > 0)
                    newStatus = StartUpTrans.DmctInfo["ma_post"].ToString();
                else
                    newStatus = StartUpTrans.tbStatus.Rows[0]["ma_post"].ToString();
            }
            else
            {
                return "";
            }    
            return newStatus;
        }
        private void SetDK_MA_DVCS()
        {
            if (FormTrans.SasO == null)
                return;
            this.DK_MA_DVCS = FormTrans.SasO.GetDmdmInfo("dmkho").Rows[0]["filter"].ToString().Trim();
            FormTrans formTrans = this;
            formTrans.DK_MA_DVCS = formTrans.DK_MA_DVCS + (string.IsNullOrEmpty(this.DK_MA_DVCS) ? "" : " AND ") + " ma_dvcs LIKE '" + FormTrans.SasO.M_ma_dvcs.Trim() + "'";
        }

        public void SetRound()
        {
            if (this.C_Ma_nt == null)
                return;
            if (this.C_Ma_nt.Text != StartUpTrans.M_ma_nt0)
            {
                StartUpTrans.M_ROUND = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND"));
                StartUpTrans.M_ROUND_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_NT"));
                StartUpTrans.M_ROUND_GIA = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
                StartUpTrans.M_ROUND_GIA_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA_NT"));
            }
            else
            {
                StartUpTrans.M_ROUND = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND"));
                StartUpTrans.M_ROUND_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND"));
                StartUpTrans.M_ROUND_GIA = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
                StartUpTrans.M_ROUND_GIA_NT = Convert.ToInt32(StartupBase.SasObj.GetSysvar("M_ROUND_GIA"));
            }
        }

        private void SetInvisible()
        {
            if (this.IsNd51)
            {
                string str = StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString();
                if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("3") || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("4") || (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["status"].ToString().Equals("5") || this.IsEditMode) || (StartUpTrans.DsTrans.Tables[0].Rows.Count <= 1 || !str.Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString())))
                {
                    this.Toolbar.btnPrint.IsEnabled = false;
                    this.Toolbar.btnPrint.Opacity = 0.3;
                    this.Toolbar.btnEdit.IsEnabled = false;
                    this.Toolbar.btnEdit.Opacity = 0.3;
                    this.Toolbar.btnDelete.IsEnabled = false;
                    this.Toolbar.btnDelete.Opacity = 0.3;
                }
                else if (!this.IsEditMode)
                {
                    if (((IEnumerable<string>)this.Toolbar.ListPrint).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1)
                    {
                        this.Toolbar.btnPrint.IsEnabled = true;
                        this.Toolbar.btnPrint.Opacity = 1.0;
                    }
                    if (((IEnumerable<string>)this.Toolbar.ListEdit).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1)
                    {
                        this.Toolbar.btnEdit.IsEnabled = true;
                        this.Toolbar.btnEdit.Opacity = 1.0;
                    }
                    if (((IEnumerable<string>)this.Toolbar.ListDelete).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1)
                    {
                        this.Toolbar.btnDelete.IsEnabled = true;
                        this.Toolbar.btnDelete.Opacity = 1.0;
                    }
                }
                this.Toolbar.btnDelete.Text = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString().Equals("0") || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString()) ? (StartupBase.M_LAN.Equals("V") ? "Xóa" : "Delete") : (StartupBase.M_LAN.Equals("V") ? "Xóa bỏ" : "Remove");
            }
            else
                this.Toolbar.btnDelete.Text = StartupBase.M_LAN.Equals("V") ? "Xóa" : "Delete";
            if (StartUpTrans.DsTrans == null || !StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt") || !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("9") && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["tinh_trang_hddt"].ToString().Trim().Equals("4"))
                return;
            this.Toolbar.btnPrint.IsEnabled = false;
            this.Toolbar.btnPrint.Opacity = 0.3;
            this.Toolbar.btnEdit.IsEnabled = false;
            this.Toolbar.btnEdit.Opacity = 0.3;
            this.Toolbar.btnDelete.IsEnabled = false;
            this.Toolbar.btnDelete.Opacity = 0.3;
        }

        private void FormTrans_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                string menuItemName = string.Empty;
                if (e.Key == Key.F7 && this.Cm_In != null && this.Toolbar.btnPrint.IsEnabled)
                {
                    GridLayout gridMain = this.GetGridMain();
                    if (gridMain.DataContext != null)
                    {
                        DataView dataContext = gridMain.DataContext as DataView;
                        string str1 = dataContext[0]["stt_rec"].ToString().ToString();
                        string str2 = dataContext[0]["ma_ct"].ToString().ToString();
                        string cmdText = string.Empty;
                        if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                        {
                            cmdText = "exec LoadVoucher#1 @ma_ct, @PhFilter, @CtFilter, @Sl_ct";
                            if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 1)
                                cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[1]);
                        }
                        else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                        {
                            cmdText = "exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct";
                            if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 0)
                                cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[0]);
                        }
                        SqlCommand cmd = new SqlCommand(cmdText);
                        if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                        {
                            cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                            cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                            cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                            cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                        }
                        else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                        {
                            cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                            cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                            cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                            cmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                            cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                        }
                        if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                        {
                            DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, cmd);
                            if (dataSet.Tables[0].Rows.Count == 1)
                            {
                                StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + str1 + "'")[0].ItemArray = dataSet.Tables[0].Rows[0].ItemArray;
                                if (StartUpTrans.DsTrans.Tables.Count > 1)
                                {
                                    string rowFilter = StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter;
                                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                    foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str1 + "'"))
                                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                                    foreach (DataRow dataRow in dataSet.Tables[1].Select("stt_rec='" + str1 + "'"))
                                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow.ItemArray);
                                    StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = rowFilter;
                                }
                                if (StartUpTrans.DsTrans.Tables.Count > 2)
                                {
                                    try
                                    {
                                        string rowFilter = StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter;
                                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + str1 + "'"))
                                            StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                                        foreach (DataRow dataRow in dataSet.Tables[2].Select("stt_rec='" + str1 + "'"))
                                            StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow.ItemArray);
                                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = rowFilter;
                                    }
                                    catch (Exception ex)
                                    {
                                    }
                                }
                            }
                        }
                    }
                    this.Cm_In();
                    menuItemName = "btnPrint";
                    e.Handled = true;
                    ToolBarButton btnPrint = this.Toolbar.FindName("btnPrint") as ToolBarButton;
                    btnPrint.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => btnPrint.Focus()));
                }
                if (e.Key == Key.Escape)
                    this.Toolbar.CloseTemplate();
                if (!this.isMainWindow && e.Key == Key.Escape && this.canclose)
                {
                    this.Close();
                }
                else
                {
                    if (this.EditModeEnded != null && !string.IsNullOrEmpty(menuItemName))
                    {
                        this.EditModeEnded((object)this, menuItemName, (RoutedEventArgs)e);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                    this.SetInvisible();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        /// <summary>Thiết lập focus cho button sửa hay thêm của toolbar.</summary>
        protected void SetFocusToolbar()
        {
            ToolBarButton btnMoi = this.Toolbar.FindName("btnNew") as ToolBarButton;
            Action action = (Action)(() => btnMoi.Focus());
            btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)action);
        }

        protected Decimal GetCurrentSo_ct(string _format, string _value)
        {
            string str1 = _format.Substring(0, _format.IndexOf("{"));
            string str2 = _format.Substring(_format.IndexOf("}") + 1, _format.Length - _format.IndexOf("}") - 1);
            _value.Substring(0, str1.Length);
            _value.Substring(_value.Length - str2.Length, str2.Length);
            int result = 0;
            return int.TryParse(_value.Substring(str1.Length, _value.Length - str1.Length - str2.Length), out result) ? (Decimal)result : new Decimal(-1);
        }

        protected bool CheckSo_ct(
          string _format,
          Decimal _startNumber,
          Decimal _endNumber,
          string _value)
        {
            if (StartUpTrans.DmctInfo["m_dmqs"].ToString().Equals("0"))
                return true;
            if (this.BindingSasObj != null && StartUpTrans.DsTrans != null && (StartUpTrans.DsTrans.Tables.Count > 0 && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_qs")))
            {
                DataTable table = this.BindingSasObj.ExcuteReader(new SqlCommand(string.Format("SELECT transform, so_ct1, so_ct2 FROM dmqs WHERE ma_qs LIKE '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString()))).Tables[0];
                if (table.Rows.Count == 1)
                {
                    _format = table.Rows[0]["transform"].ToString();
                    _startNumber = Convert.ToDecimal(table.Rows[0]["so_ct1"]);
                    _endNumber = Convert.ToDecimal(table.Rows[0]["so_ct2"]);
                }
            }
            if (_format.Contains("{") && _format.Contains("}"))
            {
                string str1 = _format.Substring(0, _format.IndexOf("{"));
                string str2 = _format.Substring(_format.IndexOf("}") + 1, _format.Length - _format.IndexOf("}") - 1);
                if (_value.Length >= str1.Length + str2.Length)
                {
                    string str3 = _value.Substring(0, str1.Length);
                    string str4 = _value.Substring(_value.Length - str2.Length, str2.Length);
                    if (str1.ToUpper().Trim().Equals(str3.ToUpper().Trim()) && str2.ToUpper().Trim().Equals(str4.ToUpper().Trim()) && _value.Length > str1.Length + str2.Length)
                    {
                        int result = 0;
                        if (int.TryParse(_value.Substring(str1.Length, _value.Length - str1.Length - str2.Length), out result) && _startNumber <= (Decimal)result && (Decimal)result <= _endNumber)
                            return true;
                    }
                }
            }
            return false;
        }

        public bool CheckVoucherOutofDate()
        {
            GridLayout gridMain = this.GetGridMain();
            if (gridMain.DataContext != null)
            {
                DataView dataContext = gridMain.DataContext as DataView;
                if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1 && FormTrans.SasO.GetSysvar("M_LIST_INPUT_VC").ToString().Contains(StartUpTrans.Ma_ct))
                {
                    bool flag1 = true;
                    if (dataContext.Table.Columns.Contains("thue_dau_vao") && !dataContext[0]["thue_dau_vao"].ToString().Equals("1"))
                        flag1 = false;
                    if (dataContext.Table.Columns.Contains("thue_dau_ra") && dataContext[0]["thue_dau_ra"].ToString().Equals("1"))
                        flag1 = false;
                    if (flag1)
                    {
                        DateTime dateTime1 = Convert.ToDateTime(dataContext[0]["ngay_ct"]);
                        int int16 = (int)Convert.ToInt16(FormTrans.SasO.GetOption("M_MONTH_OFD"));
                        DateTime dateTime2 = new DateTime(dateTime1.Year, dateTime1.Month, 1).AddMonths(-int16);
                        if (this.GridCtgt.Count > 0)
                        {
                            bool flag2 = false;
                            foreach (BasicGridView basicGridView in this.GridCtgt)
                            {
                                if (basicGridView.FieldLayouts[0].Fields.Where<Field>((Func<Field, bool>)(p => p.Name == "ngay_ct0")).Count<Field>() == 1)
                                {
                                    for (int index = 0; index < basicGridView.Records.Count; ++index)
                                    {
                                        if ((basicGridView.Records[index] as DataRecord).Cells["ngay_ct0"].Value != DBNull.Value)
                                        {
                                            DateTime dateTime3 = Convert.ToDateTime((basicGridView.Records[index] as DataRecord).Cells["ngay_ct0"].Value);
                                            if (dateTime2 > dateTime3)
                                            {
                                                switch (FormTrans.SasO.GetOption("M_AC_MONTH_OFD").ToString())
                                                {
                                                    case "2":
                                                        int num = (int)ExMessageBox.Show(-966, FormTrans.SasO, "Hóa đơn đã quá hạn [" + int16.ToString() + "] tháng, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                        return false;
                                                    case "1":
                                                        if (ExMessageBox.Show(-967, FormTrans.SasO, "Hóa đơn đã quá hạn [" + int16.ToString() + "] tháng, có lưu hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                                            return false;
                                                        flag2 = true;
                                                        break;
                                                }
                                            }
                                            if (flag2)
                                                break;
                                        }
                                        if (flag2)
                                            break;
                                    }
                                }
                            }
                        }
                        else if (dataContext.Table.Columns.Contains("ngay_ct0") && dataContext[0]["ngay_ct0"] != DBNull.Value)
                        {
                            DateTime dateTime3 = Convert.ToDateTime(dataContext[0]["ngay_ct0"]);
                            if (dateTime2 > dateTime3)
                            {
                                switch (FormTrans.SasO.GetOption("M_AC_MONTH_OFD").ToString())
                                {
                                    case "2":
                                        int num = (int)ExMessageBox.Show(-966, FormTrans.SasO, "Hóa đơn đã quá hạn [" + int16.ToString() + "] tháng, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        return false;
                                    case "1":
                                        if (ExMessageBox.Show(-967, FormTrans.SasO, "Hóa đơn đã quá hạn [" + int16.ToString() + "] tháng, có lưu hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                            return false;
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            return true;
        }

        private void ToolBarButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (e.OriginalSource == null)
                    return;
                ToolBarButton btn = e.OriginalSource as ToolBarButton;
                if (btn == null || !btn.IsFocused)
                    return;
                ToolBarButton btnMoi = this.Toolbar.FindName("btnNew") as ToolBarButton;
                Action action = (Action)(() => btnMoi.Focus());
                switch (btn.Name)
                {
                    case "btnNext":
                        if (this.Cm_Sau != null)
                            this.Cm_Sau();
                        this.CheckBoxVoucher();
                        this.CheckND51();
                        this.SetInvisible();
                        break;
                    case "btnPrevious":
                        if (this.Cm_Truoc != null)
                            this.Cm_Truoc();
                        this.CheckBoxVoucher();
                        this.CheckND51();
                        this.SetInvisible();
                        break;
                    case "btnTop":
                        if (this.Cm_Dau != null)
                            this.Cm_Dau();
                        this.CheckBoxVoucher();
                        this.CheckND51();
                        this.SetInvisible();
                        break;
                    case "btnBottom":
                        if (this.Cm_Cuoi != null)
                            this.Cm_Cuoi();
                        this.CheckBoxVoucher();
                        this.CheckND51();
                        this.SetInvisible();
                        break;
                    case "btnNew":
                        if (this.Cm_Moi != null)
                        {
                            Infragistics.Windows.Controls.XamTabControl TabInfoobj = this.FindName("TabInfo") as Infragistics.Windows.Controls.XamTabControl;
                            if (TabInfoobj != null)
                            {
                                TabInfoobj.SelectedIndex = 0;
                            }
                            this.Cm_Moi();
                            if (this.IsEditMode)
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"];
                                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("loai_tg") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = (object)StartUpTrans.Getloai_tg(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("hd_thue"))
                                {
                                    if (FormTrans.currActionTask != ActionTask.Edit)
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["hd_thue"] = !StartupBase.SasObj.GetSysvar("M_CHECK_VOUCHER").ToString().Trim().Equals("2") || !StartUpTrans.hd_thue.Trim().Equals("1") && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim().ToUpper().Equals("QL1") ? (object)"0" : (object)"1";
                                    this.CheckBoxVoucher();
                                    break;
                                }
                                break;
                            }
                            break;
                        }
                        break;
                    case "btnEdit": 
                        if (this.Cm_Sua != null)
                        {
                            Infragistics.Windows.Controls.XamTabControl TabInfoobj = this.FindName("TabInfo") as Infragistics.Windows.Controls.XamTabControl;
                            if (TabInfoobj != null)
                            {
                                TabInfoobj.SelectedIndex = 0;
                            }
                            GridLayout gridMain = this.GetGridMain();
                            if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt"))
                            {
                                SqlCommand sqlcmd = new SqlCommand("SELECT 1 FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec ='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "' AND (isnull(tinh_trang_hddt,0) != 0 OR tinh_trang_hddt != '')");
                                int count = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count;
                                if (this.CheckDC() || count > 0)
                                {
                                    int num = (int)ExMessageBox.Show(-2016, StartupBase.SasObj, "Đã có phát sinh HĐĐT không được sửa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                            }
                            if (gridMain.DataContext != null)
                            {
                                DataView dataContext = gridMain.DataContext as DataView;
                                if (this.IsNd51)
                                {
                                    string str = StartUpTrans.GetSl_in(dataContext[0]["stt_rec"].ToString().ToString()).ToString();
                                    if (!str.Equals("0"))
                                    {
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
                                        if (ExMessageBox.Show(-940, StartupBase.SasObj, "Hóa đơn đã được in, có muốn sửa lại chứng từ hay không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                            return;
                                    }
                                }
                                else if (!StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString().Trim()) && FormTrans.M_MA_CT_KHTMP.IndexOf(StartUpTrans.Ma_ct) ==-1)
                                {
                                    int num = (int)ExMessageBox.Show(-950, FormTrans.SasO, "Chứng từ không thuộc đơn vị hiện thời!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1 && !SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                {
                                    int num = (int)ExMessageBox.Show(-955, FormTrans.SasO, "Dữ liệu đã khóa sổ, không sửa được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if ((!(StartUpTrans.Ma_ct == "HDA") && !(StartUpTrans.Ma_ct == "HD1") || !(dataContext[0]["stt_rec_pt"].ToString().Trim() != "")) && StartUpTrans.CheckPhanBo(dataContext[0]["stt_rec"].ToString()) == 1)
                                {
                                    int num = (int)ExMessageBox.Show(-960, FormTrans.SasO, "Không được sửa hóa đơn đã được thanh toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if (this.C_QS != null && this.C_NgayHT != null && (!string.IsNullOrEmpty(this.C_QS.Text.Trim()) && !string.IsNullOrEmpty(this.C_NgayHT.Text.Trim())) && StartUpTrans.CheckQS(this.C_QS.Text, this.C_NgayHT.dValue.ToString("yyyyMMdd"), (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString()), this.C_So_ct.Text) == 2)
                                {
                                    int num = (int)ExMessageBox.Show(-981, FormTrans.SasO, "Quyền sử dụng ký hiệu không hợp lệ, không được sửa chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_QS.IsFocus = true));
                                    return;
                                }
                                this.DsBackup = new DataSet();
                                this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                                if (StartUpTrans.DsTrans.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()))
                                    this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                                if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                    this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                                this.Cm_Sua();
                                if (dataContext.Table.Columns.Contains("ngay_ct"))
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();
                                if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_ct0")))
                                {
                                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[2].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                        dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim();
                                    break;
                                }
                                break;
                            }
                            break;
                        }
                        break;
                    case "btnSave":
                        if (!this.IsEditMode)
                            return;
                        if (this.Cm_Nhan != null && btn.IsFocused && StartUpTrans.DsTrans != null)
                        {
                            GridLayout gridMain = this.GetGridMain();
                            DateTime dateTime = (DateTime)StartupBase.SasObj.ExcuteScalar(new SqlCommand("select getdate()"));
                            DateTime date = dateTime.Date;
                            string str1 = (string)StartupBase.SasObj.ExcuteScalar(new SqlCommand("select CONVERT(varchar(8), getdate(), 108)"));
                            if (gridMain.DataContext != null)
                            {
                                DataView dataContext = gridMain.DataContext as DataView;
                                if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1 && dataContext[0]["ngay_ct"] != DBNull.Value)
                                {
                                    if (!SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                    {
                                        int num = (int)ExMessageBox.Show(-965, FormTrans.SasO, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        object obj = this.FindName("txtNgay_ct") ?? this.FindName("txtngay_ct");
                                        if (obj == null)
                                            return;
                                        FrameworkElement dNgay_ct = obj as FrameworkElement;
                                        dNgay_ct.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => dNgay_ct.Focus()));
                                        return;
                                    }
                                    if (!SysFunc.CheckValidNgayMs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"]))))
                                    {
                                        int num = (int)ExMessageBox.Show(-970, FormTrans.SasO, "Ngày hạch toán phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        object obj = this.FindName("txtNgay_ct") ?? this.FindName("txtngay_ct");
                                        if (obj == null)
                                            return;
                                        FrameworkElement dNgay_ct = obj as FrameworkElement;
                                        dNgay_ct.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => dNgay_ct.Focus()));
                                        return;
                                    }
                                }
                                if (this.C_QS != null && this.C_NgayHT != null && (!string.IsNullOrEmpty(this.C_QS.Text.Trim()) && !string.IsNullOrEmpty(this.C_NgayHT.Text.Trim())))
                                {
                                    string text1 = this.C_QS.Text;
                                    dateTime = this.C_NgayHT.dValue;
                                    string ngay_ct = dateTime.ToString("yyyyMMdd");
                                    int int16 = (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString());
                                    string text2 = this.C_So_ct.Text;
                                    switch (StartUpTrans.CheckQS(text1, ngay_ct, int16, text2))
                                    {
                                        case 1:
                                            int num1 = (int)ExMessageBox.Show(-975, FormTrans.SasO, "Ngày bắt đầu sử dụng ký hiệu không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_NgayHT.Focus()));
                                            return;
                                        case 2:
                                            int num2 = (int)ExMessageBox.Show(-980, FormTrans.SasO, "Quyền sử dụng ký hiệu không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_QS.IsFocus = true));
                                            return;
                                        case 3:
                                            if (this.IsNd51)
                                            {
                                                int num3 = (int)ExMessageBox.Show(-985, FormTrans.SasO, "Ngày chứng từ hiện tại nhỏ hơn ngày chứng từ cuối cùng trong 'Ký hiệu'. Không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_NgayHT.Focus()));
                                                return;
                                            }
                                            break;
                                    }
                                }
                                if (dataContext.Table.Columns.Contains("so_ct"))
                                {
                                    string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim());
                                    if (dataContext.Table.Columns.Contains("ma_qs") && !this.IsNd51 && (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim()) && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim())))
                                    {
                                        switch (StartUpTrans.CheckValidSo_ct(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                                        {
                                            case 2:
                                                int num1 = (int)ExMessageBox.Show(-1366, FormTrans.SasO, "Số c.từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                break;
                                            case 3:
                                                int num2 = (int)ExMessageBox.Show(-1366, FormTrans.SasO, "Số c.từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                return;
                                        }
                                    }
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct"), ' ');
                                }
                                if (this.GridCt != null && this.GridCt.Count > 0)
                                {
                                    for (int i = 0; i < this.GridCt.Count; ++i)
                                    {
                                        if (this.GridCt[i].DataSource != null && this.GridCt[i].Visibility == Visibility.Visible)
                                        {
                                            IEnumerable dataSource = this.GridCt[i].DataSource;
                                            int index = -1;
                                            if (!this.CheckExistsCode("ma_vv_i", this.GridCt[i], "ma_vv_yn", out index))
                                            {
                                                int num = (int)ExMessageBox.Show(-1367, FormTrans.SasO, "Mã dự án không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                                this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_vv_i"];
                                                this.GridCt[i].Focus();
                                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                                return;
                                            }
                                            if (!this.CheckExistsCode("ma_px_i", this.GridCt[i], "ma_px_yn", out index))
                                            {
                                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                                int num = (int)ExMessageBox.Show(-1368, FormTrans.SasO, "Mã phân xưởng không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_px_i"];
                                                this.GridCt[i].Focus();
                                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                                return;
                                            }
                                            if (!this.CheckExistsCode("ma_bpht_i", this.GridCt[i], "ma_bpht_yn", out index))
                                            {
                                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                                int num = (int)ExMessageBox.Show(-1369, FormTrans.SasO, "Mã bộ phận hạch toán không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_bpht_i"];
                                                this.GridCt[i].Focus();
                                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                                return;
                                            }
                                            if (!this.CheckExistsCode("ma_phi_i", this.GridCt[i], "ma_phi_yn", out index))
                                            {
                                                SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                                int num = (int)ExMessageBox.Show(-1371, FormTrans.SasO, "Mã phí không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_phi_i"];
                                                this.GridCt[i].Focus();
                                                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                                return;
                                            }
                                        }
                                    }
                                }
                                if (StartUpTrans.DsTrans.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[1].Columns.Contains("so_ct0")))
                                {
                                    bool flag = StartUpTrans.DsTrans.Tables[1].Columns.Contains("so_seri0");
                                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[1].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                    {
                                        dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct0"), ' ');
                                        if (flag)
                                            dataRow["so_seri0"] = (object)dataRow["so_seri0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_seri0"), ' ');
                                    }
                                }
                                if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_ct0")))
                                {
                                    bool flag = StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_seri0");
                                    foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[2].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                    {
                                        dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct0"), ' ');
                                        if (flag)
                                            dataRow["so_seri0"] = (object)dataRow["so_seri0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_seri0"), ' ');
                                    }
                                }
                                if (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy)
                                {
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["date0"] = (object)date;
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["time0"] = (object)str1;
                                    if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("user_name0"))
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                    if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("date0")) && (StartUpTrans.DsTrans.Tables[2].Columns.Contains("time0") && StartUpTrans.DsTrans.Tables[2].Columns.Contains("user_id0")))
                                    {
                                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                        {
                                            dataRowView["date0"] = (object)date;
                                            dataRowView["time0"] = (object)str1;
                                            dataRowView["user_id0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                        }
                                    }
                                }
                                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("user_id"))
                                {
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                    if (dataContext.Table.Columns.Contains("user_name"))
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["date"] = (object)date;
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["time"] = (object)str1;
                                    if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("date")) && (StartUpTrans.DsTrans.Tables[2].Columns.Contains("time") && StartUpTrans.DsTrans.Tables[2].Columns.Contains("user_id")))
                                    {
                                        foreach (DataRowView dataRowView in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                        {
                                            dataRowView["date"] = (object)date;
                                            dataRowView["time"] = (object)str1;
                                            dataRowView["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                        }
                                    }
                                }
                                string str2 = "";
                                if (dataContext.Table.Columns.Contains("ma_qs"))
                                    str2 = dataContext[0]["ma_qs"].ToString().Trim();
                                this.Cm_Nhan();
                                if (!this.IsEditMode && dataContext.Table.Columns.Contains("ma_qs") && StartUpTrans.Ma_ct != "HD1")
                                {
                                    string format = "EXEC  {0} '" + str2 + "', '" + dataContext[0]["so_ct"].ToString().Trim() + "'";
                                    string cmdText = StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 2 ? string.Format(format, (object)"SetSoct") : string.Format(format, (object)StartUpTrans.Process_Store[2]);
                                    FormTrans.SasO.ExcuteNonQuery(new SqlCommand(cmdText));
                                }
                                if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                                {
                                    if (dataContext.Table.Columns.Contains("ten_post"))
                                        dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act"];
                                    if (dataContext.Table.Columns.Contains("ten_post2"))
                                        dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act2"];
                                    if (dataContext.Table.Columns.Contains("ten_act"))
                                        dataContext[0]["ten_act"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act"];
                                    if (dataContext.Table.Columns.Contains("ten_act2"))
                                        dataContext[0]["ten_act2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act2"];
                                }
                            }
                        }
                        else if (this.Cm_Nhan != null)
                            this.Cm_Nhan();
                        btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)action);
                        break;
                    case "btnDelete":
                        if (this.Cm_Xoa != null) 
                        {
                            GridLayout gridMain = this.GetGridMain();
                            if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt"))
                            {
                                SqlCommand sqlcmd = new SqlCommand("SELECT 1 FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec ='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "' AND (isnull(tinh_trang_hddt,0) != 0 OR tinh_trang_hddt != '')");
                                int count = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count;
                                if (this.CheckDC() || count > 0)
                                {
                                    int num = (int)ExMessageBox.Show(-2017, StartupBase.SasObj, "Đã có phát sinh HĐĐT không được xóa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                            }
                            if (gridMain.DataContext != null)
                            {
                                DataView dataContext = gridMain.DataContext as DataView;
                                if (this.IsNd51 && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString().Equals("0"))
                                {
                                    string str = StartUpTrans.GetSl_in(dataContext[0]["stt_rec"].ToString().ToString()).ToString();
                                    if (!str.Equals("0"))
                                    {
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
                                        this.SetInvisible();
                                        int num = (int)ExMessageBox.Show(-990, StartupBase.SasObj, "Hóa đơn đã được in, không xóa được!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        return;
                                    }
                                }
                                else if (!this.IsNd51 && !StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString().Trim().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString().Trim()) && FormTrans.M_MA_CT_KHTMP.IndexOf(StartUpTrans.Ma_ct) == -1)
                                {
                                    int num = (int)ExMessageBox.Show(-995, FormTrans.SasO, "Chứng từ không thuộc đơn vị hiện thời!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if (this.IsNd51 && this.Cm_HuyHD != null && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString().Equals("0"))
                                {
                                    if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1)
                                    {
                                        if (!SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                        {
                                            int num = (int)ExMessageBox.Show(-1000, FormTrans.SasO, "Không thể xóa bỏ được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            return;
                                        }
                                        if (ExMessageBox.Show(-1005, FormTrans.SasO, "Có chắc chắn xóa bỏ hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                                        {
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["date"] = (object)DateTime.Now.Date;
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                            {
                                                foreach (DataRowView dataRowView1 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                                {
                                                    dataRowView1["date"] = DateTime.Now.Date;
                                                    string str = DateTime.Now.Date.ToString("HH:mm:ss");
                                                    dataRowView1["time"] = str;
                                                    dataRowView1["user_id"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                                }
                                            }
                                            this.Cm_HuyHD();
                                            if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                                            {
                                                if (dataContext.Table.Columns.Contains("ten_post"))
                                                    dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post"];
                                                if (dataContext.Table.Columns.Contains("ten_post2"))
                                                    dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post2"];
                                            }
                                            this.Toolbar.OnIsInEditModeChanged();
                                            this.SetFocusToolbar();
                                            break;
                                        }
                                        break;
                                    }
                                    break;
                                }
                                bool flag = true;
                                if (this.IsNd51 && this.C_QS != null)
                                {
                                    this.C_QS.SearchInit();
                                    DataSet dataSet = StartUpTrans.CheckQS(this.C_QS.Text, this.C_NgayHT.dValue.ToString("yyyyMMdd"), (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString()), dataContext[0]["so_ct"].ToString().Trim(), dataContext[0]["so_ct"].ToString().Trim(), "1;2;3;5");
                                    if (dataSet.Tables.Count > 0 && Convert.ToInt32(dataSet.Tables[0].Rows[0][0]) == 3)
                                    {
                                        flag = false;
                                        if (ExMessageBox.Show(-1010, FormTrans.SasO, "Đã có hóa đơn sau hóa đơn này. Có chuyển sang tình trạng xóa bỏ hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes && dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1)
                                        {
                                            if (!SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                            {
                                                int num = (int)ExMessageBox.Show(-1015, FormTrans.SasO, "Không thể xóa bỏ được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                return;
                                            }
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["date"] = (object)DateTime.Now.Date;
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                            {
                                                foreach (DataRowView dataRowView1 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                                {
                                                    dataRowView1["date"] = DateTime.Now.Date;
                                                    dataRowView1["time"] = DateTime.Now.ToString("HH:mm:ss");
                                                    dataRowView1["user_id"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                                }
                                            }
                                            this.Cm_HuyHD();
                                            if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                                            {
                                                if (dataContext.Table.Columns.Contains("ten_post"))
                                                    dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post"];
                                                if (dataContext.Table.Columns.Contains("ten_post2"))
                                                    dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post2"];
                                            }
                                            this.Toolbar.OnIsInEditModeChanged();
                                            this.SetFocusToolbar();
                                        }
                                    }
                                }
                                if ((!(StartUpTrans.Ma_ct == "HDA") && !(StartUpTrans.Ma_ct == "HD1") || !(dataContext[0]["stt_rec_pt"].ToString().Trim() != "")) && StartUpTrans.CheckPhanBo(dataContext[0]["stt_rec"].ToString()) == 1)
                                {
                                    int num = (int)ExMessageBox.Show(-1020, FormTrans.SasO, "Không xoá được, đã có chứng từ thanh toán cho cho hoá đơn này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if (dataContext.Table.Rows.Count > 1 && flag)
                                {
                                    if (dataContext.Table.Columns.Contains("ngay_ct") && !SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                    {
                                        int num = (int)ExMessageBox.Show(-1025, FormTrans.SasO, "Không thể xóa được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        return;
                                    }
                                    if (!this.CheckCanDelete())
                                        return;
                                    if (ExMessageBox.Show(-1030, FormTrans.SasO, "Có chắc chắn xóa không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                                    {
                                        SqlCommand sqlcmd = new SqlCommand("Exec UpdateInfoVoucherDelete @stt_rec, @ma_ct, @date, @time, @user_id");
                                        sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = (object)dataContext[0]["stt_rec"].ToString();
                                        sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)dataContext[0]["ma_ct"].ToString();
                                        sqlcmd.Parameters.Add("@date", SqlDbType.Char).Value = (object)DateTime.Now.Date.ToString("yyyyMMdd");
                                        sqlcmd.Parameters.Add("@time", SqlDbType.Char).Value = (object)DateTime.Now.ToString("HH:mm:ss");
                                        sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                        StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                                        this.Cm_Xoa();
                                        this.Toolbar.OnIsInEditModeChanged();
                                        this.SetFocusToolbar();
                                        break;
                                    }
                                    break;
                                }
                                break;
                            }
                            break;
                        }
                        break;
                    case "btnCancel":
                        if (this.Cm_Huy != null)
                        {
                            this.isCurrentFormActive = true;
                            if (ExMessageBox.Show(-1035, FormTrans.SasO, "Hủy bỏ các thay đổi?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                            {
                                if (FormTrans.currActionTask == ActionTask.Edit && StartUpTrans.DmctInfo != null)
                                {
                                    DataProvider.UpdateDataTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_phdbf"].ToString(), "stt_rec", this.DsBackup.Tables[0], "stt_rec;row_id");
                                    if (this.DsBackup.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()))
                                        DataProvider.UpdateCtTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_ctdbf"].ToString(), this.DsBackup.Tables[1], StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                                    if (this.DsBackup.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                        DataProvider.UpdateCtTable(this.BindingSasObj, StartUpTrans.DmctInfo["m_ctgtdbf"].ToString(), this.DsBackup.Tables[2], StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                                }
                                this.Cm_Huy();
                            }
                            btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)action);
                            this.isCurrentFormActive = false;
                            break;
                        }
                        break;
                    case "btnView":
                        if (this.Cm_Xem != null)
                        {
                            this.Cm_Xem();
                            break;
                        }
                        break;
                    case "btnSearch":
                        if (this.Cm_Tim != null)
                        {
                            this.Cm_Tim();
                            break;
                        }
                        break;
                    case "btnPrint":
                        if (this.Cm_In != null)
                        {
                            GridLayout gridMain = this.GetGridMain();
                            if (gridMain.DataContext != null)
                            {
                                DataView dataContext = gridMain.DataContext as DataView;
                                string str1 = dataContext[0]["stt_rec"].ToString().ToString();
                                string str2 = dataContext[0]["ma_ct"].ToString().ToString();
                                string cmdText = string.Empty;
                                if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                                {
                                    cmdText = "exec LoadVoucher#1 @ma_ct, @PhFilter, @CtFilter, @Sl_ct";
                                    if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 1)
                                        cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[1]);
                                }
                                else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                                {
                                    cmdText = "exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct";
                                    if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 0)
                                        cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[0]);
                                }
                                SqlCommand cmd = new SqlCommand(cmdText);
                                if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                                {
                                    cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                                    cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                                    cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                    cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                                }
                                else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                                {
                                    cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                                    cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                                    cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                    cmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                    cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                                }
                                if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                                {
                                    DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, cmd);
                                    if (dataSet.Tables[0].Rows.Count == 1)
                                    {
                                        StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + str1 + "'")[0].ItemArray = dataSet.Tables[0].Rows[0].ItemArray;
                                        if (StartUpTrans.DsTrans.Tables.Count > 1)
                                        {
                                            string rowFilter = StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter;
                                            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                            foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str1 + "'"))
                                                StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                                            foreach (DataRow dataRow in dataSet.Tables[1].Select("stt_rec='" + str1 + "'"))
                                                StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow.ItemArray);
                                            StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = rowFilter;
                                        }
                                        if (StartUpTrans.DsTrans.Tables.Count > 2)
                                        {
                                            try
                                            {
                                                string rowFilter = StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter;
                                                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                                foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + str1 + "'"))
                                                    StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                                                foreach (DataRow dataRow in dataSet.Tables[2].Select("stt_rec='" + str1 + "'"))
                                                    StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow.ItemArray);
                                                StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = rowFilter;
                                            }
                                            catch (Exception ex)
                                            {
                                            }
                                        }
                                    }
                                }
                            }
                            this.Cm_In();
                        }
                        ToolBarButton btnPrint = this.Toolbar.FindName("btnPrint") as ToolBarButton;
                        btnPrint.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => btnPrint.Focus()));
                        break;
                    case "btnCopy":
                        if (this.Cm_Copy != null && !this.IsEditMode)
                        {
                            this.Cm_Copy();
                            if (StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("ngay_lct") && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("ngay_ct"))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                            if (this.IsEditMode)
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"];
                                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("loai_tg") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = (object)StartUpTrans.Getloai_tg(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                            }
                            this.C_QS_Text_Changed();
                        }
                        this.SetInvisible();
                        break;
                    case "btnVoucherConextMenu":
                        if (this.StartUpMain != null && StartUpTrans.DmctInfo != null && (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables.Count > 0) && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                        {
                            DataTable table = FormTrans.SasO.ExcuteReader(new SqlCommand("SELECT * from " + StartUpTrans.DmctInfo["m_phdbf"] + " WHERE stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'")).Tables[0];
                            if (table.Rows.Count == 1)
                            {
                                this.Toolbar.vc.VoucherRow = table.Rows[0];
                                this.Toolbar.vc.DmctRow = StartUpTrans.DmctInfo;
                                break;
                            }
                            break;
                        }
                        break;
                    case "btnReportFromVoucher":
                        this.InitReportMenu();
                        break;
                    case "btnTemplate":
                        GridLayout gridMain1 = this.GetGridMain();
                        if (gridMain1.DataContext != null)
                        {
                            DataView dataContext = gridMain1.DataContext as DataView;
                            if (dataContext.Table.Columns.Contains("stt_rec"))
                            {
                                this.Toolbar.currentStt_rec = dataContext[0]["stt_rec"].ToString();
                                break;
                            }
                            break;
                        }
                        break;
                    case "btnOptions":
                        DataTable dataTable = (DataTable)null;
                        SqlCommand sqlcmd1 = new SqlCommand();
                        try
                        {
                            string str = "select * from v_dmct where ma_ct = @ma_ct";
                            sqlcmd1.Parameters.Add("@ma_ct", SqlDbType.VarChar).Value = (object)StartUpTrans.Ma_ct;
                            sqlcmd1.CommandText = str;
                            dataTable = this.BindingSasObj.ExcuteReader(sqlcmd1).Tables[0];
                        }
                        catch (SqlException ex)
                        {
                            ErrorLog.CatchMessage(ex);
                        }
                        if ((bool)this.CallModule("Sasct.exe;Sasct.StartUp;Extend_oBrowse_Command", new object[3]{
                                  (object) this.BindingSasObj,
                                  (object) ActionTask.Edit,
                                  (object) dataTable
                        }))
                        {
                            string[] parameters = new string[2]
                            {
                StartUpTrans.CommandInfo["menu_id"].ToString(),
                StartUpTrans.Editing_Stt_Rec
                            };
                            SysFunc.CallModule(StartUpTrans.CommandInfo["procedure"].ToString(), parameters, Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), StartupBase.SasObj.M_ProcessName, StartupBase.SasObj);
                            this.Close();
                            break;
                        }
                        break;
                    case "btnHelp":
                        SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.None, Key.F1);
                        break;
                }
                if (this.EditModeEnded != null && btn != null)
                    this.EditModeEnded((object)this, btn.Name, e);
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (btn == null)
                       return;
                   this.CheckND51();
                   this.C_QS_Text_Changed();
                   this.SetInvisible();
                   this.SetRound();
                   GC.Collect();
                   GC.WaitForPendingFinalizers();
                   GC.Collect();
               }));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private bool CheckExistsCode(
          string column,
          BasicGridView grid,
          string columncheck,
          out int index)
        {
            bool flag = false;
            index = -1;
            try
            {
                for (int index1 = 0; index1 < grid.FieldLayouts[0].Fields.Count; ++index1)
                {
                    if (grid.FieldLayouts[0].Fields[index1].Name.Equals(column))
                        flag = true;
                }
                if (flag)
                {
                    string[] strArray = this.BindingSasObj.DmdmInfo.Select("ma_dm LIKE 'dmtk'")[0]["doi_ma"].ToString().Split(';');
                    DataTable dataTable = StartUpTrans.Getdmtk(columncheck + " = 1");
                    for (int index1 = 0; index1 < grid.Records.Count; ++index1)
                    {
                        DataRecord record = grid.Records[index1] as DataRecord;
                        for (int index2 = 0; index2 < record.Cells.Count; ++index2)
                        {
                            Cell cell = record.Cells[index2];
                            if (cell != null && ((IEnumerable<string>)strArray).Contains<string>(cell.Field.Name) && dataTable.Select("tk LIKE '" + cell.Value.ToString().Trim() + "'").Length > 0 && (record.Cells[column].Value == DBNull.Value || string.IsNullOrEmpty(record.Cells[column].Value.ToString().Trim())))
                            {
                                index = index1;
                                return false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
            return true;
        }

        private object CallModule(string StrExecute, object[] parameters)
        {
            try
            {
                string startUpPath = this.BindingSasObj.M_StartUp_Path;
                string[] strArray = StrExecute.Split(';');
                string str = startUpPath + (startUpPath.Substring(startUpPath.Length - 1, 1) == "\\" ? "" : "\\");
                FormTrans.SasO.SynchroFile(".", strArray[0].Trim());
                Type type = Assembly.LoadFile(str + strArray[0]).GetType(strArray[1]);
                return type.GetMethod(strArray[2]).Invoke(Activator.CreateInstance(type), parameters);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return (object)null;
        }

        private void InitReportMenu()
        {
            if (this.cm == null)
            {
                List<string> stringList = new List<string>((IEnumerable<string>)FormTrans.SasO.UserInfo.Rows[0]["rights"].ToString().Split('/'));
                this.ListCommand = (IEnumerable<SasDefine.Command>)FormTrans.SasO.CommandInfo.AsEnumerable().OrderBy<DataRow, string>((Func<DataRow, string>)(currMenuList => currMenuList.Field<string>("menu_id0"))).ThenBy<DataRow, string>((Func<DataRow, string>)(currMenuList => currMenuList.Field<string>("menu_id"))).Select<DataRow, SasDefine.Command>((Func<DataRow, SasDefine.Command>)(currMenuList => new SasDefine.Command()
                {
                    menu_id = currMenuList.Field<string>("menu_id"),
                    menu_id0 = currMenuList.Field<string>("menu_id0"),
                    bar = currMenuList.Field<string>("bar"),
                    bar2 = currMenuList.Field<string>("bar2"),
                    nh_menu = currMenuList.Field<string>("nh_menu"),
                    procedure = currMenuList.Field<string>("procedure"),
                    picture = currMenuList.Field<string>("picture"),
                    Vouchers = currMenuList.Field<string>("Vouchers")
                }));
                this.cm = new ContextMenu();
                string menu_id = StartupBase.Menu_Id;
                foreach (SasDefine.Command command in this.ListCommand.Where<SasDefine.Command>((Func<SasDefine.Command, bool>)(p => p.Vouchers == "1" && p.menu_id.Substring(0, 2) == menu_id.Substring(0, 2))))
                {
                    MenuItem parentItem = new MenuItem();
                    if (StartupBase.M_LAN.Equals("V"))
                        parentItem.Header = (object)command.bar;
                    else if (StartupBase.M_LAN.Equals("E"))
                        parentItem.Header = (object)command.bar2;
                    parentItem.Tag = (object)command;
                    if (string.IsNullOrEmpty(command.procedure))
                    {
                        this.AddChildItem(ref parentItem, command.menu_id);
                        this.AddChildItemEvent(parentItem);
                    }
                    else
                        parentItem.Click += new RoutedEventHandler(this.miCommand_Click);
                    this.cm.Items.Add((object)parentItem);
                }
            }
            this.cm.IsOpen = true;
        }

        private void AddChildItem(ref MenuItem parentItem, string menu_id)
        {
            foreach (SasDefine.Command command in this.ListCommand.Where<SasDefine.Command>((Func<SasDefine.Command, bool>)(p => p.menu_id0.Trim() == menu_id)))
            {
                MenuItem parentItem1 = new MenuItem();
                if (StartupBase.M_LAN.Equals("V"))
                    parentItem1.Header = (object)command.bar;
                else if (StartupBase.M_LAN.Equals("E"))
                    parentItem1.Header = (object)command.bar2;
                parentItem1.Tag = (object)command;
                if (string.IsNullOrEmpty(command.procedure))
                    this.AddChildItem(ref parentItem1, command.menu_id);
                parentItem.Items.Add((object)parentItem1);
            }
        }

        private void AddChildItemEvent(MenuItem it)
        {
            it.Click += new RoutedEventHandler(this.miCommand_Click);
            foreach (MenuItem it1 in (IEnumerable)it.Items)
                this.AddChildItemEvent(it1);
        }

        private void RemoveChildItemEvent(MenuItem it)
        {
            it.Click -= new RoutedEventHandler(this.miCommand_Click);
            foreach (MenuItem it1 in (IEnumerable)it.Items)
                this.RemoveChildItemEvent(it1);
        }

        private void miCommand_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (string.IsNullOrEmpty((menuItem.Tag as SasDefine.Command).procedure))
                return;
            string[] parameters = new string[1]
            {
        (menuItem.Tag as SasDefine.Command).menu_id
            };
            SysFunc.CallModule((menuItem.Tag as SasDefine.Command).procedure, parameters, FormTrans.SasO.M_StartUp_Path, FormTrans.SasO.M_ProcessName, FormTrans.SasO);
        }

        private void FormTrans_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (Keyboard.IsKeyDown(Key.RightAlt) || Keyboard.IsKeyDown(Key.LeftAlt))
                    return;
                string btnName = string.Empty;
                if (Keyboard.IsKeyDown(Key.Home))
                {
                    if (this.Cm_Dau != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnTop") as Button).Focus();
                        this.Cm_Dau();
                        this.CheckBoxVoucher();
                        btnName = "btnTop";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.End))
                {
                    if (this.Cm_Cuoi != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnBottom") as Button).Focus();
                        this.Cm_Cuoi();
                        this.CheckBoxVoucher();
                        btnName = "btnBottom";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.Prior))
                {
                    if (this.Cm_Truoc != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnPrevious") as Button).Focus();
                        this.Cm_Truoc();
                        this.CheckBoxVoucher();
                        btnName = "btnPrevious";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.Next))
                {
                    if (this.Cm_Sau != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnNext") as Button).Focus();
                        this.Cm_Sau();
                        this.CheckBoxVoucher();
                        btnName = "btnNext";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.N) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) || Keyboard.IsKeyDown(Key.F4) && Keyboard.Modifiers != ModifierKeys.Control)
                {
                    if (this.Cm_Moi != null && !this.IsEditMode && (((IEnumerable<string>)this.Toolbar.ListAdd).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1))
                    {
                        (this.Toolbar.FindName("btnNew") as Button).Focus();
                        this.Cm_Moi();
                        if (this.IsEditMode)
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"];
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("loai_tg") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = (object)StartUpTrans.Getloai_tg(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("hd_thue"))
                            {
                                if (FormTrans.currActionTask != ActionTask.Edit)
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["hd_thue"] = !StartupBase.SasObj.GetSysvar("M_CHECK_VOUCHER").ToString().Trim().Equals("2") || !StartUpTrans.hd_thue.Trim().Equals("1") && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim().ToUpper().Equals("QL1") ? (object)"0" : (object)"1";
                                this.CheckBoxVoucher();
                            }
                        }
                        btnName = "btnNew";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F4) && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (this.Cm_Copy != null && !this.IsEditMode && (((IEnumerable<string>)this.Toolbar.ListAdd).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1))
                    {
                        (this.Toolbar.FindName("btnCopy") as Button).Focus();
                        this.Cm_Copy();
                        if (StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("ngay_lct") && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("ngay_ct"))
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                        if (this.IsEditMode)
                        {
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"] = FormTrans.SasO.DmdvcsInfo.Rows[0]["ma_dvcs"];
                            if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("loai_tg") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["loai_tg"] = (object)StartUpTrans.Getloai_tg(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString());
                        }
                        btnName = "btnCopy";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.S) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                {
                    Button name = this.Toolbar.FindName("btnSave") as Button;
                    name.Focus();
                    if (this.Cm_Nhan != null && this.IsEditMode && (name.IsFocused && StartUpTrans.DsTrans != null))
                    {
                        GridLayout gridMain = this.GetGridMain();
                        if (gridMain.DataContext != null)
                        {
                            DataView dataContext = gridMain.DataContext as DataView;
                            if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1)
                            {
                                if (dataContext[0]["ngay_ct"] != DBNull.Value && !SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                {
                                    int num = (int)ExMessageBox.Show(-1040, FormTrans.SasO, "Ngày hạch toán phải sau ngày khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    object obj = this.FindName("txtNgay_ct") ?? this.FindName("txtngay_ct");
                                    if (obj == null)
                                        return;
                                    FrameworkElement dNgay_ct = obj as FrameworkElement;
                                    dNgay_ct.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => dNgay_ct.Focus()));
                                    return;
                                }
                                if (!SysFunc.CheckValidNgayMs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"]))))
                                {
                                    int num = (int)ExMessageBox.Show(-1045, FormTrans.SasO, "Ngày hạch toán phải sau ngày mở sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    object obj = this.FindName("txtNgay_ct") ?? this.FindName("txtngay_ct");
                                    if (obj == null)
                                        return;
                                    FrameworkElement dNgay_ct = obj as FrameworkElement;
                                    dNgay_ct.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => dNgay_ct.Focus()));
                                    return;
                                }
                            }
                            DateTime dateTime;
                            if (this.C_QS != null && this.C_NgayHT != null && (!string.IsNullOrEmpty(this.C_QS.Text.Trim()) && !string.IsNullOrEmpty(this.C_NgayHT.Text.Trim())))
                            {
                                string text1 = this.C_QS.Text;
                                dateTime = this.C_NgayHT.dValue;
                                string ngay_ct = dateTime.ToString("yyyyMMdd");
                                int int16 = (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString());
                                string text2 = this.C_So_ct.Text;
                                switch (StartUpTrans.CheckQS(text1, ngay_ct, int16, text2))
                                {
                                    case 1:
                                        int num1 = (int)ExMessageBox.Show(-1050, FormTrans.SasO, "Ngày bắt đầu sử dụng ký hiệu không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_NgayHT.Focus()));
                                        return;
                                    case 2:
                                        int num2 = (int)ExMessageBox.Show(-1055, FormTrans.SasO, "Quyền sử dụng ký hiệu không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_QS.IsFocus = true));
                                        return;
                                    case 3:
                                        if (this.IsNd51)
                                        {
                                            int num3 = (int)ExMessageBox.Show(-1060, FormTrans.SasO, "Ngày c.từ hiện tại nhỏ hơn ngày c.từ cuối cùng của quyển c.từ, không lưu được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.C_NgayHT.Focus()));
                                            return;
                                        }
                                        break;
                                }
                            }
                            if (dataContext.Table.Columns.Contains("so_ct"))
                            {
                                string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim());
                                if (dataContext.Table.Columns.Contains("ma_qs") && !this.IsNd51 && (!string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim()) && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim())))
                                {
                                    switch (StartUpTrans.CheckValidSo_ct(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_qs"].ToString().Trim(), StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim()))
                                    {
                                        case 2:
                                            int num1 = (int)ExMessageBox.Show(-1366, FormTrans.SasO, "Số c.từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            break;
                                        case 3:
                                            int num2 = (int)ExMessageBox.Show(-1366, FormTrans.SasO, "Số c.từ không hợp lệ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            return;
                                    }
                                }
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct"), ' ');
                            }
                            if (this.GridCt != null && this.GridCt.Count > 0)
                            {
                                for (int i = 0; i < this.GridCt.Count; ++i)
                                {
                                    if (this.GridCt[i].DataSource != null && this.GridCt[i].Visibility == Visibility.Visible)
                                    {
                                        IEnumerable dataSource = this.GridCt[i].DataSource;
                                        int index = -1;
                                        if (!this.CheckExistsCode("ma_vv_i", this.GridCt[i], "ma_vv_yn", out index))
                                        {
                                            int num = (int)ExMessageBox.Show(-1367, FormTrans.SasO, "Mã dự án không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                            this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_vv_i"];
                                            this.GridCt[i].Focus();
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                            return;
                                        }
                                        if (!this.CheckExistsCode("ma_px_i", this.GridCt[i], "ma_px_yn", out index))
                                        {
                                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                            int num = (int)ExMessageBox.Show(-1368, FormTrans.SasO, "Mã phân xưởng không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_px_i"];
                                            this.GridCt[i].Focus();
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                            return;
                                        }
                                        if (!this.CheckExistsCode("ma_bpht_i", this.GridCt[i], "ma_bpht_yn", out index))
                                        {
                                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                            int num = (int)ExMessageBox.Show(-1369, FormTrans.SasO, "Mã bộ phận hạch toán không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_bpht_i"];
                                            this.GridCt[i].Focus();
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                            return;
                                        }
                                        if (!this.CheckExistsCode("ma_phi_i", this.GridCt[i], "ma_phi_yn", out index))
                                        {
                                            SasFormBrowes.WinAPISenkey.SenKey(ModifierKeys.Alt, Key.D1);
                                            int num = (int)ExMessageBox.Show(-1371, FormTrans.SasO, "Mã phí không được phép rỗng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                            this.GridCt[i].ActiveCell = (this.GridCt[i].Records[index] as DataRecord).Cells["ma_phi_i"];
                                            this.GridCt[i].Focus();
                                            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => this.GridCt[i].ExecuteCommand(DataPresenterCommands.StartEditMode)));
                                            return;
                                        }
                                    }
                                }
                            }
                            if (StartUpTrans.DsTrans.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[1].Columns.Contains("so_ct0")))
                            {
                                bool flag = StartUpTrans.DsTrans.Tables[1].Columns.Contains("so_seri0");
                                foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[1].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                {
                                    dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct0"), ' ');
                                    if (flag)
                                        dataRow["so_seri0"] = (object)dataRow["so_seri0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_seri0"), ' ');
                                }
                            }
                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_ct0")))
                            {
                                bool flag = StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_seri0");
                                foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[2].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                {
                                    dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_ct0"), ' ');
                                    if (flag)
                                        dataRow["so_seri0"] = (object)dataRow["so_seri0"].ToString().Trim().PadLeft(FormTrans.SasO.GetDatabaseFieldLength("so_seri0"), ' ');
                                }
                            }
                            if (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy)
                            {
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                DataRowView dataRowView1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
                                dataRowView1["date0"] = DateTime.Now.Date;
                                dataRowView1["time0"] = DateTime.Now.ToString("HH:mm:ss");
                                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("user_name0"))
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("date0")) && (StartUpTrans.DsTrans.Tables[2].Columns.Contains("time0") && StartUpTrans.DsTrans.Tables[2].Columns.Contains("user_id0")))
                                {
                                    foreach (DataRowView dataRowView3 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                    {
                                        dataRowView3["date0"] = DateTime.Now.Date;
                                        dataRowView3["time0"] = DateTime.Now.ToString("HH:mm:ss");
                                        dataRowView3["user_id0"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                    }
                                }
                            }
                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                            if (dataContext.Table.Columns.Contains("user_name"))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                            DataRowView dataRowView6 = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
                            dataRowView6["date"] = DateTime.Now.Date;
                            dataRowView6["time"] = DateTime.Now.ToString("HH:mm:ss"); ;
                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("date")) && (StartUpTrans.DsTrans.Tables[2].Columns.Contains("time") && StartUpTrans.DsTrans.Tables[2].Columns.Contains("user_id")))
                            {
                                foreach (DataRowView dataRowView1 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                {
                                    dataRowView1["date"] = DateTime.Now.Date;
                                    dataRowView1["time"] = DateTime.Now.ToString("HH:mm:ss");
                                    dataRowView1["user_id"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                }
                            }
                          (this.Toolbar.FindName("btnSave") as ToolBarButton).Focus();
                            string str4 = "";
                            if (dataContext.Table.Columns.Contains("ma_qs"))
                                str4 = dataContext[0]["ma_qs"].ToString().Trim();
                            this.Cm_Nhan();
                            if (!this.IsEditMode && dataContext.Table.Columns.Contains("ma_qs") && StartUpTrans.Ma_ct != "HD1")
                            {
                                string format = "EXEC  {0} '" + str4 + "', '" + dataContext[0]["so_ct"].ToString().Trim() + "'";
                                string cmdText = StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 2 ? string.Format(format, (object)"SetSoct") : string.Format(format, (object)StartUpTrans.Process_Store[2]);
                                FormTrans.SasO.ExcuteNonQuery(new SqlCommand(cmdText));
                            }
                            if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                            {
                                if (dataContext.Table.Columns.Contains("ten_post"))
                                    dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act"];
                                if (dataContext.Table.Columns.Contains("ten_post2"))
                                    dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_act2"];
                            }
                            btnName = "btnSave";
                            e.Handled = true;
                            ToolBarButton btnMoi = this.Toolbar.FindName("btnNew") as ToolBarButton;
                            Action action = (Action)(() => btnMoi.Focus());
                            btnMoi.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)action);
                        }
                    }
                    else if (this.Cm_Nhan != null && this.IsEditMode && name.IsFocused)
                        this.Cm_Nhan();
                }
                else if (Keyboard.IsKeyDown(Key.P) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                {
                    if (this.Cm_In != null && this.Toolbar.btnPrint.IsEnabled)
                    {
                        GridLayout gridMain = this.GetGridMain();
                        if (gridMain.DataContext != null)
                        {
                            DataView dataContext = gridMain.DataContext as DataView;
                            string str1 = dataContext[0]["stt_rec"].ToString().ToString();
                            string str2 = dataContext[0]["ma_ct"].ToString().ToString();
                            string cmdText = string.Empty;
                            if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                            {
                                cmdText = "exec LoadVoucher#1 @ma_ct, @PhFilter, @CtFilter, @Sl_ct";
                                if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 1)
                                    cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[1]);
                            }
                            else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                            {
                                cmdText = "exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct";
                                if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 0)
                                    cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[0]);
                            }
                            SqlCommand cmd = new SqlCommand(cmdText);
                            if (str2.ToString().ToUpper().Trim().Equals("QL2") || str2.ToString().ToUpper().Trim().Equals("QL3") || str2.ToString().ToUpper().Trim().Equals("QL4"))
                            {
                                cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                                cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                                cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                            }
                            else if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                            {
                                cmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.Ma_ct;
                                cmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + str1 + "'");
                                cmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                cmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
                                cmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
                            }
                            if (!str2.ToString().ToUpper().Trim().Equals("QL1"))
                            {
                                DataSet dataSet = DataProvider.FillCommand(StartupBase.SasObj, cmd);
                                if (dataSet.Tables[0].Rows.Count == 1)
                                {
                                    StartUpTrans.DsTrans.Tables[0].Select("stt_rec = '" + str1 + "'")[0].ItemArray = dataSet.Tables[0].Rows[0].ItemArray;
                                    if (StartUpTrans.DsTrans.Tables.Count > 1)
                                    {
                                        string rowFilter = StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter;
                                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                        foreach (DataRow row in StartUpTrans.DsTrans.Tables[1].Select("stt_rec='" + str1 + "'"))
                                            StartUpTrans.DsTrans.Tables[1].Rows.Remove(row);
                                        foreach (DataRow dataRow in dataSet.Tables[1].Select("stt_rec='" + str1 + "'"))
                                            StartUpTrans.DsTrans.Tables[1].Rows.Add(dataRow.ItemArray);
                                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = rowFilter;
                                    }
                                    if (StartUpTrans.DsTrans.Tables.Count > 2)
                                    {
                                        try
                                        {
                                            string rowFilter = StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter;
                                            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = "stt_rec= '" + StartUpTrans.DsTrans.Tables[0].Rows[0]["stt_rec"].ToString() + "'";
                                            foreach (DataRow row in StartUpTrans.DsTrans.Tables[2].Select("stt_rec='" + str1 + "'"))
                                                StartUpTrans.DsTrans.Tables[2].Rows.Remove(row);
                                            foreach (DataRow dataRow in dataSet.Tables[2].Select("stt_rec='" + str1 + "'"))
                                                StartUpTrans.DsTrans.Tables[2].Rows.Add(dataRow.ItemArray);
                                            StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = rowFilter;
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                    }
                                }
                            }
                        }
                        this.Cm_In();
                        btnName = "btnPrint";
                        e.Handled = true;
                        ToolBarButton btnPrint = this.Toolbar.FindName("btnPrint") as ToolBarButton;
                        btnPrint.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() => btnPrint.Focus()));
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F3))
                {
                    if (this.Cm_Sua != null && this.Toolbar.btnEdit.IsEnabled && (((IEnumerable<string>)this.Toolbar.ListEdit).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1))
                    {
                        GridLayout gridMain = this.GetGridMain();
                        if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt"))
                        {
                            SqlCommand sqlcmd = new SqlCommand("SELECT 1 FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec ='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "' AND (isnull(tinh_trang_hddt,0) != 0 OR tinh_trang_hddt != '')");
                            int count = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count;
                            if (this.CheckDC() || count > 0)
                            {
                                int num = (int)ExMessageBox.Show(-2016, StartupBase.SasObj, "Đã có phát sinh HĐĐT không được sửa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                        }
                        if (gridMain.DataContext != null)
                        {
                            DataView dataContext = gridMain.DataContext as DataView;
                            if (this.IsNd51)
                            {
                                string str = StartUpTrans.GetSl_in(dataContext[0]["stt_rec"].ToString().ToString()).ToString();
                                if (!str.Equals("0"))
                                {
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
                                    if (ExMessageBox.Show(-1070, StartupBase.SasObj, "Hóa đơn đã được in, có muốn sửa lại chứng từ hay không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                                        return;
                                }
                            }
                            else if (!StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(-1075, FormTrans.SasO, "Chứng từ không thuộc đơn vị hiện thời!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                            if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1 && !SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                            {
                                int num = (int)ExMessageBox.Show(-1080, FormTrans.SasO, "Dữ liệu đã khóa sổ, không sửa được!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                            if ((!(StartUpTrans.Ma_ct == "HDA") && !(StartUpTrans.Ma_ct == "HD1") || !(dataContext[0]["stt_rec_pt"].ToString().Trim() != "")) && StartUpTrans.CheckPhanBo(dataContext[0]["stt_rec"].ToString()) == 1)
                            {
                                int num = (int)ExMessageBox.Show(-1085, FormTrans.SasO, "Không được sửa hóa đơn đã được thanh toán!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                            if (this.C_QS != null && this.C_NgayHT != null && (!string.IsNullOrEmpty(this.C_QS.Text.Trim()) && !string.IsNullOrEmpty(this.C_NgayHT.Text.Trim())) && StartUpTrans.CheckQS(this.C_QS.Text, this.C_NgayHT.dValue.ToString("yyyyMMdd"), (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString()), this.C_So_ct.Text) == 2)
                            {
                                int num = (int)ExMessageBox.Show(-981, FormTrans.SasO, "Quyền sử dụng ký hiệu không hợp lệ, không được sửa chứng từ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                          (this.Toolbar.FindName("btnEdit") as Button).Focus();
                            this.DsBackup = new DataSet();
                            this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[0].DefaultView.ToTable());
                            if (StartUpTrans.DsTrans.Tables.Count >= 2 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim()))
                                this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[1].DefaultView.ToTable());
                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                this.DsBackup.Tables.Add(StartUpTrans.DsTrans.Tables[2].DefaultView.ToTable());
                            this.Cm_Sua();
                            btnName = "btnEdit";
                            e.Handled = true;
                            if (dataContext.Table.Columns.Contains("ngay_ct"))
                                StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"] = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"].ToString().Trim();
                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && (!string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()) && StartUpTrans.DsTrans.Tables[2].Columns.Contains("so_ct0")))
                            {
                                foreach (DataRow dataRow in StartUpTrans.DsTrans.Tables[2].Select(string.Format("stt_rec = '{0}'", (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                                    dataRow["so_ct0"] = (object)dataRow["so_ct0"].ToString().Trim();
                            }
                        }
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F8))
                {
                    if (this.Cm_Xoa != null && this.Toolbar.btnDelete.IsEnabled && (((IEnumerable<string>)this.Toolbar.ListDelete).Contains<string>(this.Toolbar.Menu_id) || this.Toolbar.IsAdmin == 1))
                    {
                        (this.Toolbar.FindName("btnDelete") as Button).Focus();
                        GridLayout gridMain = this.GetGridMain();
                        if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("tinh_trang_hddt"))
                        {
                            SqlCommand sqlcmd = new SqlCommand("SELECT 1 FROM " + StartUpTrans.DmctInfo["m_phdbf"].ToString() + " WHERE stt_rec ='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "' AND (isnull(tinh_trang_hddt,0) != 0 OR tinh_trang_hddt != '')");
                            int count = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Rows.Count;
                            if (this.CheckDC() || count > 0)
                            {
                                int num = (int)ExMessageBox.Show(-2017, StartupBase.SasObj, "Đã có phát sinh HĐĐT không được xóa!", "Xac nhan nhap lieu", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                        }
                        if (gridMain.DataContext != null)
                        {
                            DataView dataContext = gridMain.DataContext as DataView;
                            if (this.IsNd51 && StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString().Equals("0"))
                            {
                                string str = StartUpTrans.GetSl_in(dataContext[0]["stt_rec"].ToString().ToString()).ToString();
                                if (!str.Equals("0"))
                                {
                                    StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"] = (object)str;
                                    this.SetInvisible();
                                    int num = (int)ExMessageBox.Show(-1090, StartupBase.SasObj, "Hóa đơn đã được in, không xóa được!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                            }
                            else if (!this.IsNd51 && !StartupBase.SasObj.DmdvcsInfo.Rows[0]["ma_dvcs"].ToString().Equals(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString()))
                            {
                                int num = (int)ExMessageBox.Show(-1095, FormTrans.SasO, "Chứng từ không thuộc đơn vị hiện thời!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                return;
                            }
                            if (this.IsNd51 && !StartUpTrans.DsTrans.Tables[0].DefaultView[0]["sl_in"].ToString().Equals("0"))
                            {
                                if (dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1)
                                {
                                    if (!SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                    {
                                        int num = (int)ExMessageBox.Show(-1100, FormTrans.SasO, "Không thể hủy được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        return;
                                    }
                                    if (ExMessageBox.Show(-1105, FormTrans.SasO, "Có chắc chắn hủy hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                                    {
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["date"] = (object)DateTime.Now.Date;
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0]["time"] = (object)DateTime.Now.ToString("HH:mm:ss");
                                        if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                        {
                                            foreach (DataRowView dataRowView1 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                            {
                                                dataRowView1["date"] = DateTime.Now.Date;
                                                dataRowView1["time"] = DateTime.Now.ToString("HH:mm:ss");
                                                dataRowView1["user_id"] = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                            }
                                        }
                                        this.Cm_HuyHD();
                                        if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                                        {
                                            if (dataContext.Table.Columns.Contains("ten_post"))
                                                dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post"];
                                            if (dataContext.Table.Columns.Contains("ten_post2"))
                                                dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post2"];
                                        }
                                        this.Toolbar.OnIsInEditModeChanged();
                                        this.SetFocusToolbar();
                                    }
                                }
                            }
                            else
                            {
                                bool flag = true;
                                DateTime dateTime;
                                if (this.IsNd51 && this.C_QS != null)
                                {
                                    this.C_QS.SearchInit();
                                    string text = this.C_QS.Text;
                                    dateTime = this.C_NgayHT.dValue;
                                    string ngay_ct = dateTime.ToString("yyyyMMdd");
                                    int int16 = (int)Convert.ToInt16(FormTrans.SasO.UserInfo.Rows[0]["user_id"].ToString());
                                    string lstso_ct1 = dataContext[0]["so_ct"].ToString().Trim();
                                    string lstso_ct2 = dataContext[0]["so_ct"].ToString().Trim();
                                    DataSet dataSet = StartUpTrans.CheckQS(text, ngay_ct, int16, lstso_ct1, lstso_ct2, "1;2;3;5");
                                    if (dataSet.Tables.Count > 0 && Convert.ToInt32(dataSet.Tables[0].Rows[0][0]) == 3)
                                    {
                                        flag = false;
                                        if (ExMessageBox.Show(-1110, FormTrans.SasO, "Đã có hóa đơn sau hóa đơn này. Có chuyển sang tình trạng xóa bỏ hóa đơn không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes && dataContext.Table.Columns.Contains("ngay_ct") && dataContext.Table.Rows.Count > 1)
                                        {
                                            if (!SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                            {
                                                int num = (int)ExMessageBox.Show(-1115, FormTrans.SasO, "Không thể hủy được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                return;
                                            }
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                            StartUpTrans.DsTrans.Tables[0].DefaultView[0]["user_name"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_name"].ToString();
                                            DataRowView dataRowView1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0];
                                            dataRowView1["date"] = DateTime.Now.Date;
                                            dataRowView1["time"] = DateTime.Now.ToString("HH:mm:ss");
                                            if (StartUpTrans.DsTrans.Tables.Count >= 3 && StartUpTrans.DmctInfo != null && !string.IsNullOrEmpty(StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim()))
                                            {
                                                foreach (DataRowView dataRowView3 in StartUpTrans.DsTrans.Tables[2].DefaultView)
                                                {
                                                    dataRowView3["date"] = DateTime.Now.Date;
                                                    string str2 = dateTime.ToString("HH:mm:ss");
                                                    dataRowView3["time"] = DateTime.Now.ToString("HH:mm:ss");
                                                    dataRowView3["user_id"] = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                                }
                                            }
                                            this.Cm_HuyHD();
                                            if (!this.IsEditMode && dataContext.Table.Columns.Contains("status"))
                                            {
                                                if (dataContext.Table.Columns.Contains("ten_post"))
                                                    dataContext[0]["ten_post"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post"];
                                                if (dataContext.Table.Columns.Contains("ten_post2"))
                                                    dataContext[0]["ten_post2"] = ((IEnumerable<DataRow>)StartUpTrans.tbStatus.Select("ma_post = " + dataContext[0]["status"].ToString())).First<DataRow>()["ten_post2"];
                                            }
                                            this.Toolbar.OnIsInEditModeChanged();
                                            this.SetFocusToolbar();
                                        }
                                    }
                                }
                                if ((!(StartUpTrans.Ma_ct == "HDA") && !(StartUpTrans.Ma_ct == "HD1") || !(dataContext[0]["stt_rec_pt"].ToString().Trim() != "")) && StartUpTrans.CheckPhanBo(dataContext[0]["stt_rec"].ToString()) == 1)
                                {
                                    int num = (int)ExMessageBox.Show(-1120, FormTrans.SasO, "Không xoá được, đã có chứng từ thanh toán cho cho hoá đơn này!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    return;
                                }
                                if (dataContext.Table.Rows.Count > 1 && flag)
                                {
                                    if (dataContext.Table.Columns.Contains("ngay_ct") && !SysFunc.CheckValidNgayKs(FormTrans.SasO, new DateTime?(Convert.ToDateTime(dataContext[0]["ngay_ct"])), StartUpTrans.Ma_ct))
                                    {
                                        int num = (int)ExMessageBox.Show(-1125, FormTrans.SasO, "Không thể xóa được chứng từ đã khóa sổ!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                        return;
                                    }
                                    if (!this.CheckCanDelete())
                                        return;
                                    if (ExMessageBox.Show(-1130, FormTrans.SasO, "Có chắc chắn xóa không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                                    {
                                        SqlCommand sqlcmd = new SqlCommand("Exec UpdateInfoVoucherDelete @stt_rec, @ma_ct, @date, @time, @user_id");
                                        sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char).Value = (object)dataContext[0]["stt_rec"].ToString();
                                        sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)dataContext[0]["ma_ct"].ToString();
                                        SqlParameter sqlParameter1 = sqlcmd.Parameters.Add("@date", SqlDbType.Char);
                                        dateTime = DateTime.Now;
                                        dateTime = dateTime.Date;
                                        string str1 = dateTime.ToString("yyyyMMdd");
                                        sqlParameter1.Value = (object)str1;
                                        SqlParameter sqlParameter2 = sqlcmd.Parameters.Add("@time", SqlDbType.Char);
                                        dateTime = DateTime.Now;
                                        string str2 = dateTime.ToString("HH:mm:ss");
                                        sqlParameter2.Value = (object)str2;
                                        sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString();
                                        StartupBase.SasObj.ExcuteNonQuery(sqlcmd);
                                        this.Cm_Xoa();
                                        e.Handled = true;
                                        this.Toolbar.OnIsInEditModeChanged();
                                        this.SetFocusToolbar();
                                    }
                                }
                            }
                        }
                        btnName = "btnDelete";
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F2))
                {
                    if (this.Cm_Xem != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnView") as Button).Focus();
                        this.Cm_Xem();
                        btnName = "btnView";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F9))
                {
                    if (this.Cm_Tim != null && !this.IsEditMode)
                    {
                        (this.Toolbar.FindName("btnSearch") as Button).Focus();
                        this.Cm_Tim();
                        btnName = "btnSearch";
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.F11))
                {
                    if (this.BindingSasObj != null && StartUpTrans.DmctInfo != null && (StartUpTrans.DmctInfo["f11"].ToString().Equals("1") && this.Cm_Nhan != null) && (!this.IsEditMode && StartUpTrans.DsTrans.Tables[0].Rows.Count > 1 && ExMessageBox.Show(-1135, FormTrans.SasO, "Có chắc chắn chạy Post tự động không?", "Thông báo", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes))
                    {
                        this.IsSequenceSave = true;
                        for (int index = 1; index < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index)
                        {
                            string str = StartUpTrans.DmctInfo["procedure"].ToString().Trim();
                            if (!string.IsNullOrEmpty(str))
                                this.BindingSasObj.ExcuteNonQuery(new SqlCommand(str.ToLower().Replace("@stt_rec", "'" + StartUpTrans.DsTrans.Tables[0].Rows[index]["stt_rec"].ToString().Trim() + "'")));
                        }
                        this.IsSequenceSave = false;
                        e.Handled = true;
                    }
                }
                else if (Keyboard.IsKeyDown(Key.D1) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                {
                    if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables.Count > 0 && StartUpTrans.DsTrans.Tables[0].Columns.Contains("stt_rec"))
                    {
                        Clipboard.SetText(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                        int num = (int)MessageBox.Show(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString());
                    }
                }
                else if (Keyboard.IsKeyDown(Key.D4) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                {
                    if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables.Count > 0 && StartUpTrans.DsTrans.Tables[0].Columns.Contains("stt_rec"))
                    {
                        e.Handled = true;
                        SaveFileDialog saveFileDialog = new SaveFileDialog();
                        saveFileDialog.Filter = "Xml|*.xml";
                        bool? nullable = saveFileDialog.ShowDialog();
                        if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                        {
                            try
                            {
                                DataSet dataSet = StartUpTrans.DsTrans.Clone();
                                for (int index1 = 0; index1 < dataSet.Tables.Count; ++index1)
                                {
                                    for (int index2 = 0; index2 < StartUpTrans.DsTrans.Tables[index1].DefaultView.Count; ++index2)
                                        dataSet.Tables[index1].Rows.Add(StartUpTrans.DsTrans.Tables[index1].DefaultView[index2].Row.ItemArray);
                                }
                                dataSet.WriteXml(saveFileDialog.FileName, XmlWriteMode.WriteSchema);
                                int num = (int)ExMessageBox.Show(-2690, FormTrans.SasO, "Thành công!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                            catch (Exception ex)
                            {
                                ErrorLog.CatchMessage(ex);
                            }
                        }
                    }
                }
                else if (this.IsEditMode && Keyboard.IsKeyDown(Key.D5) && (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)) && (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables.Count > 0 && (StartUpTrans.DsTrans.Tables[0].Columns.Contains("stt_rec") && StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_ct"))))
                {
                    e.Handled = true;
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Xml|*.xml";
                    openFileDialog.Multiselect = false;
                    bool? nullable = openFileDialog.ShowDialog();
                    if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                    {
                        try
                        {
                            DataSet dataSet = new DataSet();
                            int num1 = (int)dataSet.ReadXml(openFileDialog.FileName);
                            bool flag = true;
                            if (dataSet.Tables.Count != StartUpTrans.DsTrans.Tables.Count)
                            {
                                int num2 = (int)ExMessageBox.Show(-2695, FormTrans.SasO, "Số lượng bảng dữ liệu không đúng", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                flag = false;
                            }
                            else
                            {
                                if (!dataSet.Tables[0].Columns.Contains("ma_ct") || !dataSet.Tables[0].Rows[0]["ma_ct"].ToString().Trim().Equals(StartUpTrans.Ma_ct))
                                {
                                    int num2 = (int)ExMessageBox.Show(-2700, FormTrans.SasO, "File xml không thuộc chứng từ hiện hành!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                    flag = false;
                                }
                                if (flag)
                                {
                                    for (int index1 = 0; index1 < StartUpTrans.DsTrans.Tables.Count; ++index1)
                                    {
                                        for (int index2 = 0; index2 < StartUpTrans.DsTrans.Tables[index1].Columns.Count; ++index2)
                                        {
                                            if (!dataSet.Tables[index1].Columns.Contains(StartUpTrans.DsTrans.Tables[index1].Columns[index2].ColumnName))
                                            {
                                                int num2 = (int)ExMessageBox.Show(-2695, FormTrans.SasO, "Cột [" + StartUpTrans.DsTrans.Tables[index1].Columns[index2].ColumnName + "] không tồn tại trong file xml!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                                                flag = false;
                                            }
                                        }
                                    }
                                }
                            }
                            if (flag)
                            {
                                string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
                                string rowFilter = StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter;
                                string str2 = "";
                                string str3 = "";
                                if (StartUpTrans.DsTrans.Tables.Count > 1)
                                    str2 = StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter;
                                if (StartUpTrans.DsTrans.Tables.Count > 2)
                                    str3 = StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter;
                                dataSet.Tables[0].Rows[0]["stt_rec"] = (object)str1;
                                if (StartUpTrans.DsTrans.Tables.Count > 1)
                                {
                                    while (StartUpTrans.DsTrans.Tables[1].DefaultView.Count > 0)
                                    {
                                        StartUpTrans.DsTrans.Tables[1].Rows.Remove(StartUpTrans.DsTrans.Tables[1].DefaultView[0].Row);
                                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = str2;
                                    }
                                    for (int index = 0; index < dataSet.Tables[1].Rows.Count; ++index)
                                    {
                                        dataSet.Tables[1].Rows[index]["stt_rec"] = (object)str1;
                                        StartUpTrans.DsTrans.Tables[1].Rows.Add(dataSet.Tables[1].Rows[index].ItemArray);
                                        StartUpTrans.DsTrans.Tables[1].DefaultView.RowFilter = str2;
                                    }
                                }
                                if (StartUpTrans.DsTrans.Tables.Count > 2)
                                {
                                    while (StartUpTrans.DsTrans.Tables[2].DefaultView.Count > 0)
                                    {
                                        StartUpTrans.DsTrans.Tables[2].Rows.Remove(StartUpTrans.DsTrans.Tables[2].DefaultView[0].Row);
                                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = str3;
                                    }
                                    for (int index = 0; index < dataSet.Tables[2].Rows.Count; ++index)
                                    {
                                        dataSet.Tables[2].Rows[index]["stt_rec"] = (object)str1;
                                        StartUpTrans.DsTrans.Tables[2].Rows.Add(dataSet.Tables[2].Rows[index].ItemArray);
                                        StartUpTrans.DsTrans.Tables[2].DefaultView.RowFilter = str3;
                                    }
                                }
                                for (int index = 0; index < StartUpTrans.DsTrans.Tables[0].Columns.Count; ++index)
                                {
                                    if (string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].Columns[index].Expression))
                                        StartUpTrans.DsTrans.Tables[0].DefaultView[0].Row[index] = dataSet.Tables[0].Rows[0][StartUpTrans.DsTrans.Tables[0].Columns[index].ColumnName];
                                }
                                StartUpTrans.DsTrans.Tables[0].DefaultView.RowFilter = rowFilter;
                                int num2 = (int)ExMessageBox.Show(-2690, FormTrans.SasO, "Thành công!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                            }
                        }
                        catch (Exception ex)
                        {
                            ErrorLog.CatchMessage(ex);
                        }
                    }
                }
                if (this.EditModeEnded != null && !string.IsNullOrEmpty(btnName))
                    this.EditModeEnded((object)this, btnName, (RoutedEventArgs)e);
                this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
               {
                   if (string.IsNullOrEmpty(btnName))
                       return;
                   this.CheckND51();
                   this.C_QS_Text_Changed();
                   this.SetInvisible();
                   this.SetRound();
                   if (!this.IsNd51)
                       return;
                   if (FormTrans.currActionTask == ActionTask.Add || FormTrans.currActionTask == ActionTask.Copy || btnName.Equals("btnCopy"))
                   {
                       this.C_QS.IsReadOnly = false;
                       this.C_NgayHT.IsReadOnly = false;
                       this.C_QS.IsTabStop = true;
                       this.C_NgayHT.IsTabStop = true;
                   }
                   else
                   {
                       this.C_QS.IsReadOnly = true;
                       this.C_NgayHT.IsReadOnly = true;
                       this.C_QS.IsTabStop = false;
                       this.C_NgayHT.IsTabStop = false;
                   }
               }));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        /// <summary>Phương thức kiểm tra trùng số chứng từ.</summary>
        /// <param name="SasObj">SasObject dùng để thao tác dự liệu.</param>
        /// <param name="ma_qs">Mã quyển sổ chứa số chứng từ cần kiểm tra.</param>
        /// <param name="so_ct">Số chứng từ cần kiểm tra.</param>
        /// <param name="stt_rec">stt_rec của phiếu để xác định thêm mới hay sửa phiếu</param>
        /// <returns>Kết quả kiểm tra.(True- trùng,False- không trùng)</returns>
        public bool CheckValidSoct(SasObject SasObj, string ma_qs, string so_ct, string stt_rec)
        {
            string format = "{0}"; 
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Check_Valid_Store.Length <= 3 ? string.Format(format, (object)nameof(CheckValidSoct)) : string.Format(format, (object)StartUpTrans.Check_Valid_Store[3]));
            sqlcmd.Parameters.Add(new SqlParameter("@ma_qs", (object)ma_qs));
            sqlcmd.Parameters.Add(new SqlParameter("@so_ct", (object)so_ct));
            sqlcmd.Parameters.Add(new SqlParameter("@stt_rec", (object)stt_rec));
            sqlcmd.CommandType = CommandType.StoredProcedure;
            return (bool)SasObj.ExcuteScalar(sqlcmd);
        }

        protected bool CheckValidSoctLT(SasObject SasObj, string ma_qs, string so_ct)
        {
            string cmdText = "Select 1 FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "' and (so_ct + 1= " + so_ct + " or exists (select 1 from cthhd where ma_qs like '" + ma_qs.Trim() + "' and  so_ct2 + 1 = " + so_ct + "))";
            return SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0].Rows.Count == 1;
        }

        /// <summary>Phương thức tạo mới số chứng từ đảm bảo không trùng.</summary>
        /// <param name="SasObj">SasObject dùng để thao tác dự liệu.</param>
        /// <param name="ma_qs">Mã quyển sổ cần tạo mới số chứng từ.</param>
        /// <returns>Số chứng từ được tạo mới.</returns>
        public string GetNewSoct(SasObject SasObj, string ma_qs)
        {
            return this.GetNewSoct(SasObj, ma_qs, false);
        }

        protected string GetNewSoct(SasObject SasObj, string ma_qs, bool isUpdateNewSoCt)
        {
            try
            {
                string cmdText;
                if (!this.IsNd51 && isUpdateNewSoCt)
                {
                    string format = "EXEC  {0} '" + ma_qs.Trim() + "'";
                    cmdText = StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 4 ? string.Format(format, (object)"[GetNewSoct#2]") : string.Format(format, (object)StartUpTrans.Process_Store[4]);
                }
                else if (Convert.ToInt16(SasObj.GetOption("M_AUTO_SOCT").ToString()) == (short)1)
                {
                    cmdText = "SELECT transform, so_ct + 1 as so_ct FROM dmqs WHERE ma_qs = '" + ma_qs.Trim() + "'";
                }
                else
                {
                    string str;
                    string format = str = "EXEC  {0} '" + ma_qs.Trim() + "'";
                    cmdText = StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 3 ? string.Format(format, (object)nameof(GetNewSoct)) : string.Format(format, (object)StartUpTrans.Process_Store[3]);
                }
                DataTable table = SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0];
                if (table.Rows.Count > 0)
                {
                    DataRow row = table.Rows[0];
                    if (row[1] != null)
                    {
                        if (row[1] != DBNull.Value)
                        {
                            string str = row[1].ToString();
                            return string.Format(row[0].ToString(), (object)Convert.ToDouble(str));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return "";
        }

        protected void UpdateNewSoCt(SasObject SasObj, string ma_qs)
        {
            try
            {
                if (this.C_NgayHT == null || this.C_So_ct == null)
                    return;
                string cmdText = "UPDATE dmqs SET so_ct = " + this.C_So_ct.Text + ", ngay_ct = '" + string.Format("{0:yyyyMMdd}", (object)this.C_NgayHT.dValue) + "' WHERE ma_qs='" + ma_qs.Trim() + "';";
                SasObj.ExcuteNonQuery(new SqlCommand(cmdText));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        protected void UpdateNewNgayCt(SasObject SasObj, string ma_qs, Decimal so_ct)
        {
            try
            {
                if (this.C_NgayHT == null)
                    return;
                string cmdText = "IF EXISTS(SELECT 1 FROM dmqs WHERE ma_qs='" + ma_qs.Trim() + "' AND so_ct = " + (object)so_ct + ")" + "UPDATE dmqs SET ngay_ct = '" + string.Format("{0:yyyyMMdd}", (object)this.C_NgayHT.dValue) + "' WHERE ma_qs='" + ma_qs.Trim() + "';";
                SasObj.ExcuteNonQuery(new SqlCommand(cmdText));
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        /// <summary>Lấy số chứng từ cuối cùng trong một quyển sổ.</summary>
        /// <param name="SasObj">SasObject dùng để thao tác dự liệu.</param>
        /// <param name="ma_qs">Mã quyển sổ cần lấy số chứng từ.</param>
        /// <returns>số chứng từ cuối cùng trong quyển sổ.</returns>
        public string GetLastSoct(SasObject SasObj, string ma_qs)
        {
            try
            {
                string cmdText = "EXEC GetLastSoct '" + ma_qs.Trim() + "';";
                DataRow row = SasObj.ExcuteReader(new SqlCommand(cmdText)).Tables[0].Rows[0];
                if (row[0] != null)
                {
                    if (row[0] != DBNull.Value)
                        return row[0].ToString();
                }
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return "";
        }

        /// <summary>
        /// Lấy iRow của phiếu(hàn dùng để sắp xếp lại các phiếu theo ngày sau khi thêm, sửa).
        /// </summary>
        /// <param name="tbView">TableView chứa phiếu</param>
        /// <param name="stt_rec">stt_rec của phiếu</param>
        /// <returns>vị trị của row cần xét trang dataview.</returns>
        protected int GetiRow(DataTable tbView, string stt_rec)
        {
            try
            {
                DataTable dataTable1 = tbView.Copy();
                dataTable1.DefaultView.Sort = tbView.DefaultView.Sort;
                DataTable dataTable2 = dataTable1;
                return dataTable2.Rows.IndexOf(dataTable2.Select("stt_rec= '" + stt_rec + "'")[0]);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            return -1;
        }

        /// <summary>Mặc định quyển sổ khi thêm mới</summary>
        /// <param name="SasObj"></param>
        /// <param name="ma_ct"></param>
        /// <returns></returns>
        protected string GetDMQS(SasObject SasObj, string ma_ct, DateTime ngay_ct, int user_id)
        {
            string format = "exec {0} @ma_ct, @ngay_ct, @user_id";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 5 ? string.Format(format, (object)"GetDmQs") : string.Format(format, (object)StartUpTrans.Process_Store[5]));
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.VarChar).Value = (object)ma_ct;
            sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.DateTime).Value = (object)ngay_ct;
            sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)user_id;
            return SasObj.ExcuteReader(sqlcmd).Tables[0].Rows[0]["ma_qs"].ToString().Trim();
        }

        protected string GetDMQS(
          SasObject SasObj,
          string ma_ct,
          DateTime ngay_ct,
          int user_id,
          string _ma_qs)
        {
            SqlCommand sqlcmd = new SqlCommand("SELECT [status] FROM dmqs WHERE ma_qs = @ma_qs " + (" AND ((EXISTS (SELECT * FROM dmuserqs WHERE [user_id] = " + FormTrans.SasO.UserInfo.Rows[0][nameof(user_id)].ToString() + ") AND EXISTS (SELECT * FROM dmuserqs WHERE dmqs.ma_qs = dmuserqs.ma_qs AND [user_id] = " + FormTrans.SasO.UserInfo.Rows[0][nameof(user_id)].ToString() + ")) OR NOT EXISTS (SELECT * FROM dmuserqs WHERE [user_id] =  " + FormTrans.SasO.UserInfo.Rows[0][nameof(user_id)].ToString() + ") OR NOT EXISTS(SELECT 1 FROM dmuserqs WHERE dmqs.ma_qs = ma_qs))") + "AND ngay_qs1 <= @ngay_ct");
            sqlcmd.Parameters.Add("@ma_qs", SqlDbType.VarChar).Value = (object)_ma_qs;
            sqlcmd.Parameters.Add("@ngay_ct", SqlDbType.DateTime).Value = (object)ngay_ct;
            sqlcmd.Parameters.Add("@user_id", SqlDbType.Int).Value = (object)user_id;
            DataTable table = SasObj.ExcuteReader(sqlcmd).Tables[0];
            return table.Rows.Count > 0 && table.Rows[0]["status"].ToString() == "1" ? _ma_qs : this.GetDMQS(SasObj, ma_ct, ngay_ct, user_id);
        }

        protected virtual bool CheckCanDelete()
        {
            return true;
        }

        public void CheckBoxVoucher()
        {
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (StartUpTrans.DsTrans == null || !StartUpTrans.DsTrans.Tables[0].DefaultView.Table.Columns.Contains("hd_thue"))
                   return;
               if (StartUpTrans.DsTrans.Tables[0].DefaultView[0]["hd_thue"].Equals((object)"1") || StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_ct"].ToString().Trim().ToUpper().Equals("QL1"))
                   (this.Toolbar.FindName("cb_hd_thue") as CheckBox).IsChecked = new bool?(true);
               else
                   (this.Toolbar.FindName("cb_hd_thue") as CheckBox).IsChecked = new bool?(false);
           }));
        }

        protected virtual void InsertRecord(Action action, DataGridView GrdCt, string activecolumn)
        {
            string format = StartupBase.SasObj.GetSysvar("M_FORMAT_stt_rec0").ToString();
            int result = 1;
            DataRowView r = (GrdCt.ActiveRecord as DataRecord).DataItem as DataRowView;
            string stt_rec = r["stt_rec"].ToString();
            string stt_rec0 = r["stt_rec0"].ToString();
            DataView dataView = r.DataView;
            action();
            foreach (DataRowView dataRowView in dataView.Cast<DataRowView>().Where<DataRowView>((Func<DataRowView, bool>)(x => !(x["stt_rec"].ToString() != stt_rec) && x["stt_rec0"].ToString().CompareTo(stt_rec0) >= 0)).ToArray<DataRowView>())
            {
                int.TryParse(dataRowView["stt_rec0"].ToString(), out result);
                ++result;
                dataRowView["stt_rec0"] = (object)string.Format(format, (object)result);
            }
            DataRowView r1 = (GrdCt.Records[GrdCt.Records.Count - 1] as DataRecord).DataItem as DataRowView;
            r1["stt_rec0"] = (object)stt_rec0;
            DataView CtTableView = GrdCt.DataSource as DataView;
            this.Dispatcher.BeginInvoke((Delegate)new Action(() =>
           {
               if (GrdCt.Records.Count > 0)
               {
                   DataTable freeCodeInfo = StartupBase.SasObj.GetFreeCodeInfo(StartUpTrans.Ma_ct);
                   string index = "field_ct";
                   foreach (DataRow row in (InternalDataCollectionBase)freeCodeInfo.Rows)
                   {
                       if (!string.IsNullOrEmpty(row[index].ToString().Trim()) && row["carry"].ToString().Trim() == "1" && CtTableView.ToTable().Columns.Contains(row[index].ToString().Trim()))
                           r1[row[index].ToString().Trim()] = r[row[index].ToString().Trim()];
                   }
               }
               GrdCt.ActiveRecord = GrdCt.Records.FirstOrDefault<Infragistics.Windows.DataPresenter.Record>((Func<Infragistics.Windows.DataPresenter.Record, bool>)(x => ((x as DataRecord).DataItem as DataRowView)["stt_rec0"].ToString() == stt_rec0));
               if (GrdCt.ActiveRecord == null)
                   return;
               GrdCt.ActiveCell = (GrdCt.ActiveRecord as DataRecord).Cells[0];
           }), DispatcherPriority.Background);
        }

        protected void MoveUp(DataRecord rec1)
        {
            if (rec1 == null)
                return;
            DataGridView dataPresenter = rec1.DataPresenter as DataGridView;
            if (rec1.Index <= 0)
                return;
            DataRecord record = dataPresenter.Records[rec1.Index - 1] as DataRecord;
            this.ChangeRecord(rec1, record);
        }

        protected void MoveDown(DataRecord rec1)
        {
            if (rec1 == null)
                return;
            DataGridView dataPresenter = rec1.DataPresenter as DataGridView;
            if (rec1.Index >= dataPresenter.Records.Count - 1)
                return;
            DataRecord record = dataPresenter.Records[rec1.Index + 1] as DataRecord;
            this.ChangeRecord(rec1, record);
        }

        private void ChangeRecord(DataRecord rec1, DataRecord rec2)
        {
            DataRowView dataItem1 = rec1.DataItem as DataRowView;
            DataRowView dataItem2 = rec2.DataItem as DataRowView;
            string str = dataItem1["stt_rec0"].ToString();
            dataItem1["stt_rec0"] = dataItem2["stt_rec0"];
            dataItem2["stt_rec0"] = (object)str;
            DataGridView dataPresenter = rec1.DataPresenter as DataGridView;
            dataItem1.DataView.Table.AcceptChanges();
            dataPresenter.ActiveRecord = (Infragistics.Windows.DataPresenter.Record)rec2;
        }

        private bool CheckDC()
        {
            SqlCommand sqlcmd = new SqlCommand("exec CheckSttRecTT @stt_rec");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = (object)StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            return !((int)StartupBase.SasObj.ExcuteScalar(sqlcmd)).Equals(0);
        }

        public delegate void Execute();

        public delegate void OnEditMode(object sender, string menuItemName, RoutedEventArgs e);


    }
}
