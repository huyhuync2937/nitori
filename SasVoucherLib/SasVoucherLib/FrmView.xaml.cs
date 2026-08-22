using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasFormBrowes;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Threading;

namespace SasVoucherLib
{
  /// <summary>Interaction logic for FrmBrowse.xaml</summary>
  /// <summary>FrmView</summary>
  public partial class FrmView : Form
  {
    public static readonly DependencyProperty IsShowStatusBarProperty = DependencyProperty.Register(nameof (IsShowStatusBar), typeof (bool), typeof (FrmView), (PropertyMetadata) new UIPropertyMetadata((object) true));
    public BasicGridView oBrowse;
    public BasicGridView oBrowseCt;
   
    public bool IsShowStatusBar
    {
      get
      {
        return (bool) this.GetValue(FrmView.IsShowStatusBarProperty);
      }
      set
      {
        this.SetValue(FrmView.IsShowStatusBarProperty, (object) value);
        if (!value)
          return;
        this.statusBorder.Visibility = Visibility.Collapsed;
      }
    }

    public FrmView()
    {
      this.InitializeComponent();
      this.oBrowse = this.GrdBrowse;
      this.oBrowseCt = this.GrdBrowseCt;
      SysFunc.LoadIcon((Window) this);
      if (StartupBase.SasObj == null || !StartupBase.SasObj.GetOption("M_FTRAN_MAXIMIZE").ToString().Trim().Equals("1"))
        return;
      this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.WindowState = WindowState.Maximized));
      double dWidth = this.Width;
      double dHeight = this.Height;
      this.Width = SystemParameters.WorkArea.Width;
      this.Height = SystemParameters.WorkArea.Height;
      this.StateChanged += (EventHandler) ((s, e) =>
      {
        this.Width = dWidth;
        this.Height = dHeight;
      });
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      this.Dispatcher.BeginInvoke((Delegate) new Action(() => this.ChangeLanguage()), DispatcherPriority.Background);
    }

    private void GrdBrowse_RecordUpdated(object sender, RecordUpdatedEventArgs e)
    {
    }

    public void RefreshData()
    {
      DataView dataSource = (DataView) this.GrdBrowse.DataSource;
      this.GrdBrowse.DataSource = (IEnumerable) null;
      this.GrdBrowse.DataSource = (IEnumerable) dataSource;
    }

  }
}
