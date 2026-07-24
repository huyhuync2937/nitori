using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace Socthda
{
  public partial class FrmView : Form
  {
    public bool isOk;
    public DataSet dsHdm;

    public FrmView(string filter)
    {
      this.LanguageID = "Socthda_FrmChonHd";
      this.InitializeComponent();
      SysFunc.LoadIcon((Window) this);
      this.dsHdm = StartUp.GetHdb(filter);
      this.GrdBrowse.DataSource = (IEnumerable) this.dsHdm.Tables[0].DefaultView;
      this.GrdBrowseCt.DataSource = (IEnumerable) this.dsHdm.Tables[1].DefaultView;
      this.GrdBrowse.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowse, StartUp.stringBrowse3, this.dsHdm.Tables[0]));
      SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowse, StartUp.stringBrowse3);
      this.GrdBrowseCt.FieldLayouts.Add(SysFunc.CreateFieldLayout(StartupBase.SasObj, this.GrdBrowseCt, StartUp.stringBrowse4, this.dsHdm.Tables[1]));
      SysFunc.CreateSumFieldList(StartupBase.SasObj, this.GrdBrowseCt, StartUp.stringBrowse4);
      this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() =>
      {
        if (this.GrdBrowse.Records.Count <= 0)
          return;
        this.GrdBrowse.ActiveRecord = this.GrdBrowse.Records[0];
      }));
    }

    private void GrdBrowse_RecordActivated(object sender, RecordActivatedEventArgs e)
    {
      try
      {
        BasicGridView basicGridView = sender as BasicGridView;
        if (basicGridView.ActiveRecord == null || (basicGridView.ActiveRecord.Index < 0 || basicGridView.ActiveRecord.RecordType != RecordType.DataRecord))
          return;
        this.dsHdm.Tables[1].DefaultView.RowFilter = "";
        string str = "1 =1 ";
        this.dsHdm.Tables[1].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object) "" : (object) " and ", (object) "stt_rec", (object) ((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
        this.GrdBrowseCt.DataSource = (IEnumerable) this.dsHdm.Tables[1].DefaultView;
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      this.isOk = false;
    }

    private void Form_KeyUp(object sender, KeyEventArgs e)
    {
      if (!e.Key.Equals((object) Key.Escape))
        return;
      this.Close();
    }

    private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
    {
      if (this.GrdBrowse.ActiveRecord == null || this.GrdBrowse.ActiveRecord.RecordType != RecordType.DataRecord)
        return;
      string str = "1 =1 ";
      this.dsHdm.Tables[0].DefaultView.RowFilter += str + string.Format("{0} {1} = '{2}'", str == "" ? (object) "" : (object) " and ", (object) "stt_rec", (object) ((this.GrdBrowse.ActiveRecord as DataRecord).DataItem as DataRowView)["stt_rec"].ToString());
      this.isOk = true;
      this.Close();
    }
  }
}
