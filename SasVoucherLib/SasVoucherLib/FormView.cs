using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using Microsoft.Win32;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasLib;
using System;
using System.Collections;
using System.Data;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace SasVoucherLib
{
  public class FormView
  {
    public FrmView frmBrw = new FrmView();
    public string constCtFilter = "";
    public string TongPsVND = "";
    public string TongPSNT = "";
    public string ListFieldSum = "";
    public string TongCongLabel = "";
    private string strNgayHienTai = DateTime.Now.Date.ToShortDateString();
    public int TongSoCt;
    public DataView ObrowseView;
    public DataView ObrowseViewCt;
    private string sRelationKey;
    protected SasObject _SasObj;

    /// <summary>Record được Actived</summary>
    public DataRecord ActiveRecord
    {
      get
      {
        return (DataRecord) this.frmBrw.GrdBrowse.ActiveRecord;
      }
    }

    /// <summary>DataGridView Ph của form</summary>
    public XamDataGrid DataGrid
    {
      get
      {
        return (XamDataGrid) this.frmBrw.GrdBrowse;
      }
    }

    protected string _fields { get; set; }

    private string BrowseFields
    {
      get
      {
        return this._fields;
      }
      set
      {
        this._fields = value;
      }
    }

    protected string _fieldsCt { get; set; }

    private string BrowseFieldsCt
    {
      get
      {
        return this._fieldsCt;
      }
      set
      {
        this._fieldsCt = value;
      }
    }

    private SasObject SasObject
    {
      get
      {
        return this._SasObj;
      }
      set
      {
        this._SasObj = value;
      }
    }

    public event FormView.Form_Loaded FormBrowse_Loaded;

    public event FormView.GridKeyUp_Esc Esc;

    public event FormView.GridKeyUp_F1 F1;

    public event FormView.GridKeyUp_F2 F2;

    public event FormView.GridKeyUp_F3 F3;

    public event FormView.GridKeyUp_F4 F4;

    public event FormView.GridKeyUp_F5 F5;

    public event FormView.GridKeyUp_F6 F6;

    public event FormView.GridKeyUp_F7 F7;

    public event FormView.GridKeyUp_F8 F8;

    public event FormView.GridKeyUp_F9 F9;

    public event FormView.GridKeyUp_F10 F10;

    public event FormView.GridKeyUp_F11 F11;

    public event FormView.GridKeyUp_F12 F12;

    public event FormView.GridKeyUp_All KeyUp_All;

    public FormView(
      SasObject SasObj,
      DataView dv,
      DataView dvCt,
      string strBrowse,
      string strBrowseCt,
      string sFieldKey)
    {
      this.ObrowseView = dv.Table.Copy().DefaultView;
      this.ObrowseViewCt = dvCt.Table.Copy().DefaultView;
      try
      {
        if (string.IsNullOrEmpty(this.ObrowseViewCt.Sort))
          this.ObrowseViewCt.Sort = "stt_rec0 asc";
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
      this._SasObj = SasObj;
      this.sRelationKey = sFieldKey;
      this.BrowseFields = strBrowse;
      this.BrowseFieldsCt = strBrowseCt;
      this.InitFormBrowse(ref this.frmBrw, ref SasObj, dv, dvCt, strBrowse, strBrowseCt, sFieldKey);
      this.frmBrw.Loaded += new RoutedEventHandler(this.frmBrw_Loaded);
      this.frmBrw.GrdBrowse.RecordActivated += new EventHandler<RecordActivatedEventArgs>(this.GrdBrowse_RecordActivated);
      this.frmBrw.GrdBrowse.RecordFilterChanged += new EventHandler<RecordFilterChangedEventArgs>(this.GrdBrowse_RecordFilterChanged);
      this.frmBrw.GrdBrowse.FieldLayoutSettings.AllowFieldMoving = AllowFieldMoving.WithinLogicalRow;
      this.frmBrw.GrdBrowse.FieldSettings.AllowFixing = AllowFieldFixing.Near;
      this.frmBrw.GrdBrowseCt.FieldLayoutSettings.AllowFieldMoving = AllowFieldMoving.WithinLogicalRow;
      this.frmBrw.GrdBrowseCt.FieldSettings.AllowFixing = AllowFieldFixing.Near;
      if (this.frmBrw.GrdBrowse.FieldLayouts.Count <= 0)
        return;
      bool flag = false;
      for (int index = 0; index < this.frmBrw.GrdBrowse.FieldLayouts[0].Fields.Count && !flag; ++index)
      {
        if (this.frmBrw.GrdBrowse.FieldLayouts[0].Fields[index].Name == "ma_nt" && dv.ToTable().Select("ma_nt <> '' AND ma_nt <> '" + StartupBase.M_MA_NT0 + "'").Length > 0)
          flag = true;
      }
      if (flag)
        return;
      DataTable columnForeignCurrency = StartUpTrans.GetListColumnForeignCurrency();
      columnForeignCurrency.Rows.Add((object) "ma_nt");
      columnForeignCurrency.Rows.Add((object) "ty_gia");
      columnForeignCurrency.Rows.Add((object) "ty_giaf");
      for (int index = 0; index < this.frmBrw.GrdBrowse.FieldLayouts[0].Fields.Count; ++index)
      {
        if (columnForeignCurrency.Select("column_name = '" + this.frmBrw.GrdBrowse.FieldLayouts[0].Fields[index].Name + "'").Length > 0)
          this.frmBrw.GrdBrowse.FieldLayouts[0].Fields[index].Visibility = Visibility.Collapsed;
      }
      for (int index = 0; index < this.frmBrw.GrdBrowseCt.FieldLayouts[0].Fields.Count; ++index)
      {
        if (columnForeignCurrency.Select("column_name = '" + this.frmBrw.GrdBrowseCt.FieldLayouts[0].Fields[index].Name + "'").Length > 0)
          this.frmBrw.GrdBrowseCt.FieldLayouts[0].Fields[index].Visibility = Visibility.Collapsed;
      }
    }

    private void GrdBrowse_RecordFilterChanged(object sender, RecordFilterChangedEventArgs e)
    {
      this.frmBrw.GrdBrowseCt.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(() => this.frmBrw.GrdBrowseCt.DataSource = (IEnumerable) this.ObrowseViewCt.Table.Clone().DefaultView));
    }

    private void GrdBrowse_RecordActivated(object sender, RecordActivatedEventArgs e)
    {
      try
      {
        BasicGridView basicGridView = sender as BasicGridView;
        if (basicGridView.ActiveRecord == null || basicGridView.ActiveRecord.Index < 0 || basicGridView.ActiveRecord.RecordType != RecordType.DataRecord)
          return;
        this.ObrowseViewCt.RowFilter = "";
        string str = "1 =1 ";
        string sRelationKey = this.sRelationKey;
        char[] chArray = new char[1]{ ';' };
        foreach (string index in sRelationKey.Split(chArray))
          str += string.Format("{0} {1} = '{2}'", str == "" ? (object) "" : (object) " and ", (object) index, (object) ((basicGridView.ActiveRecord as DataRecord).DataItem as DataRowView)[index].ToString());
        this.ObrowseViewCt.RowFilter += str;
        if (this.constCtFilter.Length > 0)
        {
          if (str.Length > 0)
          {
            DataView obrowseViewCt = this.ObrowseViewCt;
            obrowseViewCt.RowFilter = obrowseViewCt.RowFilter + " and " + this.constCtFilter;
          }
          else
            this.ObrowseViewCt.RowFilter = this.constCtFilter;
        }
        this.frmBrw.GrdBrowseCt.DataSource = (IEnumerable) this.ObrowseViewCt;
      }
      catch (Exception ex)
      {
        ErrorLog.CatchMessage(ex);
      }
    }

    private void frmBrw_Loaded(object sender, RoutedEventArgs e)
    {
      this.frmBrw.GrdBrowse.DataSource = (IEnumerable) this.ObrowseView;
      this.frmBrw.GrdBrowseCt.DataSource = (IEnumerable) this.ObrowseViewCt;
      if (this.frmBrw.GrdBrowse.Records.Count > 0)
        Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate) new Action(this.SetDefaultActiveRecord));
      this.frmBrw.PreviewKeyDown += new KeyEventHandler(this.frmBrw_PreviewKeyDown);
      this.frmBrw.UpdateLayout();
      if (this.FormBrowse_Loaded != null)
      {
        RoutedEventArgs routedEventArgs = new RoutedEventArgs();
        this.FormBrowse_Loaded(sender, (EventArgs) routedEventArgs);
      }
      this.TongSoCt = this.ObrowseView.Count;
      string[] strArray = this.ListFieldSum.Split(';');
      string str1 = "";
      for (int index = 0; index < strArray.Length; ++index)
      {
        string name = strArray[index].Trim();
        if (this.frmBrw.GrdBrowse.DefaultFieldLayout.Fields.IndexOf(name) != -1)
        {
          string str2 = this.frmBrw.GrdBrowse.DefaultFieldLayout.Fields[name].Tag.ToString().Trim();
          string format = str2.Substring(3, str2.Length - 4);
          object obj = this.ObrowseView.Table.Compute("Sum(" + name + ")", "1=1");
          double result = 0.0;
          double.TryParse(obj.ToString(), out result);
          string str3 = result.ToString(format);
          str1 += str3;
          if (index < strArray.Length - 1)
            str1 += " /";
        }
      }
      this.frmBrw.numSoChungTu.Text = this.TongSoCt.ToString();
      this.frmBrw.lblTongPhatSinh.Text = string.IsNullOrEmpty(this.TongCongLabel) ? "Tổng ps" : this.TongCongLabel;
      this.frmBrw.numTongPhatSinh.Text = str1;
      this.frmBrw.lblNgay.Text = this.strNgayHienTai;
      this.frmBrw.GrdBrowse.Focus();
    }

    public void SetDefaultActiveRecord()
    {
      this.frmBrw.GrdBrowse.Focus();
      Record record = (Record) null;
      if (StartUpTrans.DsTrans != null && StartUpTrans.DsTrans.Tables.Count > 0 && StartUpTrans.DsTrans.Tables[0].Rows.Count > 0)
      {
        string stt_rec = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
        record = this.frmBrw.GrdBrowse.Records.FirstOrDefault<Record>((Func<Record, bool>) (x => ((x as DataRecord).DataItem as DataRowView)["stt_rec"].ToString() == stt_rec));
      }
      if (record == null)
        this.frmBrw.GrdBrowse.ActiveRecord = (Record) (this.frmBrw.GrdBrowse.Records[0] as DataRecord);
      else
        this.frmBrw.GrdBrowse.ActiveRecord = record;
    }

    private void frmBrw_PreviewKeyDown(object sender, KeyEventArgs e)
    {
      switch (e.Key)
      {
        case Key.Escape:
          if (this.Esc != null)
          {
            this.Esc((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F1:
          if (this.F1 != null)
          {
            this.F1((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F2:
          if (this.F2 != null)
          {
            this.F2((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F3:
          if (this.F3 != null)
          {
            this.F3((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F4:
          if (this.F4 != null)
          {
            this.F4((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F5:
          if (this.F5 != null)
          {
            this.F5((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F6:
          if (this.F6 != null)
          {
            this.F6((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F7:
          if (this.F7 != null)
          {
            this.F7((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F8:
          if (this.F8 != null)
          {
            this.F8((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F9:
          if (this.F9 != null)
          {
            this.F9((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F10:
          if (this.F10 != null)
          {
            this.F10((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F11:
          if (this.F11 != null)
          {
            this.F11((object) this, new EventArgs());
            break;
          }
          break;
        case Key.F12:
          if (this.F12 != null)
          {
            this.F12((object) this, new EventArgs());
            break;
          }
          break;
      }
      if (e.Key == Key.S && Keyboard.IsKeyDown(Key.LeftCtrl))
        this.ExportToExcel();
      if (e.Key == Key.Tab && Keyboard.IsKeyDown(Key.LeftCtrl))
      {
        if (this.frmBrw.GrdBrowse.IsFocused)
        {
          if (this.frmBrw.GrdBrowseCt.Records.Count > 0)
            this.frmBrw.GrdBrowseCt.ActiveRecord = this.frmBrw.GrdBrowseCt.Records[0];
        }
        else if (this.frmBrw.GrdBrowse.Records.Count > 0)
          this.frmBrw.GrdBrowse.ActiveRecord = this.frmBrw.GrdBrowse.Records[0];
      }
      if (!Keyboard.IsKeyDown(Key.C) || !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl) || !Keyboard.IsKeyDown(Key.RightShift) && !Keyboard.IsKeyDown(Key.LeftShift))
        return;
      if (this.frmBrw.GrdBrowse.IsKeyboardFocusWithin)
        Clipboard.SetDataObject((object) SysFunc.GenerateStringBrowse((XamDataGrid) this.frmBrw.GrdBrowse, this.BrowseFields));
      if (!this.frmBrw.GrdBrowseCt.IsKeyboardFocusWithin)
        return;
      Clipboard.SetDataObject((object) SysFunc.GenerateStringBrowse((XamDataGrid) this.frmBrw.GrdBrowseCt, this.BrowseFieldsCt));
    }

    /// <summary>Gọi showDiaglog formview.</summary>
    public void ShowDialog()
    {
      this.frmBrw.ShowDialog();
    }

    /// <summary>Gọi show formview</summary>
    public void Show()
    {
      this.frmBrw.Show();
    }

    private void InitFormBrowse(
      ref FrmView oForm,
      ref SasObject SasObj,
      DataView dt,
      DataView dtCt,
      string strBrowse,
      string strBrowseCt,
      string sFieldsKey)
    {
      oForm.Language = XmlLanguage.GetLanguage(Thread.CurrentThread.CurrentCulture.Name);
      FieldLayout fieldLayout1 = SysFunc.CreateFieldLayout(SasObj, oForm.GrdBrowse, strBrowse, dt.Table);
      oForm.GrdBrowse.FieldLayouts.Add(fieldLayout1);
      SysFunc.CreateSumFieldList(SasObj, oForm.GrdBrowse, strBrowse);
      FieldLayout fieldLayout2 = SysFunc.CreateFieldLayout(SasObj, oForm.GrdBrowseCt, strBrowseCt, dtCt.Table);
      oForm.GrdBrowseCt.FieldLayouts.Add(fieldLayout2);
      SysFunc.CreateSumFieldList(SasObj, oForm.GrdBrowseCt, strBrowseCt);
      oForm.Left = 0.0;
      oForm.Top = 0.0;
      oForm.GrdBrowse.KeyUp += new KeyEventHandler(this.GrdBrowse_KeyUp);
    }

    private void Form_KeyDown(object sender, KeyEventArgs e)
    {
    }

    private void GrdBrowse_KeyUp(object sender, KeyEventArgs e)
    {
      if (this.KeyUp_All != null)
        return;
      switch (e.Key)
      {
        case Key.F1:
          if (this.F1 == null)
            break;
          this.F1((object) this, new EventArgs());
          break;
        case Key.F2:
          if (this.F2 == null)
            break;
          this.F2((object) this, new EventArgs());
          break;
        case Key.F3:
          if (this.F3 == null)
            break;
          this.F3((object) this, new EventArgs());
          break;
        case Key.F4:
          if (this.F4 == null)
            break;
          this.F4((object) this, new EventArgs());
          break;
        case Key.F5:
          if (this.F5 == null)
            break;
          this.F5((object) this, new EventArgs());
          break;
        case Key.F6:
          if (this.F6 == null)
            break;
          this.F6((object) this, new EventArgs());
          break;
        case Key.F7:
          if (this.F7 == null)
            break;
          this.F7((object) this, new EventArgs());
          break;
        case Key.F8:
          if (this.F8 == null)
            break;
          this.F8((object) this, new EventArgs());
          break;
        case Key.F9:
          if (this.F9 == null)
            break;
          this.F9((object) this, new EventArgs());
          break;
        case Key.F10:
          if (this.F10 == null)
            break;
          this.F10((object) this, new EventArgs());
          break;
        case Key.F11:
          if (this.F11 == null)
            break;
          this.F11((object) this, new EventArgs());
          break;
        case Key.F12:
          if (this.F12 == null)
            break;
          this.F12((object) this, new EventArgs());
          break;
      }
    }

    private void ExportToExcel()
    {
      SaveFileDialog saveFileDialog = new SaveFileDialog();
      saveFileDialog.Filter = "Excel 2003 and before format|*.xls|Excel 2007|*.xlsx";
      saveFileDialog.Title = "Xuất ra excel";
      saveFileDialog.ShowDialog();
      if (string.IsNullOrEmpty(saveFileDialog.FileName))
        return;
      new GridViewExporter2().Export((XamDataGrid) this.frmBrw.oBrowse, (XamDataGrid) this.frmBrw.oBrowseCt, this.sRelationKey, saveFileDialog.FileName, false);
    }

    public delegate void Form_Loaded(object sender, EventArgs e);

    public delegate void GridKeyUp_Esc(object sender, EventArgs e);

    public delegate void GridKeyUp_F1(object sender, EventArgs e);

    public delegate void GridKeyUp_F2(object sender, EventArgs e);

    public delegate void GridKeyUp_F3(object sender, EventArgs e);

    public delegate void GridKeyUp_F4(object sender, EventArgs e);

    public delegate void GridKeyUp_F5(object sender, EventArgs e);

    public delegate void GridKeyUp_F6(object sender, EventArgs e);

    public delegate void GridKeyUp_F7(object sender, EventArgs e);

    public delegate void GridKeyUp_F8(object sender, EventArgs e);

    public delegate void GridKeyUp_F9(object sender, EventArgs e);

    public delegate void GridKeyUp_F10(object sender, EventArgs e);

    public delegate void GridKeyUp_F11(object sender, EventArgs e);

    public delegate void GridKeyUp_F12(object sender, EventArgs e);

    public delegate void GridKeyUp_All(object sender, EventArgs e);
  }
}
