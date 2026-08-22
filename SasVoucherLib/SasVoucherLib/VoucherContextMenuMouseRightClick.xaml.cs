using Microsoft.Win32;
using SasControls;
using SasFormBrowes;
using SasLib;
using System;
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

namespace SasVoucherLib
{
    /// <summary>
    /// Interaction logic for VoucherContextMenuMouseRightClick.xaml
    /// </summary>
    /// <summary>VoucherContextMenuMouseRightClick</summary>
    public partial class VoucherContextMenuMouseRightClick : ContextMenu
    {
        public static readonly DependencyProperty ToolbarProperty = DependencyProperty.Register(nameof(Toolbar), typeof(ToolBarControl), typeof(VoucherContextMenuMouseRightClick), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));

        public ToolBarControl Toolbar
        {
            get
            {
                return (ToolBarControl)this.GetValue(VoucherContextMenuMouseRightClick.ToolbarProperty);
            }
            set
            {
                this.SetValue(VoucherContextMenuMouseRightClick.ToolbarProperty, (object)value);
            }
        }

        public VoucherContextMenuMouseRightClick(ToolBarControl _toolbar)
        {
            this.InitializeComponent();
            this.Toolbar = _toolbar;
            string AppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (this.Vcmmright.Items.Count != 0)
                return;
            EnumerableRowCollection<DataRow> dmcts = StartupBase.SasObj.DmctInfo.AsEnumerable();
            string ma_ct = "";
            DataRow[] array = ((IEnumerable<DataRow>)StartupBase.SasObj.CommandInfo.Select("ma_ct NOT LIKE ''")).OrderBy<DataRow, Decimal>((Func<DataRow, Decimal>)(x =>
              {
                  ma_ct = x.Field<string>("ma_ct").Trim();
                  if (ma_ct == "DHB")
                      ma_ct = "HDB";
                  DataRow row = dmcts.FirstOrDefault<DataRow>((Func<DataRow, bool>)(a => a["ma_ct"].ToString().Trim() == ma_ct));
                  return row == null ? new Decimal(0) : row.Field<Decimal>("stt");
              })).ToArray<DataRow>();
            if (array.Length <= 0)
                return;
            string str = "";
            string sma_ct = ";";
            for (int index = 0; index < array.Length; ++index)
            {
                DataRow dr = array[index];
                if (!sma_ct.Contains(dr["ma_ct"].ToString().Trim()))
                {

                    sma_ct = sma_ct + dr["ma_ct"].ToString().Trim() + ";";
                    if (str != dr["ma_phan_he"].ToString().Trim() && !string.IsNullOrEmpty(str))
                        this.Vcmmright.Items.Add((object)new Separator());
                    ExMenuItem exMenuItem = new ExMenuItem();
                    exMenuItem.Header = (object)dr["bar"].ToString().Trim();
                    exMenuItem.Header2 = dr["bar2"].ToString().Trim();
                    exMenuItem.Click += (RoutedEventHandler)((ss, ee) =>
                   {
                       string[] parameters = new string[2] { dr["menu_id"].ToString(), "" };
                       VoucherContextMenuMouseRightClick.CallModule(dr["procedure"].ToString(), parameters, AppPath, StartupBase.SasObj.M_ProcessName, StartupBase.SasObj);
                   });
                    this.Vcmmright.Items.Add((object)exMenuItem);
                    str = dr["ma_phan_he"].ToString().Trim();
                }
            }
        }

        public static void CallModule(
          string StrExecute,
          string[] parameters,
          string AppPath,
          string ProcessName,
          SasObject SasObj)
        {
            try
            {
                SasObj.SynchroFile(".", StrExecute);
                Type type = Assembly.LoadFrom(AppPath + "\\" + StrExecute).GetType(StrExecute.Trim().Substring(0, StrExecute.Trim().Length - 3) + "StartUp");
                if (type != null || string.IsNullOrEmpty(ProcessName))
                {
                    if (type.GetMethod("SasLoader") != null)
                    {
                        IntPtr handleWaiting = SasObj.HandleWaiting;
                        if (!SysFunc.SendMessToProcess("SasProcess", AppPath + "\\" + StrExecute + ";" + StrExecute.Trim().Substring(0, StrExecute.Trim().Length - 3) + "StartUp;SasLoader" + "|" + string.Join(";", parameters) + "|" + (object)Process.GetCurrentProcess().Handle + "|" + (object)handleWaiting, AppPath, ProcessName, SasObj, handleWaiting))
                            new Process()
                            {
                                StartInfo = new ProcessStartInfo()
                                {
                                    UseShellExecute = true,
                                    WorkingDirectory = AppPath,
                                    FileName = StrExecute,
                                    Arguments = (parameters[0].ToString() + " " + parameters[1].ToString())
                                }
                            }.Start();
                    }
                    else
                        new Process()
                        {
                            StartInfo = new ProcessStartInfo()
                            {
                                UseShellExecute = true,
                                WorkingDirectory = AppPath,
                                FileName = StrExecute,
                                Arguments = (parameters[0].ToString() + " " + parameters[1].ToString())
                            }
                        }.Start();
                }
                else
                    new Process()
                    {
                        StartInfo = new ProcessStartInfo()
                        {
                            UseShellExecute = true,
                            WorkingDirectory = AppPath,
                            FileName = StrExecute,
                            Arguments = (parameters[0].ToString() + " " + parameters[1].ToString())
                        }
                    }.Start();
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show(ex.Message);
            }
        }

        private void mnuInfo_Click(object sender, RoutedEventArgs e)
        {
            this.Toolbar.vc.VoucherRow = (DataRow)null;
            if (StartUpTrans.DmctInfo != null && StartUpTrans.DsTrans != null && (StartUpTrans.DsTrans.Tables.Count > 0 && !string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
            {
                DataTable table = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT * from " + StartUpTrans.DmctInfo["m_phdbf"] + " WHERE stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'")).Tables[0];
                if (table.Rows.Count == 1)
                {
                    this.Toolbar.vc.VoucherRow = table.Rows[0];
                    this.Toolbar.vc.DmctRow = StartUpTrans.DmctInfo;
                }
            }
            if (this.Toolbar.vc.VoucherRow == null)
                return;
            this.Toolbar.vc.mnuInfo.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
        }

        private void mnuSearch_Click(object sender, RoutedEventArgs e)
        {
            IInputElement focusedElement = Keyboard.FocusedElement;
            this.Toolbar.btnSearch.Focus();
            this.Toolbar.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent, (object)this.Toolbar.btnSearch));
            focusedElement?.Focus();
        }

        private void mnuView_Click(object sender, RoutedEventArgs e)
        {
            IInputElement focusedElement = Keyboard.FocusedElement;
            this.Toolbar.btnView.Focus();
            this.Toolbar.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent, (object)this.Toolbar.btnView));
            focusedElement?.Focus();
        }

        private void mnuPrint_Click(object sender, RoutedEventArgs e)
        {
            IInputElement focusedElement = Keyboard.FocusedElement;
            this.Toolbar.btnPrint.Focus();
            this.Toolbar.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent, (object)this.Toolbar.btnPrint));
            focusedElement?.Focus();
        }

        private void mnuAddBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.DsTrans == null || StartUpTrans.DsTrans.Tables.Count <= 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString()))
                return;
            DataTable table = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT * from " + StartUpTrans.DmctInfo["m_phdbf"] + " WHERE stt_rec='" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'")).Tables[0];
            if (table.Rows.Count != 1)
                return;
            this.Toolbar.VoucherRow = table.Rows[0];
            string str1 = "";
            if (this.Toolbar.VoucherRow.Table.Columns.Contains("ngay_ct"))
                str1 = str1 + "[" + string.Format("{0:dd/MM/yyyy}", (object)Convert.ToDateTime(this.Toolbar.VoucherRow["ngay_ct"])) + "]";
            string str2 = str1 + " - ";
            if (this.Toolbar.VoucherRow.Table.Columns.Contains("so_ct"))
                str2 = str2 + "[" + this.Toolbar.VoucherRow["so_ct"].ToString().Trim() + "]";
            string ghi_chu = str2 + " - ";
            if (this.Toolbar.VoucherRow.Table.Columns.Contains("dien_giai"))
                ghi_chu += this.Toolbar.VoucherRow["dien_giai"].ToString().Trim();
            this.Toolbar.AddBookmark(ghi_chu);
        }

        private void mnuExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (StartUpTrans.DmctInfo == null || StartUpTrans.DsTrans == null || (StartUpTrans.DsTrans.Tables.Count <= 0 || string.IsNullOrEmpty(StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString())))
                return;
            string cmdText = "exec LoadVoucher @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct";
            if (StartUpTrans.Process_Store != null && StartUpTrans.Process_Store.Length > 0)
                cmdText = string.Format("exec {0} @ma_ct, @PhFilter, @CtFilter, @GtFilter, @Sl_ct", (object)StartUpTrans.Process_Store[0]);
            SqlCommand sqlcmd = new SqlCommand(cmdText);
            sqlcmd.Parameters.Add("@ma_ct", SqlDbType.Char).Value = (object)StartUpTrans.DmctInfo["ma_ct"].ToString();
            sqlcmd.Parameters.Add("@PhFilter", SqlDbType.NVarChar, 4000).Value = (object)("stt_rec = '" + StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString() + "'");
            sqlcmd.Parameters.Add("@CtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
            sqlcmd.Parameters.Add("@GtFilter", SqlDbType.NVarChar, 4000).Value = (object)"1=1";
            sqlcmd.Parameters.Add("@Sl_ct", SqlDbType.Int).Value = (object)1;
            DataSet source = StartupBase.SasObj.ExcuteReader(sqlcmd);
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel 2003|*.xls|Excel 2007|*.xlsx";
            saveFileDialog.Title = "Save a file";
            if (!saveFileDialog.ValidateNames)
                saveFileDialog.FileName = "";
            bool? nullable = saveFileDialog.ShowDialog();
            if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) == 0 || string.IsNullOrEmpty(saveFileDialog.FileName))
                return;
            VoucherContextMenuMouseRightClick.exportToExcel(source, saveFileDialog.FileName);
            if (!File.Exists(saveFileDialog.FileName))
                return;
            Process.Start(saveFileDialog.FileName);
        }

        public static void exportToExcel(DataSet source, string fileName)
        {
            StreamWriter streamWriter = new StreamWriter(fileName);
            int num1 = 0;
            streamWriter.Write("<xml version>\r\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-    microsoft-com:office:excel\"\r\n xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">\r\n <Styles>\r\n <Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n <Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>\r\n <Font/>\r\n <Interior/>\r\n <NumberFormat/>\r\n <Protection/>\r\n </Style>\r\n <Style ss:ID=\"BoldColumn\">\r\n <Font x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n </Style>\r\n <Style     ss:ID=\"StringLiteral\">\r\n <NumberFormat ss:Format=\"@\"/>\r\n </Style>\r\n <Style ss:ID=\"Decimal\">\r\n <NumberFormat ss:Format=\"0.0000\"/>\r\n </Style>\r\n <Style ss:ID=\"Integer\">\r\n <NumberFormat ss:Format=\"0\"/>\r\n </Style>\r\n <Style ss:ID=\"DateLiteral\">\r\n <NumberFormat ss:Format=\"mm/dd/yyyy;@\"/>\r\n </Style>\r\n </Styles>\r\n ");
            bool flag = false;
            int num2 = 0;
            int num3 = 0;
            while (!flag)
            {
                ++num2;
                ++num3;
                streamWriter.Write("<Worksheet ss:Name=\"Sheet" + (object)num2 + "\">");
                streamWriter.Write("<Table>");
                streamWriter.Write("<Row>");
                for (int index = 0; index < source.Tables[num3 - 1].Columns.Count; ++index)
                {
                    streamWriter.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
                    streamWriter.Write(source.Tables[num3 - 1].Columns[index].ColumnName);
                    streamWriter.Write("</Data></Cell>");
                }
                streamWriter.Write("</Row>");
                foreach (DataRow row in (InternalDataCollectionBase)source.Tables[num3 - 1].Rows)
                {
                    ++num1;
                    if (num1 == 64000)
                    {
                        num1 = 0;
                        ++num2;
                        streamWriter.Write("</Table>");
                        streamWriter.Write(" </Worksheet>");
                        streamWriter.Write("<Worksheet ss:Name=\"Sheet" + (object)num2 + "\">");
                        streamWriter.Write("<Table>");
                    }
                    streamWriter.Write("<Row>");
                    for (int index = 0; index < source.Tables[num3 - 1].Columns.Count; ++index)
                    {
                        Type type = row[index].GetType();
                        switch (type.ToString())
                        {
                            case "System.String":
                                string str1 = row[index].ToString().Trim().Replace("&", "&").Replace(">", ">").Replace("<", "<");
                                streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                streamWriter.Write(str1);
                                streamWriter.Write("</Data></Cell>");
                                break;
                            case "System.DateTime":
                                DateTime dateTime = (DateTime)row[index];
                                string str2 = dateTime.Year.ToString() + "-" + (dateTime.Month < 10 ? "0" + dateTime.Month.ToString() : dateTime.Month.ToString()) + "-" + (dateTime.Day < 10 ? "0" + dateTime.Day.ToString() : dateTime.Day.ToString()) + "T" + (dateTime.Hour < 10 ? "0" + dateTime.Hour.ToString() : dateTime.Hour.ToString()) + ":" + (dateTime.Minute < 10 ? "0" + dateTime.Minute.ToString() : dateTime.Minute.ToString()) + ":" + (dateTime.Second < 10 ? "0" + dateTime.Second.ToString() : dateTime.Second.ToString()) + ".000";
                                streamWriter.Write("<Cell ss:StyleID=\"DateLiteral\"><Data ss:Type=\"DateTime\">");
                                streamWriter.Write(str2);
                                streamWriter.Write("</Data></Cell>");
                                break;
                            case "System.Boolean":
                                streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                streamWriter.Write(row[index].ToString());
                                streamWriter.Write("</Data></Cell>");
                                break;
                            case "System.Int16":
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                streamWriter.Write("<Cell ss:StyleID=\"Integer\"><Data ss:Type=\"Number\">");
                                streamWriter.Write(row[index].ToString());
                                streamWriter.Write("</Data></Cell>");
                                break;
                            case "System.Decimal":
                            case "System.Double":
                                streamWriter.Write("<Cell ss:StyleID=\"Decimal\"><Data ss:Type=\"Number\">");
                                streamWriter.Write(row[index].ToString());
                                streamWriter.Write("</Data></Cell>");
                                break;
                            case "System.DBNull":
                                streamWriter.Write("<Cell ss:StyleID=\"StringLiteral\"><Data ss:Type=\"String\">");
                                streamWriter.Write("");
                                streamWriter.Write("</Data></Cell>");
                                break;
                            default:
                                throw new Exception(type.ToString() + " not handled.");
                        }
                    }
                    streamWriter.Write("</Row>");
                }
                streamWriter.Write("</Table>");
                streamWriter.Write(" </Worksheet>");
                if (num3 == source.Tables.Count)
                    flag = true;
            }
            streamWriter.Write("</Workbook>");
            streamWriter.Close();
        }

    }
}
