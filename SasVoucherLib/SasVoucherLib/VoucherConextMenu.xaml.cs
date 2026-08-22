using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.DataPresenter.Events;
using SasControls;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using SasDefine;
using SasDataLib;
using System.Windows.Threading;
using Microsoft.Win32;
using System.IO;
using System.Windows.Media;
using System.Text;

namespace SasVoucherLib
{
    /// <summary>Interaction logic for VoucherConextMenu.xaml</summary>
    /// <summary>VoucherConextMenu</summary>
    public partial class VoucherConextMenu : ContextMenu
    {
        public static readonly DependencyProperty SasObjProperty = DependencyProperty.Register(nameof(SasObj), typeof(SasObject), typeof(VoucherConextMenu), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        private static readonly string[] pnmua_list = new string[4]
        {
      "PNA",
      "PNB",
      "PNC",
      "PNG"
        };
        private Dictionary<string, string> _fieldList;

        public VoucherConextMenu()
        {
            this.InitializeComponent();
        }

        public SasObject SasObj
        {
            get
            {
                return (SasObject)this.GetValue(VoucherConextMenu.SasObjProperty);
            }
            set
            {
                this.SetValue(VoucherConextMenu.SasObjProperty, (object)value);
            }
        }

        public string Stt_rec
        {
            get
            {
                return this.VoucherRow == null ? "" : this.VoucherRow["stt_rec"].ToString();
            }
        }

        public DataRow VoucherRow { get; set; }

        public DataRow DmctRow { get; set; }

        private void Verify(DataTable data, string fieldList)
        {
            if (data == null || data.Columns.Count == 0)
                return;
            switch (fieldList)
            {
                case "":
                    break;
                case null:
                    break;
                default:
                    DataColumnCollection columns = data.Columns;
                    foreach (string str in fieldList.Split(";".ToCharArray()))
                    {
                        string[] strArray = str.Split(":".ToCharArray());
                        columns.Contains(strArray[0]);
                    }
                    break;
            }
        }

        private void ShowUpdateInfo(DataRow row)
        {
            if (row == null)
                return;
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("stt", typeof(int));
            dataTable.Columns.Add("info");
            dataTable.Columns.Add("descript");
            SqlCommand sqlcmd = new SqlCommand("Select * from userinfo where user_id = @user_id;Select * from userinfo where user_id = @user_id0");
            sqlcmd.Parameters.Add("@user_id", SqlDbType.VarChar).Value = row["user_id"];
            sqlcmd.Parameters.Add("@user_id0", SqlDbType.VarChar).Value = row["user_id0"];
            DataSet dataSet = this.SasObj.ExcuteReader(sqlcmd);
            object obj1 = row["date"];
            object obj2 = row["time"];
            object obj3 = row["user_id"];
            object obj4 = (object)"";
            object obj5 = (object)"";
            if (dataSet.Tables[0].Rows.Count > 0)
            {
                obj4 = dataSet.Tables[0].Rows[0]["user_name"];
                obj5 = dataSet.Tables[0].Rows[0]["comment"];
            }
            object obj6 = row["date0"];
            object obj7 = row["time0"];
            object obj8 = row["user_id0"];
            object obj9 = (object)"";
            object obj10 = (object)"";
            if (dataSet.Tables[1].Rows.Count > 0)
            {
                obj9 = dataSet.Tables[1].Rows[0]["user_name"];
                obj10 = dataSet.Tables[1].Rows[0]["comment"];
            }
            object obj11 = this.SasObj.UserInfo.Rows[0]["user_name"];
            object obj12 = this.SasObj.UserInfo.Rows[0]["comment"];
            if (this.SasObj != null)
            {
                if (this.SasObj.GetOption("M_LAN").Equals((object)"V"))
                {
                    dataTable.Rows.Add((object)1, (object)"Người sử dụng hiện thời", (object)obj11.ToString());
                    dataTable.Rows.Add((object)2, (object)"Tên người sử dụng hiện thời", (object)obj12.ToString());
                    dataTable.Rows.Add((object)3, (object)"Người tạo", (object)obj9.ToString());
                    dataTable.Rows.Add((object)4, (object)"Tên người tạo", (object)obj10.ToString());
                    dataTable.Rows.Add((object)5, (object)"Ngày tạo", obj6 is DBNull ? (object)"" : (object)((DateTime)obj6).ToString("dd-MM-yyyy"));
                    dataTable.Rows.Add((object)6, (object)"Thời gian tạo", obj7);
                    dataTable.Rows.Add((object)7, (object)"Người sửa", (object)obj4.ToString());
                    dataTable.Rows.Add((object)8, (object)"Tên người sửa", (object)obj5.ToString());
                    dataTable.Rows.Add((object)9, (object)"Ngày sửa", obj1 is DBNull ? (object)"" : (object)((DateTime)obj1).ToString("dd-MM-yyyy"));
                    dataTable.Rows.Add((object)10, (object)"Thời gian sửa", obj2);
                    dataTable.Rows.Add((object)11, (object)"Mã đơn vị cơ sở", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"]);
                }
                else
                {
                    dataTable.Rows.Add((object)1, (object)"Current  user", (object)obj11.ToString());
                    dataTable.Rows.Add((object)2, (object)"Current user's name", (object)obj12.ToString());
                    dataTable.Rows.Add((object)3, (object)"Created by", (object)obj9.ToString());
                    dataTable.Rows.Add((object)4, (object)"Name", (object)obj10.ToString());
                    dataTable.Rows.Add((object)5, (object)"Created on", obj6 is DBNull ? (object)"" : (object)((DateTime)obj6).ToString("dd-MM-yyyy"));
                    dataTable.Rows.Add((object)6, (object)"Created at", obj7);
                    dataTable.Rows.Add((object)7, (object)"Edited by", (object)obj4.ToString());
                    dataTable.Rows.Add((object)8, (object)"Name", (object)obj5.ToString());
                    dataTable.Rows.Add((object)9, (object)"Edited on", obj1 is DBNull ? (object)"" : (object)((DateTime)obj1).ToString("dd-MM-yyyy"));
                    dataTable.Rows.Add((object)10, (object)"Edited at", obj2);
                    dataTable.Rows.Add((object)11, (object)"Unit code", StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"]);
                }
            }
            string strBrowse = "stt:H=Stt:45:HR;info:H=Thông tin:250;descript:H=Nội dung:250";
            SasFormBrowes.FormBrowse formBrowse = new SasFormBrowes.FormBrowse(this.SasObj, dataTable.DefaultView, strBrowse);
            formBrowse.frmBrw.Title = SysFunc.Cat_Dau("Thông tin");
            formBrowse.frmBrw.LanguageID = "SasVoucherLib_2";
            formBrowse.ShowDialog();
        }

        private void ShowPrintInfo(DataRow row)
        {
            if (row == null)
                return;
            SqlCommand sqlcmd = new SqlCommand("SELECT a.sl_in,b.[user_name],a.[date],a.[time],a.[hostname] FROM phuserininfo a,userinfo b WHERE stt_rec = @stt_rec AND a.[user_id] = b.[user_id]ORDER BY a.[date],a.[time]");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.VarChar).Value = row["stt_rec"];
            SasFormBrowes.FormBrowse formBrowse = new SasFormBrowes.FormBrowse(this.SasObj, this.SasObj.ExcuteReader(sqlcmd).Tables[0].DefaultView, "sl_in:H=Lần in:55:F=0;user_name:H=Người in:250;date:105:h=Ngày in:D;time:100:h=Giờ in;hostname:H=Tên máy in:250");
            formBrowse.frmBrw.Title = SysFunc.Cat_Dau("Thông tin in");
            formBrowse.frmBrw.LanguageID = "SasVoucherLib_3";
            formBrowse.ShowDialog();
        }

        private void ShowReport(string menu)
        {
            if (this.SasObj == null)
                return;
            DataRow commandInfo = SysFunc.GetCommandInfo(this.SasObj, menu);
            SqlCommand cmd = new SqlCommand(commandInfo["store_proc"].ToString());
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = (object)this.Stt_rec;
            DataSet ds = this.SasObj.ExcuteReader(cmd);
            if (ds == null || ds.Tables.Count == 0)
                return;
            string strBrowse = !(StartUpTrans.M_LAN == "V") ? commandInfo["Ebrowse1"].ToString().Trim() : commandInfo["vbrowse1"].ToString().Trim();
            SasFormBrowes.FormBrowse oBrowse = new SasFormBrowes.FormBrowse(this.SasObj, ds.Tables[0].DefaultView, strBrowse);
            if (StartUpTrans.M_LAN == "V")
                oBrowse.frmBrw.Title = SysFunc.Cat_Dau(commandInfo["title"].ToString());
            else
                oBrowse.frmBrw.Title = commandInfo["title"].ToString();
            if (menu == "01.50.25")
                oBrowse.DataGrid.FieldLayoutInitialized += new EventHandler<FieldLayoutInitializedEventArgs>(this.ctgt20_FieldLayoutInitialized);
            if (menu == "01.50.20")
                oBrowse.DataGrid.FieldLayoutInitialized += new EventHandler<FieldLayoutInitializedEventArgs>(this.ctgt30_FieldLayoutInitialized);
            oBrowse.AllowModifyVoucher = false;
            oBrowse.frmBrw.LanguageID = commandInfo["parameter"].ToString();
            oBrowse.ShowDialog();
            oBrowse.CTRL_R += (SasFormBrowes.FormBrowse.GridKeyUp_CTRL_R)((s, e) =>
           {
               ds = this.SasObj.ExcuteReader(cmd);
               if (ds.Tables.Count <= 0)
                   return;
               oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)ds.Tables[0].DefaultView;
               oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
               oBrowse.UpdateSumaryFields();
               GC.Collect();
               int num = (int)GC.WaitForFullGCComplete();
           });
        }

        private void ctgt20_FieldLayoutInitialized(object sender, FieldLayoutInitializedEventArgs e)
        {
            string str = this.SasObj.GetOption("M_DC_THUE_CK").ToString();
            XamDataGrid xamDataGrid = sender as XamDataGrid;
            if (!(str == "2") && !(str == "3") && xamDataGrid.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ten_vt")))
                xamDataGrid.FieldLayouts[0].Fields["ten_vt"].Visibility = Visibility.Collapsed;
            if (str == "1" || str == "3" || !xamDataGrid.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "dia_chi")))
                return;
            xamDataGrid.FieldLayouts[0].Fields["dia_chi"].Visibility = Visibility.Collapsed;
        }

        private void ctgt30_FieldLayoutInitialized(object sender, FieldLayoutInitializedEventArgs e)
        {
            string str = this.SasObj.GetOption("M_DC_THUE_CK").ToString();
            XamDataGrid xamDataGrid = sender as XamDataGrid;
            if (!(str == "2") && !(str == "3") && xamDataGrid.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "ten_vt")))
                xamDataGrid.FieldLayouts[0].Fields["ten_vt"].Visibility = Visibility.Collapsed;
            if (str == "1" || str == "3" || !xamDataGrid.FieldLayouts[0].Fields.Any<Field>((Func<Field, bool>)(x => x.Name == "dia_chi")))
                return;
            xamDataGrid.FieldLayouts[0].Fields["dia_chi"].Visibility = Visibility.Collapsed;
        }

        private void mnuInfo_Click(object sender, RoutedEventArgs e)
        {
            this.ShowUpdateInfo(this.VoucherRow);
        }

        private void mnuIn_Click(object sender, RoutedEventArgs e)
        {
            this.ShowPrintInfo(this.VoucherRow);
        }

        private void mnuDel_Click(object sender, RoutedEventArgs e)
        {
            VoucherContextMenuDel voucherContextMenuDel = new VoucherContextMenuDel(this.SasObj);
            SysFunc.LoadIcon((Window)voucherContextMenuDel);
            voucherContextMenuDel.Title = SysFunc.Cat_Dau("Lọc các chứng từ bị xóa");
            voucherContextMenuDel.ShowDialog();
            if (!voucherContextMenuDel.bIsOK)
                return;
            string str1 = "1=1";
            string str2 = "";
            if (!string.IsNullOrEmpty(voucherContextMenuDel.txtSo_ct1.Text))
            {
                int result;
                if (!int.TryParse(voucherContextMenuDel.txtSo_ct1.Text, out result))
                    str1 = str1 + " and so_ct >= '" + voucherContextMenuDel.txtSo_ct1.Text + "' and len(so_ct) >= " + (object)voucherContextMenuDel.txtSo_ct1.Text.Trim().Length;
                else
                    str1 = str1 + "and IsNumeric(so_ct)=1 and  so_ct >= " + (object)result;
            }
            if (!string.IsNullOrEmpty(voucherContextMenuDel.txtSo_ct2.Text))
            {
                int result;
                if (!int.TryParse(voucherContextMenuDel.txtSo_ct2.Text, out result))
                    str1 = str1 + " and so_ct <= '" + voucherContextMenuDel.txtSo_ct2.Text + "'and len(so_ct) <= " + (object)voucherContextMenuDel.txtSo_ct2.Text.Trim().Length;
                else
                    str1 = str1 + "and IsNumeric(so_ct)=1 and  so_ct <= " + (object)result;
            }
            if (!string.IsNullOrEmpty(voucherContextMenuDel.txtMaDVCS.Text.Trim()))
                str2 = " and ma_dvcs LIKE '" + voucherContextMenuDel.txtMaDVCS.Text.Trim() + "'";
            string format = "{0}";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 10 ? string.Format(format, (object)"GetDeletedVoucher") : string.Format(format, (object)StartUpTrans.Process_Store[10]));
            sqlcmd.CommandType = CommandType.StoredProcedure;
            sqlcmd.Parameters.Add("@Ma_ct", SqlDbType.VarChar).Value = (object)StartUpTrans.Ma_ct;
            sqlcmd.Parameters.Add("@Ngay_ct1", SqlDbType.VarChar).Value = (object)string.Format("{0:yyyyMMdd}", (object)voucherContextMenuDel.txtNgay_ct1.dValue);
            sqlcmd.Parameters.Add("@Ngay_ct2", SqlDbType.VarChar).Value = (object)string.Format("{0:yyyyMMdd}", (object)voucherContextMenuDel.txtNgay_ct2.dValue);
            sqlcmd.Parameters.Add("@Advance", SqlDbType.NVarChar).Value = (object)str1;
            sqlcmd.Parameters.Add("@Ma_dvcs", SqlDbType.NVarChar).Value = (object)str2;
            DataSet dataSet = this.SasObj.ExcuteReader(sqlcmd);
            if (dataSet == null || dataSet.Tables.Count == 0)
                return;
            DataRow commandInfo1 = SysFunc.GetCommandInfo(this.SasObj, StartupBase.Menu_Id);
            DataRow commandInfo2 = SysFunc.GetCommandInfo(this.SasObj, commandInfo1["menu_copy"].ToString().Trim());
            if (commandInfo2 != null)
            {
                commandInfo1["vbrowse2"] = commandInfo2["vbrowse2"];
                commandInfo1["ebrowse2"] = commandInfo2["ebrowse2"];
            }
            string[] strArray = (string[])null;
            string str3 = "";
            string strBrowseCt = "";
            if (StartupBase.M_LAN.Equals("V"))
            {
                if (commandInfo1["Vbrowse2"].ToString().Trim() != "")
                    strArray = commandInfo1["Vbrowse2"].ToString().Split('|');
            }
            else if (commandInfo1["Ebrowse2"].ToString().Trim() != "")
                strArray = commandInfo1["Ebrowse2"].ToString().Split('|');
            if (strArray != null)
            {
                str3 = strArray[0];
                strBrowseCt = strArray[1];
            }
            FormView formView = new FormView(this.SasObj, dataSet.Tables[0].DefaultView, dataSet.Tables[1].DefaultView, str3, strBrowseCt, "stt_rec");
            if (((IEnumerable<string>)SysFunc.GetFieldFromStrBrowse(str3).Split(',')).Contains<string>("t_tt_nt"))
            {
                formView.ListFieldSum = "t_tt_nt;t_tt";
                formView.TongCongLabel = "Tổng tiền";
            }
            else
                formView.ListFieldSum = "t_tien_nt;t_tien";
            if (StartupBase.M_LAN.Equals("V"))
                formView.frmBrw.Title = SysFunc.Cat_Dau(commandInfo1["bar"].ToString().Trim()) + " - các chứng từ bị xóa";
            else
                formView.frmBrw.Title = SysFunc.Cat_Dau(commandInfo1["bar2"].ToString().Trim()) + " - vouchers deleted";
            FreeCodeFieldLib.InitFreeCodeField(this.SasObj, formView.frmBrw.oBrowseCt, StartUpTrans.Ma_ct, 1);
            formView.frmBrw.LanguageID = "SasVoucherLib_1";
            formView.ShowDialog();
        }

        private string ConvertDataToSql(object value, Type ValueType)
        {
            string str;
            switch (ValueType.ToString())
            {
                case "System.String":
                    str = string.Format("'{0}'", (object)(value as string).Replace("'", "'"));
                    break;
                case "System.DateTime":
                    str = string.Format("'{0}'", (object)((DateTime)value).ToString("yyyyMMdd"));
                    break;
                default:
                    str = string.Format("'{0}'", value);
                    break;
            }
            return str;
        }

        private void mnuReport_Click(object sender, RoutedEventArgs e)
        {
            this.ShowReport((sender as MenuItem).Tag.ToString());
        }

        public DataSet GetPhieuht(string stt_rec)
        {
            SqlCommand sqlcmd = new SqlCommand("InPhieuht");
            sqlcmd.CommandType = CommandType.StoredProcedure;
            sqlcmd.Parameters.Add("@Stt_rec", SqlDbType.VarChar).Value = (object)stt_rec;
            return this.SasObj.ExcuteReader(sqlcmd);
        }

        private void mnuVcIn_Click(object sender, RoutedEventArgs e)
        {
            if (this.SasObj == null)
                return;
            DataSet ds = this.GetPhieuht(this.Stt_rec);
            if (ds == null || ds.Tables.Count == 0)
                return;
            if (ds.Tables[0].Rows.Count == 0)
            {
                DataRow row = ds.Tables[0].NewRow();
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("so_ct"))
                    row["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_kh"))
                    row["ma_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("dien_giai"))
                    row["dien_giai"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dien_giai"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ngay_lct"))
                    row["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ngay_ct"))
                    row["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ten_kh"))
                    row["ten_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"];
                if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                    row["ma_nt"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"];
                row["ps_no"] = (object)0;
                row["ps_co"] = (object)0;
                row["ps_no_nt"] = (object)0;
                row["ps_co_nt"] = (object)0;
                ds.Tables[0].Rows.Add(row);
            }
            string ma_nt = ds.Tables[0].Rows[0]["ma_nt"].ToString();
            this.GetDmnt(ds, ma_nt);
            ReportManager oReport = new ReportManager(this.SasObj, "Vc-In");
            StackPanel name1 = oReport._frm.FindName("stck") as StackPanel;
            Button button = new Button();
            button.Width = 92.0;
            button.Height = 25.0;
            button.Margin = new Thickness(12.0, 0.0, 0.0, 0.0);
            button.Name = "btnPrintAll";
            button.Click += (RoutedEventHandler)((s, a) =>
           {
               if (this.SasObj.GetOption("M_IN_HOI_CK").ToString() == "1" && ExMessageBox.Show(-11, this.SasObj, "Có chắc chắn in tất cả các chứng từ đã được lọc?", "", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
               {
                   ControlFilterReport name2 = oReport._frm.FindName("CDReport") as ControlFilterReport;
                   for (int index1 = 1; index1 < StartUpTrans.DsTrans.Tables[0].Rows.Count; ++index1)
                   {
                       DataSet phieuht = this.GetPhieuht(StartUpTrans.DsTrans.Tables[0].Rows[index1]["stt_rec"].ToString());
                       ds.Tables[0].Rows.Clear();
                       for (int index2 = 0; index2 < phieuht.Tables[0].Rows.Count; ++index2)
                           ds.Tables[0].Rows.Add(phieuht.Tables[0].Rows[index2].ItemArray);
                       if (ds == null || ds.Tables.Count == 0)
                           return;
                       if (ds.Tables[0].Rows.Count == 0)
                       {
                           DataRow row = ds.Tables[0].NewRow();
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("so_ct"))
                               row["so_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["so_ct"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_kh"))
                               row["ma_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("dien_giai"))
                               row["dien_giai"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["dien_giai"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ngay_lct"))
                               row["ngay_lct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_lct"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ngay_ct"))
                               row["ngay_ct"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ngay_ct"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ten_kh"))
                               row["ten_kh"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ten_kh"];
                           if (StartUpTrans.DsTrans.Tables[0].Columns.Contains("ma_nt"))
                               row["ma_nt"] = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"];
                           row["ps_no"] = (object)0;
                           row["ps_co"] = (object)0;
                           row["ps_no_nt"] = (object)0;
                           row["ps_co_nt"] = (object)0;
                           ds.Tables[0].Rows.Add(row);
                       }
                       ds.Tables.RemoveAt(1);
                       ma_nt = ds.Tables[0].Rows[0]["ma_nt"].ToString();
                       this.GetDmnt(ds, ma_nt);
                       name2.V_In((short)1);
                   }
                   oReport._frm.Close();
               }
               oReport._frm.Close();
           });
            button.Content = (object)"In liên tục";
            name1.Children.Insert(1, (UIElement)button);
            oReport.Preview(ds);
        }

        public DataTable GetDmnt(DataSet ds, string ma_nt)
        {
            string str = this.SasObj.GetOption("M_MA_NT0").ToString();
            SqlCommand sqlcmd = new SqlCommand("Select *,@ma_nt0 as ma_nt0, @Type as read_num_type from dmnt");
            sqlcmd.Parameters.Add("@ma_nt0", SqlDbType.Char, 3).Value = (object)str;
            sqlcmd.Parameters.Add("@Type", SqlDbType.Char, 1).Value = this.SasObj.GetOption("M_READ_NUM");
            DataTable table = this.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
            table.TableName = "TableNTInfo";
            if (ds.Tables.IndexOf("TableNTInfo") >= 0)
                ds.Tables.Remove("TableNTInfo");
            ds.Tables.Add(table);
            return table;
        }

        private void ContextMenu_Loaded(object sender, RoutedEventArgs e)
        {
            if (((IEnumerable<string>)VoucherConextMenu.pnmua_list).Any<string>((Func<string, bool>)(x => x == StartUpTrans.Ma_ct)))
                return;
            this.mnuCt70_PNA.Visibility = Visibility.Collapsed;
        }


        private void mnuCopyData_Click(object sender, RoutedEventArgs e)
        {
            this.CopyDataToVAT(this.VoucherRow);
        }

        private void CopyDataToVAT(DataRow row)
        {
            if (row == null)
                return;
            try
            {
                string stt_rec = row["stt_rec"].ToString().Trim();
                string ws_id = StartupBase.SasObj.GetOption("M_WS_ID").ToString().Trim();
                using (SqlCommand sqlcmd = new SqlCommand("Select datasource, initialcatalog, userid, password from dmconsql where ma_data = 'M_SIS_SQL_VAT'"))
                {
                    DataSet dataSet = this.SasObj.ExcuteReader(sqlcmd);
                    if (dataSet.Tables[0].Rows.Count > 0)
                    {
                        string server = dataSet.Tables[0].Rows[0]["datasource"].ToString().Trim();
                        string data = dataSet.Tables[0].Rows[0]["initialcatalog"].ToString().Trim();
                        string user = dataSet.Tables[0].Rows[0]["userid"].ToString().Trim();
                        string pass = dataSet.Tables[0].Rows[0]["password"].ToString().Trim();
                        if (string.IsNullOrEmpty(server) || string.IsNullOrEmpty(data) || string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
                        {
                            MessageBox.Show("Thông tin kết nối dữ liệu VAT không đúng. Hãy kiểm tra lại.?", "Thông báo");
                            return;
                        }
                        SisSQLClient.SisSQL cls = new SisSQLClient.SisSQL(server, data, user, pass);
                        if (!cls.CheckConnect())
                        {
                            MessageBox.Show("Thông tin kết nối dữ liệu VAT không đúng. Hãy kiểm tra lại.?", "Thông báo");
                            return;
                        }
                        else
                        {
                            string ws_id_vat = cls.ExcuteScalar(new SqlCommand("select top 1 value from options where name = 'M_WS_ID'")).ToString().Trim();
                            if (string.IsNullOrEmpty(ws_id_vat) || (ws_id.Trim() == ws_id_vat.Trim()))
                            {
                                MessageBox.Show("Work station ID giống nhau. Hãy kiểm tra lại.?", "Thông báo");
                                return;
                            }
                            string m_phdbf = StartUpTrans.DmctInfo["m_phdbf"].ToString().Trim();
                            string m_ctdbf = StartUpTrans.DmctInfo["m_ctdbf"].ToString().Trim();
                            string m_ctgtdbf = StartUpTrans.DmctInfo["m_ctgtdbf"].ToString().Trim();
                            string m_ma_ct = StartUpTrans.DmctInfo["ma_ct"].ToString().Trim();

                            SasFormBrowes.FrmWaiting frm = new SasFormBrowes.FrmWaiting(70);
                            frm.Set(5);
                            try
                            {
                                //Thuc hien copy du lieu
                                //Bang PH
                                if (!string.IsNullOrEmpty(m_phdbf))
                                {
                                    DataTable dtph = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", m_phdbf, stt_rec))).Tables[0];
                                    if (dtph != null && dtph.Rows.Count > 0)
                                    {
                                        cls.InsertTable(m_phdbf, dtph, stt_rec, "row_id");
                                    }
                                }
                                frm.Show();
                                frm.Set(10);

                                //Bang CT
                                if (!string.IsNullOrEmpty(m_ctdbf))
                                {
                                    DataTable dtct = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", m_ctdbf, stt_rec))).Tables[0];
                                    if (dtct != null && dtct.Rows.Count > 0)
                                    {
                                        cls.InsertTable(m_ctdbf, dtct, stt_rec, "row_id");
                                    }
                                }
                                frm.Set(15);
                                //Bang CTGT
                                if (!string.IsNullOrEmpty(m_ctgtdbf))
                                {
                                    DataTable dtctgt = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", m_ctgtdbf, stt_rec))).Tables[0];
                                    if (dtctgt != null && dtctgt.Rows.Count > 0)
                                    {
                                        cls.InsertTable(m_ctgtdbf, dtctgt, stt_rec, "row_id");
                                    }
                                }
                                frm.Set(20);
                                //CT00,CT70,CTGT30,CTGT20,CTTT30,CTTT20,CTHHD,CT70HDM,CT70HD
                                //CT00
                                DataTable dtct00 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT00", stt_rec))).Tables[0];
                                if (dtct00 != null && dtct00.Rows.Count > 0)
                                {
                                    cls.InsertTable("CT00", dtct00, stt_rec, "row_id");
                                }
                                frm.Set(25);
                                //CT70
                                DataTable dtct70 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70", stt_rec))).Tables[0];
                                if (dtct70 != null && dtct70.Rows.Count > 0)
                                {
                                    cls.InsertTable("CT70", dtct70, stt_rec, "row_id");

                                    //1.MA_CT = PXE
                                    if (m_ma_ct.ToString().Trim() == "PXE")
                                    {
                                        string stt_rec_pn_pxe = dtct70.Rows[0]["stt_rec_pn"].ToString().Trim();
                                        DataTable dtct70PXE = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70", stt_rec_pn_pxe))).Tables[0];
                                        if (dtct70PXE != null && dtct70PXE.Rows.Count > 0)
                                        {
                                            cls.InsertTable("CT70", dtct70PXE, stt_rec_pn_pxe, "row_id");
                                        }
                                    }
                                    //2.MA_CT = PXV
                                    if (m_ma_ct.ToString().Trim() == "PXV")
                                    {
                                        string stt_rec_pn_pxv = dtct70.Rows[0]["stt_rec_pn"].ToString().Trim();
                                        DataTable dtct70PXV = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70", stt_rec_pn_pxv))).Tables[0];
                                        if (dtct70PXV != null && dtct70PXV.Rows.Count > 0)
                                        {
                                            cls.InsertTable("CT70", dtct70PXV, stt_rec_pn_pxv, "row_id");
                                        }
                                    }
                                    //2.MA_CT = PNG
                                    if (m_ma_ct.ToString().Trim() == "PNG")
                                    {
                                        string stt_rec_pn_pxg = stt_rec.Replace("PNG", "PXG").ToString().Trim();
                                        DataTable dtct70PNG = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70", stt_rec_pn_pxg))).Tables[0];
                                        if (dtct70PNG != null && dtct70PNG.Rows.Count > 0)
                                        {
                                            cls.InsertTable("CT70", dtct70PNG, stt_rec_pn_pxg, "row_id");
                                        }
                                    }

                                }
                                frm.Set(30);
                                //CTGT30
                                DataTable dtctgt30 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CTGT30", stt_rec))).Tables[0];
                                if (dtctgt30 != null && dtctgt30.Rows.Count > 0)
                                {
                                    cls.InsertTable("CTGT30", dtctgt30, stt_rec, "row_id");
                                }
                                frm.Set(35);
                                //CTGT20
                                DataTable dtctgt20 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CTGT20", stt_rec))).Tables[0];
                                if (dtctgt20 != null && dtctgt20.Rows.Count > 0)
                                {
                                    cls.InsertTable("CTGT20", dtctgt20, stt_rec, "row_id");
                                }
                                frm.Set(40);
                                //CTTT30
                                DataTable dtcttt30 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CTTT30", stt_rec))).Tables[0];
                                if (dtcttt30 != null && dtcttt30.Rows.Count > 0)
                                {
                                    cls.InsertTable("CTTT30", dtcttt30, stt_rec, "row_id");
                                }
                                frm.Set(45);
                                //CTTT20
                                DataTable dtcttt20 = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CTTT20", stt_rec))).Tables[0];
                                if (dtcttt20 != null && dtcttt20.Rows.Count > 0)
                                {

                                    bool flag = cls.InsertTable("CTTT20", dtcttt20, stt_rec, "row_id");
                                    SisSQLClient.SisErrorLog.AddLog(flag.ToString());
                                }
                                frm.Set(50);
                                //CTHHD
                                DataTable dtcthhd = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CTHHD", stt_rec))).Tables[0];
                                if (dtcthhd != null && dtcthhd.Rows.Count > 0)
                                {
                                    cls.InsertTable("CTHHD", dtcthhd, stt_rec, "row_id");
                                }
                                frm.Set(55);
                                //CT70HDM
                                DataTable dtct70hdm = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70HDM", stt_rec))).Tables[0];
                                if (dtct70hdm != null && dtct70hdm.Rows.Count > 0)
                                {
                                    cls.InsertTable("CT70HDM", dtct70hdm, stt_rec, "row_id");
                                }
                                frm.Set(60);
                                //CT70HD
                                DataTable dtct70hd = StartupBase.SasObj.ExcuteReader(new SqlCommand(string.Format("select * from {0} where stt_rec = '{1}'", "CT70HD", stt_rec))).Tables[0];
                                if (dtct70hd != null && dtct70hd.Rows.Count > 0)
                                {
                                    cls.InsertTable("CT70HD", dtct70hd, stt_rec, "row_id");
                                }

                                frm.Set(70);
                            }
                            catch (Exception ex)
                            {
                                SasErrorLib.ErrorLog.CatchMessage(ex);
                            }
                            finally
                            {
                                frm.Set(70);
                                frm.Hide();
                                frm.Close();
                                frm = null;
                            }
                        }
                    }
                    else
                    {
                        ExMessageBox.Show(1927, StartupBase.SasObj, "Thông tin kết nối dữ liệu VAT không đúng. Hãy kiểm tra lại!", "Configuration error!", MessageBoxButton.OK, MessageBoxImage.Asterisk);                        
                    }    
                }
            }
            catch (Exception ex)
            {
                SasErrorLib.ErrorLog.CatchMessage(ex);
            }
        }

        public static bool IsLoadFromLookUp = false;
        public static string TableName = "dmctdoc";
        public static DataTable dmctdocData;
        public static FormBrowse1 oBrowse;
        public static ActionTask currActionTask = ActionTask.None;
        public static string currSqlTableKeyValue = string.Empty;
        public static string currSqlTableKey = "file_id";
        public static string currStt_rec = string.Empty;
        public static string currMa_ct = string.Empty;
        public static DateTime currngay_ct;
        public static string SqlTableKey = "file_id";
        public static string phanquyenfilter = "";
        void LoadPhanquyen(string stt_rec)
        {
            bool admin_root = Convert.ToBoolean(StartupBase.SasObj.UserInfo.Rows[0]["is_admin"]);           
            if (admin_root)
                phanquyenfilter = " AND 1=1";
            else
            {
                string u_id = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString().Trim();
                string sql = "select user_right from " + TableName + " where stt_rec = '" + stt_rec + "'";
                DataTable tblPQ = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql)).Tables[0];
                if (tblPQ.Rows.Count != 0)
                {
                    string[] strUser = tblPQ.Rows[0]["user_right"].ToString().Trim().Split(';');
                    if (strUser.Length > 0)
                    {
                        phanquyenfilter += " AND isnull(user_right,'') ='' AND (user_id2 = '" + u_id + "'  OR '" + u_id + "' IN (";
                        foreach (string str in strUser)
                        {
                            phanquyenfilter += "'" + str + "',";
                        }
                        phanquyenfilter = phanquyenfilter.Substring(0, phanquyenfilter.Length - 1);
                        phanquyenfilter += "))";
                    }
                }
            } 
        }
        public static SqlCommand MCmd;
        private void ShowAttchFile(DataRow row)
        {
            //LoadPhanquyen(row["stt_rec"].ToString().Trim());
            bool admin_root = Convert.ToBoolean(StartupBase.SasObj.UserInfo.Rows[0]["is_admin"]);
            string u_id = StartupBase.SasObj.UserInfo.Rows[0]["user_id"].ToString().Trim();
            phanquyenfilter = "";
            if (admin_root)
                phanquyenfilter = " AND 1=1";
            else
            {
                phanquyenfilter = " AND CHARINDEX('" + u_id + ";',RTRIM(user_right)+';') > 0";
            }
            currStt_rec = row["stt_rec"].ToString().Trim();
            currMa_ct = row["ma_ct"].ToString().Trim();
            currngay_ct = Convert.ToDateTime(row["ngay_ct"]);
            string sql = "Select * from " + TableName + " where stt_rec = '" + row["stt_rec"].ToString().Trim() + "'" + phanquyenfilter;
            MCmd = new SqlCommand(sql);
            dmctdocData = StartupBase.SasObj.ExcuteReader(MCmd).Tables[0];
            string strVN = "stt:100:h=Stt;ten_hs:150:h=Tên HS;ten_hs2:150:h=Tên HS 2;ten_file:250:h=Tên file;file_id:0:h=File id";
            string strEN = "stt:100:h=No;ten_hs:150:h=Name;ten_hs2:150:h=Name 2;ten_file:250:h=File name;file_id:0:h=File id";
            string strBrowse = StartupBase.M_LAN.Equals("V") ? strVN : strEN;
            oBrowse = new SasFormBrowes.FormBrowse1(StartupBase.SasObj, dmctdocData.DefaultView, strBrowse);
            oBrowse.frmBrw.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? "File đính kèm F2-Xem, F3-Sửa, F4-Thêm,  F8-Xóa." : "Attach file F2-View, F3-Edit, F4-New, F8-Delete");
            object name = oBrowse.frmBrw.ToolBar.FindName("tbList");
            if (name != null)
            {
                ToolBar toolBar = name as ToolBar;
                ToolBarButton toolBarButton = new ToolBarButton();
                toolBarButton.BorderBrush = (Brush)Brushes.Transparent;
                toolBarButton.Name = "btnF5";
                toolBarButton.Text = (StartupBase.M_LAN.Equals("V") ? "Tải file" : "Download");
                toolBarButton.ToolTip = (object)"F5";
                toolBarButton.ImagePath = "Images\\Preview.png";
                toolBarButton.Click += new RoutedEventHandler(oBrowse_F5);
                toolBar.Items.Insert(3, (object)toolBarButton);
            }
            oBrowse.F2 += new SasFormBrowes.FormBrowse1.GridKeyUp_F2(this.oBrowse_F2);
            oBrowse.F3 += new SasFormBrowes.FormBrowse1.GridKeyUp_F3(this.oBrowse_F3);
            oBrowse.F4 += new SasFormBrowes.FormBrowse1.GridKeyUp_F4(this.oBrowse_F4);
            oBrowse.F5 += new SasFormBrowes.FormBrowse1.GridKeyUp_F5(this.oBrowse_F5);
            oBrowse.F8 += new SasFormBrowes.FormBrowse1.GridKeyUp_F8(this.oBrowse_F8);

            oBrowse.frmBrw.ToolBar.Cm_Xem += ToolBar_Cm_Xem;
            oBrowse.frmBrw.ToolBar.Cm_Moi += ToolBar_Cm_Moi;
            oBrowse.frmBrw.ToolBar.Cm_Sua += ToolBar_Cm_Sua;
            oBrowse.frmBrw.ToolBar.Cm_Xoa += ToolBar_Cm_Xoa;
            oBrowse.frmBrw.LanguageID = "Sasvoucger_attachfile";
            oBrowse.ShowDialog();
        }

        private void ToolBar_Cm_Xem()
        {
            v_xem();
        }
        private void ToolBar_Cm_Xoa()
        {
            v_Xoa();
        }
        private void ToolBar_Cm_Sua()
        {
            v_Sua();
        }
        private void ToolBar_Cm_Moi()
        {
            v_Them();
        }
        private void oBrowse_F2(object sender, EventArgs e)
        {
            v_xem();
        }
        private void oBrowse_F3(object sender, EventArgs e)
        {
            v_Sua();
        }
        private void oBrowse_F4(object sender, EventArgs e)
        {
            v_Them();
        }
        private void oBrowse_F8(object sender, EventArgs e)
        {
            v_Xoa();
        }
        private void oBrowse_F5(object sender, EventArgs e)
        {
            SaveFile();
        }

        private void v_xem()
        {
            if (VoucherConextMenu.oBrowse.ActiveRecord == null || VoucherConextMenu.currActionTask != ActionTask.None)
                return;
            VoucherConextMenu.currActionTask = ActionTask.View;
            VoucherConextMenu.currSqlTableKeyValue = VoucherConextMenu.oBrowse.ActiveRecord.Cells[VoucherConextMenu.SqlTableKey].Value.ToString();
            this.showWindow();
        }
        private void v_Sua()
        {
            if (SysFunc.CheckPermission(StartupBase.SasObj, VoucherConextMenu.currActionTask, StartupBase.Menu_Id))
            {
                if (VoucherConextMenu.oBrowse.ActiveRecord == null || VoucherConextMenu.currActionTask != ActionTask.None)
                    return;
                VoucherConextMenu.currActionTask = ActionTask.Edit;
                VoucherConextMenu.currSqlTableKeyValue = VoucherConextMenu.oBrowse.ActiveRecord.Cells[VoucherConextMenu.SqlTableKey].Value.ToString();
                this.showWindow();
            }
        }
        private void v_Xoa()
        {
            DataRecord activeRecord = VoucherConextMenu.oBrowse.ActiveRecord;
            if (activeRecord == null || activeRecord.RecordType != RecordType.DataRecord || ExMessageBox.Show(1920, StartupBase.SasObj, "Có chắc chắn xóa không?", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) != MessageBoxResult.Yes || VoucherConextMenu.currActionTask != ActionTask.None)
                return;
            VoucherConextMenu.currSqlTableKeyValue = VoucherConextMenu.GetCurrKeyValueToString(activeRecord, VoucherConextMenu.SqlTableKey);
            VoucherConextMenu.currActionTask = ActionTask.Delete;
            if (SysFunc.CheckPermission(StartupBase.SasObj, VoucherConextMenu.currActionTask, StartupBase.Menu_Id))
            {
                VoucherConextMenu.dmctdocData = VoucherConextMenu.dmctdocData.DefaultView.ToTable();
                this.deleteRowByKey(VoucherConextMenu.TableName, VoucherConextMenu.SqlTableKey, VoucherConextMenu.currSqlTableKeyValue, VoucherConextMenu.dmctdocData, StartupBase.SasObj);
                VoucherConextMenu.currActionTask = ActionTask.None;
            }
            else
            {
                int num = (int)ExMessageBox.Show(1925, StartupBase.SasObj, "Không có quyền xóa [" + VoucherConextMenu.TableName + "]", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString(), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                VoucherConextMenu.currActionTask = ActionTask.None;
            }
        }
        private void deleteRowByKey(string tableName, string strKey, string key, DataTable dt, SasObject SasObj)
        {
            if (VoucherConextMenu.oBrowse.ActiveRecord == null)
                return;
            int index = VoucherConextMenu.oBrowse.ActiveRecord.Index;
            int dataItemIndex = VoucherConextMenu.oBrowse.ActiveRecord.DataItemIndex;
            int visibleIndex = VoucherConextMenu.oBrowse.ActiveRecord.VisibleIndex;
            if (ListFunc.deleteRowInDatabaseByKey(tableName, strKey, dt.Rows[dataItemIndex], SasObj) >= 1)
            {
                dt.Rows[dataItemIndex].Delete();
                dt.AcceptChanges();
                VoucherConextMenu.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)VoucherConextMenu.dmctdocData.DefaultView;
                VoucherConextMenu.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                {
                    if (VoucherConextMenu.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().Count<DataRecord>() > visibleIndex)
                        VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)VoucherConextMenu.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().ElementAt<DataRecord>(visibleIndex - 1);
                    else
                        VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)VoucherConextMenu.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
                    VoucherConextMenu.oBrowse.frmBrw.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
                    {
                        if (VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord == null || VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord.VisibleIndex != -1)
                            return;
                        VoucherConextMenu.oBrowse.frmBrw.oBrowse.ActiveRecord = (Record)VoucherConextMenu.oBrowse.frmBrw.oBrowse.RecordManager.GetFilteredInDataRecords().LastOrDefault<DataRecord>();
                    }));
                }));
            }
        }
        private void v_Them()
        {
            if (VoucherConextMenu.currActionTask != ActionTask.None)
                return;
            VoucherConextMenu.currActionTask = ActionTask.Add;
            if (SysFunc.CheckPermission(StartupBase.SasObj, VoucherConextMenu.currActionTask, StartupBase.Menu_Id))
            {                
                this.showWindow();
            }
        }        
        private void mnuFile_Click(object sender, RoutedEventArgs e)
        {
            ShowAttchFile(this.VoucherRow);
        }

        public static DataTable GetRow(string TableName)
        {
            DataTable dataTable = (DataTable)null;
            SqlCommand sqlcmd = new SqlCommand();
            try
            {
                string str = "select * from " + TableName + " where " + SqlTableKey + " ='" + currSqlTableKeyValue + "'";
                sqlcmd.CommandText = str;
                dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            }
            catch (SqlException ex)
            {

            }
            return dataTable;
        }
        public static string GetCurrKeyValueToString(DataRecord r, string SqlTableKey)
        {
            if (r == null || r.DataItem == null)
                return string.Empty;
            string[] strArray = SqlTableKey.Split(';');
            string str1 = string.Empty;
            for (int index = 0; index < strArray.Length; ++index)
            {
                string str2 = strArray[index];
                if ((r.DataItem as DataRowView)[str2] != DBNull.Value)
                    str1 = str1 + (r.DataItem as DataRowView)[str2].ToString().Trim() + ";";
            }
            string[] strlength = str1.Split(';');
            if (strlength.Length > 0)
                str1 = str1.Substring(0, str1.Length - 1);
            return str1;
        }
        public static string CreateSqlfilter(string _listSqlKey, string _listSQlKeyValue)
        {
            string[] strArray1 = _listSqlKey.Split(';');
            string[] strArray2 = _listSQlKeyValue.Split(';');
            string str = "";
            for (int index = 0; index < strArray1.Length; ++index)
                str = str + " " + strArray1[index] + " = '" + strArray2[index] + "' and";
            return str.Length > 3 ? str.Substring(0, str.Length - 3) : string.Empty;
        }
        private void showWindow()
        {
            if (SysFunc.CheckPermission(StartupBase.SasObj, currActionTask, StartupBase.Menu_Id))
            {
                NewFrm Vouvherfile = new NewFrm();
                switch (currActionTask)
                {
                    case ActionTask.View:
                        Vouvherfile.Title = StartupBase.M_LAN.Equals("V") ? "Xem thông tin file đính kèm" : "Attach file";
                        Vouvherfile.EnableEditMode(false);
                        break;
                    case ActionTask.Add:
                        Vouvherfile.Title = StartupBase.M_LAN.Equals("V") ? "Thêm mới file đính kèm" : "Attach file";
                        Vouvherfile.EnableEditMode(true);
                        break;
                    case ActionTask.Edit:
                        Vouvherfile.Title = StartupBase.M_LAN.Equals("V") ? "Sửa thông tin file đính kèm" : "Attach file";
                        Vouvherfile.EnableEditMode(true);
                        break;
                    case ActionTask.Copy:
                        Vouvherfile.Title = StartupBase.M_LAN.Equals("V") ? "Sao chép file đính kèm" : "Attach file";
                        Vouvherfile.EnableEditMode(true);
                        break;
                }
                Vouvherfile.Closed += new EventHandler(this.window_Closed);
                Vouvherfile.ShowDialog();
            }
            else
            {
                int num = (int)ExMessageBox.Show(1945, StartupBase.SasObj, "Không có quyền", StartupBase.SasObj.GetSysvar("M_SAS_VER").ToString());
                currActionTask = ActionTask.None;
            }
            // StartUp.oBrowse.frmBrw.Focus();
        }
        private void window_Closed(object sender, EventArgs e)
        {
            VoucherConextMenu.currActionTask = ActionTask.None;
        }
        void SaveFile()
        {
            if (VoucherConextMenu.oBrowse.ActiveRecord == null)
                return;
            VoucherConextMenu.currSqlTableKeyValue = VoucherConextMenu.oBrowse.ActiveRecord.Cells[VoucherConextMenu.SqlTableKey].Value.ToString();
            string sql = "select ten_file,file_con from " + TableName + " where " + SqlTableKey + "='" + VoucherConextMenu.currSqlTableKeyValue + "'";
            SqlCommand cmd = new SqlCommand(sql);
            DataTable tblFile = StartupBase.SasObj.ExcuteReader(cmd).Tables[0];
            if (tblFile == null || tblFile.Rows.Count == 0)
                return;
            try
            {
                string file_name = tblFile.Rows[0]["ten_file"].ToString().Trim();
                byte[] Myfile = (byte[])tblFile.Rows[0]["file_con"];
                string fullpath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), file_name);
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.DefaultExt = "*";
                saveFileDialog.AddExtension = true;
                saveFileDialog.OverwritePrompt = false;
                saveFileDialog.AddExtension = true;
                saveFileDialog.Filter = "All|*.*";
                saveFileDialog.FileName = fullpath;
                bool? nullable = saveFileDialog.ShowDialog();
                if ((!nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) == 0)
                    return;

                if (Myfile.Length <= 0)
                    return;
                ByteArrayToFile(Myfile, saveFileDialog.FileName);
                if (!File.Exists(fullpath))
                {
                    int num = (int)ExMessageBox.Show(19207, StartupBase.SasObj, "Lỗi lưu file", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                }
            }
            catch
            {
                int num = (int)ExMessageBox.Show(12201, StartupBase.SasObj, "Lỗi lưu file", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
        }

        public bool ByteArrayToFile(byte[] byteArray, string filepath)
        {
            try
            {
                using (var fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(byteArray, 0, byteArray.Length);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in process: {0}", ex);
                return false;
            }
        }
        private void mnuRePost_Click(object sender, RoutedEventArgs e)
        {
            if (SasObj.UserInfo.Rows[0]["is_admin"].ToString().Trim() != "1")
            {
                int num0 = (int)ExMessageBox.Show(-951, SasObj, "Bạn không có quyền admin!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            string procpost = StartUpTrans.DmctInfo["proc_post"].ToString().Trim();
            if (String.IsNullOrEmpty(procpost))
            {
                int num1 = (int)ExMessageBox.Show(-952, SasObj, "Phiếu không lưu sổ !", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }
            if (StartUpTrans.DsTrans.Tables[0].Rows.Count <= 0)
                return;

            int result = 0;
            SasFormBrowes.FrmWaiting frm = new SasFormBrowes.FrmWaiting(70);
            frm.Set(0);
            frm.Show();
            foreach (DataRow row in StartUpTrans.DsTrans.Tables[0].Rows)
            {
                if (!String.IsNullOrEmpty(row["stt_rec"].ToString()) && !String.IsNullOrEmpty(row["ngay_ct"].ToString()))
                {
                    if (!procpost.Contains("["))
                        procpost = "[" + procpost + "]";
                    string Sql = "EXEC " + procpost + " '" + row["stt_rec"].ToString() + "'";
                    object res0 = SasObj.ExcuteScalar(new SqlCommand(Sql));
                    result++;
                    frm.Set(result * 100 / StartUpTrans.DsTrans.Tables[0].Rows.Count);
                }
            }
            frm.Set(70);
            frm.Hide();
            frm.Close();
            frm = null;
            int num2 = (int)ExMessageBox.Show(-954, SasObj, "Đã post [" + result.ToString().Trim() + "] phiếu!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            return;
        }
    }
}
