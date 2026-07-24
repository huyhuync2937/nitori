using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace CACTPC1
{
    public partial class FrmCACTPC1DSHD : Form
    {
        private bool isMePress = false;
        private CodeValueBindingObject Ip_ty_gia;

        public FrmCACTPC1DSHD()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.LoadData();
            this.SetStyle();
        }

        private void SetStyle()
        {
            string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            string str2 = StartUp.M_IP_TIEN_NT;
            if (str1 == StartUpTrans.M_ma_nt0)
                str2 = StartUp.M_IP_TIEN;
            string str3 = str2 + ";#";
            Style style = new Style(typeof(XamNumericEditor));
            Setter setter = new Setter(ValueEditor.FormatProperty, (object)str3);
            style.Setters.Add((SetterBase)setter);
            this.grdDSHD.FieldLayouts[0].Fields["tien_hd"].Settings.EditorStyle = style;
            this.grdDSHD.FieldLayouts[0].Fields["t_tien_dt"].Settings.EditorStyle = style;
            this.grdDSHD.FieldLayouts[0].Fields["tien_con_pt0"].Settings.EditorStyle = style;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Ip_ty_gia = (CodeValueBindingObject)this.FindResource((object)"Ip_ty_gia");
            this.Ip_ty_gia.Text = StartUp.M_IP_TY_GIA;
        }

        public void LoadData()
        {
            string empty1 = string.Empty;
            string empty2 = string.Empty;
            string empty3 = string.Empty;
            string str1 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["stt_rec"].ToString();
            string str2 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_kh"].ToString();
            string str3 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_dvcs"].ToString();
            string str4 = StartUpTrans.DsTrans.Tables[0].DefaultView[0]["ma_nt"].ToString();
            SqlCommand sqlcmd = new SqlCommand("Exec [CACTPC1-InitTt] @stt_rec, @ma_kh, @Ma_dvcs, @ma_nt");
            sqlcmd.Parameters.Add("@stt_rec", SqlDbType.Char, 11).Value = (object)str1;
            sqlcmd.Parameters.Add("@ma_kh", SqlDbType.Char, 16).Value = (object)str2;
            sqlcmd.Parameters.Add("@Ma_dvcs", SqlDbType.Char, 16).Value = (object)str3;
            sqlcmd.Parameters.Add("@ma_nt", SqlDbType.Char, 16).Value = (object)str4;
            DataTable dataTable = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].Copy();
            dataTable.DefaultView.Sort = "ngay_ct0, so_ct0";
            this.grdDSHD.DataSource = (IEnumerable)dataTable.DefaultView;
        }

        private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (!this.isMePress || e.Key != Key.Escape && e.Key != Key.Return)
                return;
            this.Close();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            this.isMePress = true;
        }

        public void ShowHd(string so_ct0, string stt_rec)
        {
            foreach (Record record in (IEnumerable<Record>)this.grdDSHD.Records)
            {
                if (record is DataRecord dataRecord && dataRecord.DataItem != null && so_ct0.Trim() != "" && dataRecord.Cells[nameof(so_ct0)].Value.ToString().Trim().ToUpper() == so_ct0.Trim().ToUpper())
                {
                    this.grdDSHD.ActiveRecord = record;
                    this.grdDSHD.Focus();
                    return;
                }
            }
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               if (this.grdDSHD.Records.Count <= 0)
                   return;
               this.grdDSHD.Focus();
               this.grdDSHD.ActiveRecord = this.grdDSHD.Records[0];
           }));
            this.ShowDialog();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
        }

    }
}
