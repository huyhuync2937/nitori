using SasControls;
using SasDataLib;
using SasDefine;
using SasFormBrowes;
using SasVoucherLib;
using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using static System.Net.Mime.MediaTypeNames;

namespace Poctpxf
{
    public partial class FrmLoc : Form
    {
        private DataTable dtMa_vt;

        public FrmLoc()
        {
            this.InitializeComponent();
            //this.dtMa_vt = StartupBase.SasObj.ExcuteReader(new SqlCommand("SELECT CAST(0 as BIT) as tag, ma_vt, ten_vt, ten_vt2, dvt,sl_min as ton_cuoi FROM dmvt")).Tables[0];

        }
        private void FrmLoc_Loaded(object sender, RoutedEventArgs e)
        {
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtTungay.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            this.txtDenngay.Value = DateTime.Now.Date;
            this.txtTungay.Focus();
        }
        private void _confirmGridview_OnOk(object sender, RoutedEventArgs e)
        {
            Checkvalid();
            string filter = "";
            string conditionSD = "";

            SqlCommand sqlcmd = new SqlCommand();
            sqlcmd.CommandText = "SELECT link from dmlink where rtrim(ma_link) like 'code'";
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlcmd);
            string link = dataSet.Tables[0].Rows[0]["link"].ToString().Trim();

            if (!string.IsNullOrEmpty(this.txtma_kho.Text))
            {
                filter += " and ma_kho =''" + this.txtma_kho.Text.Trim() + "''";
                conditionSD += " and ma_kho =''" + this.txtma_kho.Text.Trim() + "''";

            }
            if (!string.IsNullOrEmpty(this.txtma_kh.Text))
            {
                filter += " and ma_kh =''" + this.txtma_kh.Text.Trim() + "''";
            }
            string sql = "Exec [INCD1_realtime]" + " '" + ((DateTime)this.txtTungay.Value).ToString("yyyyMMdd") + "', '" + ((DateTime)this.txtDenngay.Value).ToString("yyyyMMdd") + "', 0 ," +
                " '1=1   " + filter + "'" + " ,1," +
                "' 1=1  AND  ma_kho in (Select ma_kho From "+ link+ ".dbo.dmkho Where ma_dvcs like ''NITORI%'')" + conditionSD + "'";
            StartUp.HDBData = StartupBase.SasObj.ExcuteReader(new SqlCommand(sql));

            DataTable dataTable = StartUp.HDBData.Tables[0];

            COTKTH2Dvcs Frmlocdon = new COTKTH2Dvcs((DateTime)this.txtTungay.Value, (DateTime)this.txtDenngay.Value, filter, conditionSD);
            Frmlocdon.GrdCt.DataSource = (IEnumerable)dataTable.DefaultView;
            if (StartupBase.M_LAN != "V")
                Frmlocdon.Title = "Supplier list";
            bool? nullable = Frmlocdon.ShowDialog();

            if ((nullable.GetValueOrDefault() ? 0 : (nullable.HasValue ? 1 : 0)) != 0)
                return;
            StartUp.KhoNG = dataTable.Select("tag = 1");


            this.Close();
        }
        bool Checkvalid()
        {
            bool flag = false;
            if (Convert.ToDateTime(txtTungay.Value) > Convert.ToDateTime(txtDenngay.Value))
            {
                flag = true;
                MessageBox.Show("Từ ngày phải <= đến ngày.", "Thông báo");
                this.txtTungay.Focus();
            }
            return flag;
        }
        private void _confirmGridview_OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

       

       
    }
}
