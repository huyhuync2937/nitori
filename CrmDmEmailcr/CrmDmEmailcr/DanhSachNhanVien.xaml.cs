using Infragistics.Windows.DataPresenter;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using System;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace CrmDmEmailcr
{
    public partial class DanhSachNhanVien: FormList
    {
        private List<DmkhModel> Users = (List<DmkhModel>)null;
        public string strma_ns = string.Empty;
        private string[] arrVoucher = (string[])null;
        private CodeValueBindingObject Voucher_Lan0;

        public DanhSachNhanVien()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.LoadData();
        }

        private void DanhSachMaCT_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.strma_ns.Trim() != string.Empty)
            {
                this.arrVoucher = this.strma_ns.Trim().Split(';');
                for (int index1 = 0; index1 < ((IEnumerable<string>)this.arrVoucher).Count<string>(); ++index1)
                {
                    bool flag = false;
                    for (int index2 = 0; index2 < this.Users.Count && !flag; ++index2)
                    {
                        if (this.Users[index2].ma_kh.Trim() == this.arrVoucher[index1].Trim())
                        {
                            this.Users[index2].IsChecked = true;
                            flag = true;
                        }
                    }
                }
            }
            if (this.GrdDsCT.Records.Count <= 0)
                return;
            this.Dispatcher.BeginInvoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               this.GrdDsCT.ActiveRecord = (Record)(this.GrdDsCT.Records[0] as DataRecord);
               this.GrdDsCT.Focus();
           }));
        }

        private void LoadData()
        {
            try
            {
                DataTable dataTable = new DataTable();
                SqlCommand sqlcmd = new SqlCommand("SELECT ma_kh, ten_kh, ten_kh2, e_mail, dien_thoai FROM dmkhcr");
                this.Users = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].AsEnumerable().Select<DataRow, DmkhModel>((Func<DataRow, DmkhModel>)(v => new DmkhModel()
                {
                    ma_kh = v.Field<string>("ma_kh").ToString(),
                    ten_kh = v.Field<string>("ten_kh"),
                    e_mail = v.Field<string>("e_mail"),
                    dien_thoai = v.Field<string>("dien_thoai")
                })).ToList<DmkhModel>();
                this.GrdDsCT.DataContext = (object)new DmkhCommunityViewModel(this.Users);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void DanhSachMaCT_Closing(object sender, CancelEventArgs e)
        {
            string str1 = string.Empty;
            for (int index = 0; index < this.Users.Count; ++index)
            {
                if (this.Users[index].IsChecked)
                {
                    str1 = str1 + ";" + this.Users[index].ma_kh.Trim();
                }
            }
            this.strma_ns = str1.Length == 0 ? str1 : str1.Substring(1);
        }

        private void GrdDsCT_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!Keyboard.IsKeyDown(Key.Space))
                return;
            int index = this.GrdDsCT.ActiveRecord.Index;
            this.Users[index].IsChecked = !this.Users[index].IsChecked;
        }

        private void FormList_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.A && Keyboard.Modifiers == ModifierKeys.Control)
            {
                foreach (DmkhModel user in this.Users)
                    user.IsChecked = true;
            }
            if (e.Key == Key.U && Keyboard.Modifiers == ModifierKeys.Control)
            {
                foreach (DmkhModel user in this.Users)
                    user.IsChecked = false;
            }
            if (e.Key != Key.Escape || Keyboard.Modifiers != ModifierKeys.None)
                return;
            this.Close();
        }

        private void grdMain_OnOk(object sender, RoutedEventArgs e)
        {
            this.DialogResult = new bool?(true);
        }

    }
}
