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
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;

namespace SasVoucherLib
{
    public partial class DanhSachUsers : FormList
    {
        private List<UserModel> Users = (List<UserModel>)null;
        public string strusername = string.Empty;
        public string struserid = string.Empty;
        private string[] arrVoucher = (string[])null;
        private CodeValueBindingObject Voucher_Lan0;

        public DanhSachUsers()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
            this.LoadData();
        }

        private void DanhSachMaCT_Loaded(object sender, RoutedEventArgs e)
        {
            //this.Voucher_Lan0 = (CodeValueBindingObject)this.FindResource((object)"Voucher_Lan0");
            //this.Voucher_Lan0.Value = StartupBase.M_LAN.Equals("V");
            if (this.struserid.Trim() != string.Empty)
            {
                this.arrVoucher = this.struserid.Trim().Split(';');
                for (int index1 = 0; index1 < ((IEnumerable<string>)this.arrVoucher).Count<string>(); ++index1)
                {
                    bool flag = false;
                    for (int index2 = 0; index2 < this.Users.Count && !flag; ++index2)
                    {
                        if (this.Users[index2].user_id.Trim() == this.arrVoucher[index1].Trim())
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
                SqlCommand sqlcmd = new SqlCommand("SELECT * FROM userinfo");
                this.Users = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0].AsEnumerable().Select<DataRow, UserModel>((Func<DataRow, UserModel>)(v => new UserModel()
                {
                    user_id = v.Field<int>("user_id").ToString(),
                    user_name = v.Field<string>("user_name")
                })).ToList<UserModel>();
                this.GrdDsCT.DataContext = (object)new UserCommunityViewModel(this.Users);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        private void DanhSachMaCT_Closing(object sender, CancelEventArgs e)
        {
            string str1 = string.Empty;
            string str2 = string.Empty;
            for (int index = 0; index < this.Users.Count; ++index)
            {
                if (this.Users[index].IsChecked)
                {
                    str1 = str1 + ";" + this.Users[index].user_name.Trim();
                    str2 = str2 + ";" + this.Users[index].user_id.Trim();
                }
            }
            this.strusername = str1.Length == 0 ? str1 : str1.Substring(1);
            this.struserid = str2.Length == 0 ? str2 : str2.Substring(1);
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
                foreach (UserModel user in this.Users)
                    user.IsChecked = true;
            }
            if (e.Key == Key.U && Keyboard.Modifiers == ModifierKeys.Control)
            {
                foreach (UserModel user in this.Users)
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
