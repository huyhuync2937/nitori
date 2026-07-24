using SasControls;
using SasDataLib;
using SasFormBrowes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;

namespace SasPermission
{
    public partial class FrmPermission : Form
    {
        private IEnumerable<Menu> cList = (IEnumerable<Menu>)null;
        private List<Menu> MenuList = (List<Menu>)null;
        public int rowIndex = -1;
        private DataTable OldRow = (DataTable)null;
        private string _strPermission = "";
        private DataTable newDataTable;
        private bool flag;
        public string Is_mobile = "0";

        public FrmPermission()
        {
            this.InitializeComponent();
            this.DisplayLanguage = StartupBase.M_LAN;
            SysFunc.LoadIcon((Window)this);
        }

        private void FrmPermission_Loaded(object sender, RoutedEventArgs e)
        {           
            DataTable dataTable = new DataTable();
            //SqlCommand sqlcmd = new SqlCommand("select menu_id, menu_id0, bar, bar2, RTRIM(ma_ct) as ma_ct from command where len(rtrim(menu_id)) > 7 and hide_yn = 0 AND ma_phan_he <> '' and IsNull(is_mobile,'0') = '" + Is_mobile.Trim() + "'  order by menu_id");
            SqlCommand sqlcmd = new SqlCommand("select menu_id, menu_id0, bar, bar2, RTRIM(ma_ct) as ma_ct from command where len(rtrim(menu_id)) > 7 and hide_yn = 0 AND ma_phan_he <> '' and IsNull(is_mobile,'0') = '" + Is_mobile.Trim() + "' or IsNull(is_mobile,'0') = '2'  order by menu_id");

            DataTable table = StartupBase.SasObj.ExcuteReader(sqlcmd).Tables[0];
            this.Title = (StartupBase.M_LAN.Equals("V") ? "Phan quyen cho NSD: " : "User permissions: ") + StartUp.dt.Rows[this.rowIndex][StartUp.SqlTableObjectName].ToString().ToUpper().Trim();
            this.Title += StartupBase.M_LAN.Equals("V") ? ", Ctrl+A: Chon tat ca, Ctrl+U: Huy chon tat ca" : ", Ctrl+A: Select All, Ctrl+U: Unselect All";
            this.cList = (IEnumerable<Menu>)table.AsEnumerable().Select<DataRow, Menu>((Func<DataRow, Menu>)(currMenuList => new Menu()
            {
                menu_id = currMenuList.Field<string>("menu_id"),
                menu_id0 = currMenuList.Field<string>("menu_id0"),
                bar = currMenuList.Field<string>("bar"),
                bar2 = currMenuList.Field<string>("bar2"),
                ma_ct = currMenuList.Field<string>("ma_ct")
            }));
            this.MenuList = new GenTreeMenu().AddTree("", this.cList);
            this.setPermissionMenu("rights", "menu");
            this.setPermissionMenu("r_read", "read");
            this.setPermissionMenu("r_add", "write");
            this.setPermissionMenu("r_edit", "edit");
            this.setPermissionMenu("r_del", "delete");
            this.setPermissionMenu("r_print", "print");
            this.newDataTable = StartUp.GetRow(StartUp.sqlTableName);
            if (this.newDataTable.Rows.Count > 0)
                this.OldRow = this.newDataTable.Copy();

            this.MenuTreeView.DataContext = this.MenuList;
            this.btnOk.Focus();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "menu");
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "menu");
        }

        private void IsCheckedParentMenuList(
          List<Menu> parentMenuList,
          bool isChecked,
          string typeChecked)
        {
            foreach (Menu parentMenu in parentMenuList)
            {
                if (parentMenu != null && parentMenu.SubMenuList != null)
                {
                    switch (typeChecked)
                    {
                        case "menu":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? ischecked = p.ischecked;
                               bool flag = isChecked;
                               return (ischecked.GetValueOrDefault() != flag ? 0 : (ischecked.HasValue ? 1 : 0)) != 0 || !p.ischecked.HasValue;
                           })) == null)
                            {
                                parentMenu.ischecked = new bool?(!isChecked);
                                if (isChecked)
                                {
                                    parentMenu.isRead = new bool?(false);
                                    parentMenu.isWrite = new bool?(false);
                                    parentMenu.isEdit = new bool?(false);
                                    parentMenu.isDelete = new bool?(false);
                                    parentMenu.isPrint = new bool?(false);
                                    break;
                                }
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? ischecked = p.ischecked;
                               bool flag = isChecked;
                               return ischecked.GetValueOrDefault() == flag && ischecked.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                                parentMenu.ischecked = new bool?();
                            break;
                        case "read":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isRead = p.isRead;
                               bool flag = isChecked;
                               return (isRead.GetValueOrDefault() != flag ? 0 : (isRead.HasValue ? 1 : 0)) != 0 || !p.isRead.HasValue;
                           })) == null)
                            {
                                parentMenu.isRead = new bool?(!isChecked);
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isRead = p.isRead;
                               bool flag = isChecked;
                               return isRead.GetValueOrDefault() == flag && isRead.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                            {
                                parentMenu.isRead = new bool?();
                                break;
                            }
                            break;
                        case "write":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isWrite = p.isWrite;
                               bool flag = isChecked;
                               return (isWrite.GetValueOrDefault() != flag ? 0 : (isWrite.HasValue ? 1 : 0)) != 0 || !p.isWrite.HasValue;
                           })) == null)
                            {
                                parentMenu.isWrite = new bool?(!isChecked);
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isWrite = p.isWrite;
                               bool flag = isChecked;
                               return isWrite.GetValueOrDefault() == flag && isWrite.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                            {
                                parentMenu.isWrite = new bool?();
                                break;
                            }
                            break;
                        case "edit":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isEdit = p.isEdit;
                               bool flag = isChecked;
                               return (isEdit.GetValueOrDefault() != flag ? 0 : (isEdit.HasValue ? 1 : 0)) != 0 || !p.isEdit.HasValue;
                           })) == null)
                            {
                                parentMenu.isEdit = new bool?(!isChecked);
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isEdit = p.isEdit;
                               bool flag = isChecked;
                               return isEdit.GetValueOrDefault() == flag && isEdit.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                            {
                                parentMenu.isEdit = new bool?();
                                break;
                            }
                            break;
                        case "delete":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isDelete = p.isDelete;
                               bool flag = isChecked;
                               return (isDelete.GetValueOrDefault() != flag ? 0 : (isDelete.HasValue ? 1 : 0)) != 0 || !p.isDelete.HasValue;
                           })) == null)
                            {
                                parentMenu.isDelete = new bool?(!isChecked);
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isDelete = p.isDelete;
                               bool flag = isChecked;
                               return isDelete.GetValueOrDefault() == flag && isDelete.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                            {
                                parentMenu.isDelete = new bool?();
                                break;
                            }
                            break;
                        case "print":
                            if (parentMenu.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isPrint = p.isPrint;
                               bool flag = isChecked;
                               return (isPrint.GetValueOrDefault() != flag ? 0 : (isPrint.HasValue ? 1 : 0)) != 0 || !p.isPrint.HasValue;
                           })) == null)
                            {
                                parentMenu.isPrint = new bool?(!isChecked);
                                break;
                            }
                            if (parentMenu.SubMenuList.Where<Menu>((Func<Menu, bool>)(p =>
                           {
                               bool? isPrint = p.isPrint;
                               bool flag = isChecked;
                               return isPrint.GetValueOrDefault() == flag && isPrint.HasValue;
                           })).Count<Menu>() < parentMenu.SubMenuList.Count)
                            {
                                parentMenu.isPrint = new bool?();
                                break;
                            }
                            break;
                    }
                }
            }
        }

        private List<Menu> IsCheckedMenu(string id, bool ischecked, string typeChecked)
        {
            List<Menu> parentMenuList = new List<Menu>();
            Menu currMenu = this.MenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id));
            if (currMenu != null)
            {
                parentMenuList.Add(currMenu);
                switch (typeChecked)
                {
                    case "menu":
                        currMenu.ischecked = new bool?(ischecked);
                        if (!ischecked)
                        {
                            currMenu.isRead = new bool?(false);
                            currMenu.isWrite = new bool?(false);
                            currMenu.isEdit = new bool?(false);
                            currMenu.isDelete = new bool?(false);
                            currMenu.isPrint = new bool?(false);
                            break;
                        }
                        break;
                    case "read":
                        currMenu.isRead = new bool?(ischecked);
                        break;
                    case "write":
                        currMenu.isWrite = new bool?(ischecked);
                        break;
                    case "edit":
                        currMenu.isEdit = new bool?(ischecked);
                        break;
                    case "delete":
                        currMenu.isDelete = new bool?(ischecked);
                        break;
                    case "print":
                        currMenu.isPrint = new bool?(ischecked);
                        break;
                }
                this.MenuChecked(currMenu, ischecked, typeChecked);
            }
            else
                this.IDPosision(id, this.MenuList, ischecked, ref parentMenuList, typeChecked);
            return parentMenuList;
        }

        private void IDPosision(
          string id,
          List<Menu> mList,
          bool isChecked,
          ref List<Menu> parentMenuList,
          string typeChecked)
        {
            foreach (Menu m in mList)
            {
                this.flag = false;
                if (m.SubMenuList != null)
                {
                    if (m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)) != null)
                    {
                        parentMenuList.Add(m);
                        switch (typeChecked)
                        {
                            case "menu":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).ischecked = new bool?(isChecked);
                                if (!isChecked)
                                {
                                    m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isRead = new bool?(false);
                                    m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isWrite = new bool?(false);
                                    m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isEdit = new bool?(false);
                                    m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isDelete = new bool?(false);
                                    m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isPrint = new bool?(false);
                                    break;
                                }
                                break;
                            case "read":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isRead = new bool?(isChecked);
                                break;
                            case "write":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isWrite = new bool?(isChecked);
                                break;
                            case "edit":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isEdit = new bool?(isChecked);
                                break;
                            case "delete":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isDelete = new bool?(isChecked);
                                break;
                            case "print":
                                m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)).isPrint = new bool?(isChecked);
                                break;
                        }
                        this.MenuChecked(m.SubMenuList.FirstOrDefault<Menu>((Func<Menu, bool>)(p => p.menu_id.Trim() == id)), isChecked, typeChecked);
                        this.flag = true;
                        break;
                    }
                    this.IDPosision(id, m.SubMenuList, isChecked, ref parentMenuList, typeChecked);
                }
                if (this.flag)
                {
                    parentMenuList.Add(m);
                    break;
                }
            }
        }
        private void MenuChecked(Menu currMenu, bool isChecked, string typeChecked)
        {
            if (currMenu.SubMenuList == null)
                return;
            foreach (Menu subMenu in currMenu.SubMenuList)
            {
                bool isChecked1 = isChecked;
                if (this._strPermission != "" && !this._strPermission.Contains(subMenu.menu_id))
                    isChecked1 = false;
                switch (typeChecked)
                {
                    case "menu":
                        subMenu.ischecked = new bool?(isChecked1);
                        if (!isChecked1)
                        {
                            subMenu.isRead = new bool?(false);
                            subMenu.isWrite = new bool?(false);
                            subMenu.isEdit = new bool?(false);
                            subMenu.isDelete = new bool?(false);
                            subMenu.isPrint = new bool?(false);
                            break;
                        }
                        break;
                    case "read":
                        subMenu.isRead = new bool?(isChecked1);
                        break;
                    case "write":
                        subMenu.isWrite = new bool?(isChecked1);
                        break;
                    case "edit":
                        subMenu.isEdit = new bool?(isChecked1);
                        break;
                    case "delete":
                        subMenu.isDelete = new bool?(isChecked1);
                        break;
                    case "print":
                        int num = subMenu.SubMenuList != null ? 0 : (subMenu.SubMenuList != null ? 1 : (!(subMenu.ma_ct != "") ? 1 : 0));
                        subMenu.isPrint = num != 0 ? new bool?(false) : new bool?(isChecked1);
                        break;
                }
                this.MenuChecked(subMenu, isChecked1, typeChecked);
            }
        }

        private void MenuCheckedALL(Menu currMenu, bool isChecked, string typeChecked)
        {
            if (currMenu.SubMenuList == null)
                return;
            foreach (Menu subMenu in currMenu.SubMenuList)
            {
                bool isChecked1 = isChecked;
                //if (this._strPermission != "" && !this._strPermission.Contains(subMenu.menu_id))
                  //  isChecked1 = false;
                switch (typeChecked)
                {
                    case "menu":
                        subMenu.ischecked = new bool?(isChecked1);
                        if (!isChecked1)
                        {
                            subMenu.isRead = new bool?(false);
                            subMenu.isWrite = new bool?(false);
                            subMenu.isEdit = new bool?(false);
                            subMenu.isDelete = new bool?(false);
                            subMenu.isPrint = new bool?(false);
                            break;
                        }
                        break;
                    case "read":
                        subMenu.isRead = new bool?(isChecked1);
                        break;
                    case "write":
                        subMenu.isWrite = new bool?(isChecked1);
                        break;
                    case "edit":
                        subMenu.isEdit = new bool?(isChecked1);
                        break;
                    case "delete":
                        subMenu.isDelete = new bool?(isChecked1);
                        break;
                    case "print":
                        int num = subMenu.SubMenuList != null ? 0 : (subMenu.SubMenuList != null ? 1 : (!(subMenu.ma_ct != "") ? 1 : 0));
                        subMenu.isPrint = num != 0 ? new bool?(false) : new bool?(isChecked1);
                        break;
                }
                this.MenuCheckedALL(subMenu, isChecked1, typeChecked);
            }
        }

        private void PermissionMenu(
          List<Menu> _MenuList,
          ref StringBuilder strPermission,
          string typeChecked)
        {
            if (_MenuList == null)
                return;
            bool? nullable;
            foreach (Menu menu in _MenuList)
            {
                switch (typeChecked)
                {
                    case "menu":
                        nullable = menu.ischecked;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                            this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);
                            break;
                        }
                        break;
                    case "read":
                        nullable = menu.isRead;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                        }
                        this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);
                        break;
                    case "write":
                        nullable = menu.isWrite;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                        }
                        this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);
                        break;
                    case "edit":
                        nullable = menu.isEdit;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                        }
                        this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);
                        break;
                    case "delete":
                        nullable = menu.isDelete;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                        }
                        this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);                        
                        break;
                    case "print":
                        nullable = menu.isPrint;
                        if ((nullable.GetValueOrDefault() ? 1 : (!nullable.HasValue ? 1 : 0)) != 0)
                        {
                            strPermission.Append(menu.menu_id.Trim() + "/");
                        }
                        this.PermissionMenu(menu.SubMenuList, ref strPermission, typeChecked);                            
                        break;
                }
            }
        }

        private void CheckedOrUnChecked(object sender, bool isChecked, string typeChecked)
        {
            this._strPermission = "";
            this.IsCheckedParentMenuList(this.IsCheckedMenu(((FrameworkElement)sender).Tag.ToString().Trim(), isChecked, typeChecked), !isChecked, typeChecked);
        }

        private void write_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "write");
        }

        private void write_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "write");
        }

        private void edit_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "edit");
        }

        private void edit_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "edit");
        }

        private void delete_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "delete");
        }

        private void delete_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "delete");
        }

        private void read_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "read");
        }

        private void read_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "read");
        }

        private void print_Checked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, true, "print");
        }

        private void print_Unchecked(object sender, RoutedEventArgs e)
        {
            this.CheckedOrUnChecked(sender, false, "print");
        }

        private void setPermissionMenu(string sqlColumnName, string typeChecked)
        {
            if (StartUp.dt.Rows[this.rowIndex][sqlColumnName] == null)
                return;
            this._strPermission = StartUp.dt.Rows[this.rowIndex][sqlColumnName].ToString();
            string[] strArray = this._strPermission.Split('/');
            for (int index = 0; index < ((IEnumerable<string>)strArray).Count<string>(); ++index)
            {
                if (strArray[index] != string.Empty)
                    this.IsCheckedParentMenuList(this.IsCheckedMenu(strArray[index], true, typeChecked), false, typeChecked);
            }
        }

        private void Form_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.A)
            {
                foreach (Menu menu in this.MenuList)
                {
                    menu.ischecked = new bool?(true);
                    menu.isRead = new bool?(true);
                    menu.isWrite = new bool?(true);
                    menu.isEdit = new bool?(true);
                    menu.isDelete = new bool?(true);
                    menu.isPrint = new bool?(true);
                    this.MenuCheckedALL(menu, true, "menu");
                    this.MenuCheckedALL(menu, true, "read");
                    this.MenuCheckedALL(menu, true, "write");
                    this.MenuCheckedALL(menu, true, "edit");
                    this.MenuCheckedALL(menu, true, "delete");
                    this.MenuCheckedALL(menu, true, "print");
                }
            }
            if (Keyboard.Modifiers != ModifierKeys.Control || e.Key != Key.U)
                return;
            foreach (Menu menu in this.MenuList)
            {
                menu.ischecked = new bool?(false);
                this.MenuChecked(menu, false, "menu");
            }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder strPermission1 = new StringBuilder();
            StringBuilder strPermission2 = new StringBuilder();
            StringBuilder strPermission3 = new StringBuilder();
            StringBuilder strPermission4 = new StringBuilder();
            StringBuilder strPermission5 = new StringBuilder();
            StringBuilder strPermission6 = new StringBuilder();
            this.PermissionMenu(this.MenuList, ref strPermission1, "menu");
            this.PermissionMenu(this.MenuList, ref strPermission2, "read");
            this.PermissionMenu(this.MenuList, ref strPermission3, "write");
            this.PermissionMenu(this.MenuList, ref strPermission4, "edit");
            this.PermissionMenu(this.MenuList, ref strPermission5, "delete");
            this.PermissionMenu(this.MenuList, ref strPermission6, "print");
            string NonUpdateColumns = "user_id;user_name;user_pre;password;comment;is_admin";
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("rights", strPermission1);
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("r_read", strPermission2);
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("r_add", strPermission3);
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("r_edit", strPermission4);
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("r_del", strPermission5);
            StartUp.dt.Rows[this.rowIndex].SetField<StringBuilder>("r_print", strPermission6);
            this.newDataTable.Rows[0]["rights"] = (object)StartUp.dt.Rows[this.rowIndex]["rights"].ToString();
            this.newDataTable.Rows[0]["r_read"] = (object)StartUp.dt.Rows[this.rowIndex]["r_read"].ToString();
            this.newDataTable.Rows[0]["r_add"] = (object)StartUp.dt.Rows[this.rowIndex]["r_add"].ToString();
            this.newDataTable.Rows[0]["r_edit"] = (object)StartUp.dt.Rows[this.rowIndex]["r_edit"].ToString();
            this.newDataTable.Rows[0]["r_del"] = (object)StartUp.dt.Rows[this.rowIndex]["r_del"].ToString();
            this.newDataTable.Rows[0]["r_print"] = (object)StartUp.dt.Rows[this.rowIndex]["r_print"].ToString();
            if (ListFunc.updateRowInDatabaseByKey(StartUp.sqlTableName, StartUp.SqlTableKey, this.OldRow.Rows[0], this.newDataTable.Rows[0], StartupBase.SasObj, NonUpdateColumns) != 1)
                return;
            StartUp.dt.Rows[this.rowIndex].AcceptChanges();
            StartUp.dt.Rows[this.rowIndex].BeginEdit();
            StartUp.dt.Rows[this.rowIndex].ItemArray = StartUp.dt.Rows[this.rowIndex].ItemArray;
            StartUp.dt.Rows[this.rowIndex].EndEdit();
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}
