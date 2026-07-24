using System;
using System.Collections.Generic;
using System.Linq;

namespace SasPermission
{
    public class GenTreeMenu
    {
        private List<Menu> resultMenu = new List<Menu>();

        public List<Menu> AddTree(string sParentKey, IEnumerable<Menu> cList)
        {
            if ((from p in cList
                 where p.menu_id0.Trim() == sParentKey.Trim()
                 select p).Count<Menu>() > 0)
            {
                Menu childMenu;
                foreach (Menu childMenu2 in from p in cList
                                            where p.menu_id0.Trim() == sParentKey.Trim()
                                            select p)
                {
                    childMenu = childMenu2;
                    if (childMenu.menu_id != childMenu.menu_id0)
                    {
                        Menu menu = new Menu();
                        menu.menu_id = childMenu.menu_id;
                        menu.menu_id0 = childMenu.menu_id0;
                        menu.bar = childMenu.bar;
                        menu.bar2 = childMenu.bar2;
                        menu.ma_ct = childMenu.ma_ct;
                        menu.ischecked = childMenu.ischecked;
                        menu.isRead = childMenu.isRead;
                        menu.isWrite = childMenu.isWrite;
                        menu.isEdit = childMenu.isEdit;
                        menu.isDelete = childMenu.isDelete;
                        menu.isPrint = childMenu.isPrint;
                        if (cList.Count((Menu p) => p.menu_id0 == childMenu.menu_id) > 0)
                        {
                            menu.SubMenuList = new List<Menu>();
                            this.AddChildTree(childMenu.menu_id, cList, menu);
                        }
                        this.resultMenu.Add(menu);
                    }
                }
            }
            return this.resultMenu;
        }
        public void AddChildTree(string sParentKey, IEnumerable<Menu> cList, Menu currentMenu)
        {
            if ((from p in cList
                 where p.menu_id0.Trim() == sParentKey.Trim()
                 select p).Count<Menu>() > 0)
            {
                Menu childMenu;
                foreach (Menu childMenu2 in from p in cList
                                            where p.menu_id0.Trim() == sParentKey.Trim()
                                            select p)
                {
                    childMenu = childMenu2;
                    if (childMenu.menu_id != childMenu.menu_id0)
                    {
                        Menu menu = new Menu();
                        menu.menu_id = childMenu.menu_id;
                        menu.menu_id0 = childMenu.menu_id0;
                        menu.bar = childMenu.bar;
                        menu.bar2 = childMenu.bar2;
                        menu.ma_ct = childMenu.ma_ct;
                        menu.ischecked = childMenu.ischecked;
                        menu.isRead = childMenu.isRead;
                        menu.isWrite = childMenu.isWrite;
                        menu.isEdit = childMenu.isEdit;
                        menu.isDelete = childMenu.isDelete;
                        menu.isPrint = childMenu.isPrint;
                        if (cList.Count((Menu p) => p.menu_id0 == childMenu.menu_id) > 0)
                        {
                            menu.SubMenuList = new List<Menu>();
                            this.AddChildTree(childMenu.menu_id, cList, menu);
                        }
                        currentMenu.SubMenuList.Add(menu);
                    }
                }
            }
        }

        public List<Menu> AddTree2(string sParentKey, IEnumerable<Menu> cList)
        {
            if (cList.Where<Menu>((Func<Menu, bool>)(p => p.menu_id0.Trim() == sParentKey.Trim())).Count<Menu>() > 0)
            {
                using (IEnumerator<Menu> enumerator = cList.Where<Menu>((Func<Menu, bool>)(p => p.menu_id0.Trim() == sParentKey.Trim())).GetEnumerator())
                {
                    while (enumerator.MoveNext())
                    {
                        Menu childMenu = enumerator.Current;
                        if (childMenu.menu_id != childMenu.menu_id0)
                        {
                            Menu currentMenu = new Menu();
                            currentMenu.menu_id = childMenu.menu_id;
                            currentMenu.menu_id0 = childMenu.menu_id0;
                            currentMenu.bar = childMenu.bar;
                            currentMenu.bar2 = childMenu.bar2;
                            currentMenu.ma_ct = childMenu.ma_ct;
                            currentMenu.ischecked = childMenu.ischecked;
                            currentMenu.isRead = childMenu.isRead;
                            currentMenu.isWrite = childMenu.isWrite;
                            currentMenu.isEdit = childMenu.isEdit;
                            currentMenu.isDelete = childMenu.isDelete;
                            currentMenu.isPrint = childMenu.isPrint;
                            if (cList.Count<Menu>((Func<Menu, bool>)(p => p.menu_id0 == childMenu.menu_id)) > 0)
                            {
                                currentMenu.SubMenuList = new List<Menu>();
                                this.AddChildTree(childMenu.menu_id, cList, currentMenu);
                            }
                            this.resultMenu.Add(currentMenu);
                        }
                    }
                }
            }
            return this.resultMenu;
        }

        public void AddChildTree2(string sParentKey, IEnumerable<Menu> cList, Menu currentMenu)
        {
            if (cList.Where<Menu>((Func<Menu, bool>)(p => p.menu_id0.Trim() == sParentKey.Trim())).Count<Menu>() <= 0)
                return;
            using (IEnumerator<Menu> enumerator = cList.Where<Menu>((Func<Menu, bool>)(p => p.menu_id0.Trim() == sParentKey.Trim())).GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Menu childMenu = enumerator.Current;
                    if (childMenu.menu_id != childMenu.menu_id0)
                    {
                        Menu currentMenu1 = new Menu();
                        currentMenu1.menu_id = childMenu.menu_id;
                        currentMenu1.menu_id0 = childMenu.menu_id0;
                        currentMenu1.bar = childMenu.bar;
                        currentMenu1.bar2 = childMenu.bar2;
                        currentMenu1.ma_ct = childMenu.ma_ct;
                        currentMenu1.ischecked = childMenu.ischecked;
                        currentMenu1.isRead = childMenu.isRead;
                        currentMenu1.isWrite = childMenu.isWrite;
                        currentMenu1.isEdit = childMenu.isEdit;
                        currentMenu1.isDelete = childMenu.isDelete;
                        currentMenu1.isPrint = childMenu.isPrint;
                        if (cList.Count<Menu>((Func<Menu, bool>)(p => p.menu_id0 == childMenu.menu_id)) > 0)
                        {
                            currentMenu1.SubMenuList = new List<Menu>();
                            this.AddChildTree(childMenu.menu_id, cList, currentMenu1);
                        }
                        currentMenu.SubMenuList.Add(currentMenu1);
                    }
                }
            }
        }
    }
}
