using System.Collections.Generic;
using System.ComponentModel;

namespace SasPermission
{
    public class Menu : INotifyPropertyChanged
    {
        private bool? _ischecked = new bool?(false);
        private bool? _isRead = new bool?(false);
        private bool? _isWrite = new bool?(false);
        private bool? _isEdit = new bool?(false);
        private bool? _isDelete = new bool?(false);
        private bool? _isPrint = new bool?(false);
        private string _menu_id;
        private string _menu_id0;
        private string _bar;
        private string _bar2;
        private string _ma_ct;
        private List<Menu> _SubMenuList;

        public string menu_id
        {
            get
            {
                return this._menu_id;
            }
            set
            {
                this._menu_id = value;
                this.OnPropertyChanged(nameof(menu_id));
            }
        }

        public string menu_id0
        {
            get
            {
                return this._menu_id0;
            }
            set
            {
                this._menu_id0 = value;
                this.OnPropertyChanged(nameof(menu_id0));
            }
        }

        public string bar
        {
            get
            {
                return this._bar;
            }
            set
            {
                this._bar = value;
                this.OnPropertyChanged(nameof(bar));
            }
        }

        public string bar2
        {
            get
            {
                return this._bar2;
            }
            set
            {
                this._bar2 = value;
                this.OnPropertyChanged(nameof(bar2));
            }
        }

        public string ma_ct
        {
            get
            {
                return this._ma_ct;
            }
            set
            {
                this._ma_ct = value;
                this.OnPropertyChanged(nameof(ma_ct));
            }
        }

        public bool? ischecked
        {
            get
            {
                return this._ischecked;
            }
            set
            {
                this._ischecked = value;
                this.OnPropertyChanged(nameof(ischecked));
            }
        }

        public bool? isRead
        {
            get
            {
                return this._isRead;
            }
            set
            {
                this._isRead = value;
                this.OnPropertyChanged(nameof(isRead));
            }
        }

        public bool? isWrite
        {
            get
            {
                return this._isWrite;
            }
            set
            {
                this._isWrite = value;
                this.OnPropertyChanged(nameof(isWrite));
            }
        }

        public bool? isEdit
        {
            get
            {
                return this._isEdit;
            }
            set
            {
                this._isEdit = value;
                this.OnPropertyChanged(nameof(isEdit));
            }
        }

        public bool? isDelete
        {
            get
            {
                return this._isDelete;
            }
            set
            {
                this._isDelete = value;
                this.OnPropertyChanged(nameof(isDelete));
            }
        }

        public bool? isPrint
        {
            get
            {
                return this._isPrint;
            }
            set
            {
                this._isPrint = value;
                this.OnPropertyChanged(nameof(isPrint));
            }
        }

        public List<Menu> SubMenuList
        {
            get
            {
                return this._SubMenuList;
            }
            set
            {
                this._SubMenuList = value;
                this.OnPropertyChanged(nameof(SubMenuList));
            }
        }

        private void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged == null)
                return;
            this.PropertyChanged((object)this, new PropertyChangedEventArgs(prop));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
