using System.ComponentModel;

namespace SasVoucherLib
{
    public class UserModel : INotifyPropertyChanged
    {
        private bool _isChecked;

        public bool IsChecked
        {
            get
            {
                return this._isChecked;
            }
            set
            {
                if (value == this._isChecked)
                    return;
                this._isChecked = value;
                this.OnPropertyChanged(nameof(IsChecked));
            }
        }

        public string user_id { get; set; }

        public string user_name { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged == null)
                return;
            this.PropertyChanged((object)this, new PropertyChangedEventArgs(prop));
        }
    }
}
