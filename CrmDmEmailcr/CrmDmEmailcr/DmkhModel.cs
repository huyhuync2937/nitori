using System.ComponentModel;

namespace CrmDmEmailcr
{
    public class DmkhModel : INotifyPropertyChanged
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

        public string ma_kh { get; set; }

        public string ten_kh { get; set; }
        public string e_mail { get; set; }
        public string dien_thoai { get; set; }


        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged == null)
                return;
            this.PropertyChanged((object)this, new PropertyChangedEventArgs(prop));
        }
    }
}
