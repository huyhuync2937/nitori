using System.Collections.Generic;
using System.ComponentModel;

namespace SasVoucherLib
{
    public class UserCommunityViewModel : INotifyPropertyChanged
    {
        private List<UserModel> _User;

        public UserCommunityViewModel(List<UserModel> users)
        {
            this.Users = users;
            foreach (UserModel user in users)
                user.PropertyChanged += (PropertyChangedEventHandler)((sender, e) =>
               {
                   if (!(e.PropertyName == "IsChecked"))
                       return;
                   this.OnPropertyChanged(nameof(AllVouchersAreChecked));
               });
        }

        public List<UserModel> Users
        {
            get
            {
                return this._User;
            }
            set
            {
                if (value == this._User)
                    return;
                this._User = value;
                this.OnPropertyChanged(nameof(Users));
            }
        }

        public bool? AllVouchersAreChecked
        {
            get
            {
                bool? nullable1 = new bool?();
                for (int index = 0; index < this.Users.Count; ++index)
                {
                    if (index == 0)
                    {
                        nullable1 = new bool?(this.Users[0].IsChecked);
                    }
                    else
                    {
                        bool? nullable2 = nullable1;
                        bool isChecked = this.Users[index].IsChecked;
                        if ((nullable2.GetValueOrDefault() != isChecked ? 1 : (!nullable2.HasValue ? 1 : 0)) != 0)
                        {
                            nullable1 = new bool?();
                            break;
                        }
                    }
                }
                return nullable1;
            }
            set
            {
                if (!value.HasValue)
                    return;
                foreach (UserModel user in this.Users)
                    user.IsChecked = value.Value;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged == null)
                return;
            this.PropertyChanged((object)this, new PropertyChangedEventArgs(prop));
        }
    }
}
