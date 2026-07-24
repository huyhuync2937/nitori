using SasControls;
using SasFormBrowes;
using SasFormReport;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;
namespace PODHD1
{
    /// <summary>
    /// Interaction logic for Arttkb1F3.xaml
    /// </summary>
    public partial class Arttkb1F3 : FormFilter
    {
        private CheckBox[] list = (CheckBox[])null;
        public Arttkb1F3()
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.list = new CheckBox[3]
            {
        this.chk1,
        this.chk2,
        this.chk3
            };
            this.chk1.Focus();
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {         
            this.DialogResult = new bool?(true);
        }

        public string Selected
        {
            get
            {
                bool? isChecked = this.chk1.IsChecked;
                if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                    return "1";
                isChecked = this.chk2.IsChecked;
                return (!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0 ? "2" : "3";
            }
        }

        private void chk_Checked(object sender, RoutedEventArgs e)
        {
            if (this.list == null)
                return;
            foreach (CheckBox checkBox in this.list)
            {
                if (checkBox != sender)
                    checkBox.IsChecked = new bool?(false);
            }
        }

        private void chk_Unchecked(object sender, RoutedEventArgs e)
        {
            foreach (ToggleButton toggleButton in this.list)
            {
                bool? isChecked = toggleButton.IsChecked;
                if ((!isChecked.GetValueOrDefault() ? 0 : (isChecked.HasValue ? 1 : 0)) != 0)
                    return;
            }
          (sender as CheckBox).IsChecked = new bool?(true);
        }

    }
}
