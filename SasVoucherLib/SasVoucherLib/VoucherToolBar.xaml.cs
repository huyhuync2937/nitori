using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace SasVoucherLib
{
    /// <summary>Interaction logic for VoucherToolBar.xaml</summary>
    /// <summary>VoucherToolBar</summary>
    public partial class VoucherToolBar : ToolBar
    {
        public static readonly DependencyProperty DisableOpacityProperty = DependencyProperty.Register(nameof(DisableOpacity), typeof(double), typeof(VoucherToolBar), (PropertyMetadata)new UIPropertyMetadata((object)0.3));
        /// <summary>
        /// Using a DependencyProperty as the backing store for IsInEditMode.  This enables animation, styling, binding, etc...
        /// </summary>
        public static readonly DependencyProperty IsInEditModeProperty = DependencyProperty.Register(nameof(IsInEditMode), typeof(bool), typeof(VoucherToolBar), (PropertyMetadata)new UIPropertyMetadata((object)false, new PropertyChangedCallback(VoucherToolBar.OnEditModeChanged)));
        public Button a = new Button();

        /// <summary>VoucherToolBar constructor.</summary>
        public VoucherToolBar()
        {
            this.InitializeComponent();
            this.OnIsInEditModeChanged();
        }

        /// <summary>Add tool bar button click handle.</summary>
        /// <param name="handler"></param>
        public void AddButtonClick(Delegate handler)
        {
            this.AddHandler(ButtonBase.ClickEvent, handler);
        }

        /// <summary>Remove tool bar button click handle.</summary>
        /// <param name="handler"></param>
        public void RemoveButtonClick(RoutedEventHandler handler)
        {
            this.RemoveHandler(ButtonBase.ClickEvent, (Delegate)handler);
        }

        /// <summary>return tool bar button collection.</summary>
        public ButtonCollection Buttons
        {
            get
            {
                if (this.Items.Count == 0)
                    return (ButtonCollection)null;
                ButtonCollection buttonCollection = new ButtonCollection();
                for (int index = 0; index < this.Items.Count; ++index)
                    buttonCollection.Add(this.Items[index] as ToolBarButton);
                return buttonCollection;
            }
        }

        /// <summary>
        /// Button disable Opacity. This is a dependency property.
        /// </summary>
        public double DisableOpacity
        {
            get
            {
                return (double)this.GetValue(VoucherToolBar.DisableOpacityProperty);
            }
            set
            {
                this.SetValue(VoucherToolBar.DisableOpacityProperty, (object)value);
            }
        }

        /// <summary>
        /// Get/set tool bar mode. View or Edit moding. This is a dependency property.
        /// </summary>
        public bool IsInEditMode
        {
            get
            {
                return (bool)this.GetValue(VoucherToolBar.IsInEditModeProperty);
            }
            set
            {
                this.SetValue(VoucherToolBar.IsInEditModeProperty, (object)value);
                this.OnIsInEditModeChanged();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void OnEditModeChanged(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {
            (sender as VoucherToolBar).OnIsInEditModeChanged();
        }

        /// <summary>
        /// Setup button setting when IsInEditMode property changed.
        /// </summary>
        /// <param name="d"></param>
        /// <param name="e"></param>
        private void OnIsInEditModeChanged()
        {
            this.btnSave.IsEnabled = this.IsInEditMode;
            this.btnCancel.IsEnabled = this.IsInEditMode;
            this.btnNew.IsEnabled = !this.IsInEditMode;
            this.btnEdit.IsEnabled = !this.IsInEditMode;
            this.btnCopy.IsEnabled = !this.IsInEditMode;
            this.btnPrint.IsEnabled = !this.IsInEditMode;
            this.btnDelete.IsEnabled = !this.IsInEditMode;
            this.btnView.IsEnabled = !this.IsInEditMode;
            this.btnSearch.IsEnabled = !this.IsInEditMode;
            this.btnTop.IsEnabled = !this.IsInEditMode;
            this.btnPrevious.IsEnabled = !this.IsInEditMode;
            this.btnNext.IsEnabled = !this.IsInEditMode;
            this.btnBottom.IsEnabled = !this.IsInEditMode;
            if (this.IsInEditMode)
            {
                this.btnSave.Opacity = 1.0;
                this.btnCancel.Opacity = 1.0;
                this.btnNew.Opacity = this.DisableOpacity;
                this.btnEdit.Opacity = this.DisableOpacity;
                this.btnCopy.Opacity = this.DisableOpacity;
                this.btnPrint.Opacity = this.DisableOpacity;
                this.btnDelete.Opacity = this.DisableOpacity;
                this.btnView.Opacity = this.DisableOpacity;
                this.btnSearch.Opacity = this.DisableOpacity;
                this.btnTop.Opacity = this.DisableOpacity;
                this.btnPrevious.Opacity = this.DisableOpacity;
                this.btnNext.Opacity = this.DisableOpacity;
                this.btnBottom.Opacity = this.DisableOpacity;
            }
            else
            {
                this.btnSave.Opacity = this.DisableOpacity;
                this.btnCancel.Opacity = this.DisableOpacity;
                this.btnNew.Opacity = 1.0;
                this.btnEdit.Opacity = 1.0;
                this.btnCopy.Opacity = 1.0;
                this.btnPrint.Opacity = 1.0;
                this.btnDelete.Opacity = 1.0;
                this.btnView.Opacity = 1.0;
                this.btnSearch.Opacity = 1.0;
                this.btnTop.Opacity = 1.0;
                this.btnPrevious.Opacity = 1.0;
                this.btnNext.Opacity = 1.0;
                this.btnBottom.Opacity = 1.0;
            }
        }

    }
}

