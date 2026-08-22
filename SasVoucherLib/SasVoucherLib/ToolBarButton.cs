using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SasVoucherLib
{
    public class ToolBarButton : Button
    {
        public static readonly DependencyProperty IsNoTextProperty = DependencyProperty.Register(nameof(IsNoText), typeof(bool), typeof(ToolBarButton), (PropertyMetadata)new UIPropertyMetadata((object)false));
        public static readonly DependencyProperty ToolTip2Property = DependencyProperty.Register(nameof(ToolTip2), typeof(string), typeof(ToolBarButton), (PropertyMetadata)new UIPropertyMetadata((object)string.Empty));
        public static readonly DependencyProperty ImagePathProperty = DependencyProperty.Register(nameof(ImagePath), typeof(string), typeof(ToolBarButton), (PropertyMetadata)new UIPropertyMetadata((object)""));
        public static readonly DependencyProperty TextProperty = DependencyProperty.Register(nameof(Text), typeof(string), typeof(ToolBarButton), (PropertyMetadata)new UIPropertyMetadata((object)""));
        private TextBlock txt = new TextBlock();

        public bool IsNoText
        {
            get
            {
                return (bool)this.GetValue(ToolBarButton.IsNoTextProperty);
            }
            set
            {
                this.SetValue(ToolBarButton.IsNoTextProperty, (object)value);
            }
        }

        public string ToolTip2
        {
            get
            {
                return (string)this.GetValue(ToolBarButton.ToolTip2Property);
            }
            set
            {
                this.SetValue(ToolBarButton.ToolTip2Property, (object)value);
            }
        }

        public ToolBarButton()
        {
            this.Loaded += new RoutedEventHandler(this.ToolBarButton_Loaded);
        }

        private void ToolBarButton_Loaded(object sender, RoutedEventArgs e)
        {
            StackPanel stackPanel = new StackPanel();
            stackPanel.Orientation = Orientation.Horizontal;
            Image image = new Image();
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.UriSource = new Uri(this.ImagePath, UriKind.RelativeOrAbsolute);
            bitmapImage.EndInit();
            image.Stretch = Stretch.Fill;
            image.Source = (ImageSource)bitmapImage;
            image.Width = 24.0;
            image.Height = 24.0;
            stackPanel.Children.Add((UIElement)image);
            if (!this.IsNoText)
            {
                //image.Width = 25.0;
                //image.Height = 26.5;
                this.txt.Text = this.Text;
                this.txt.Name = "lblName";
                this.txt.Margin = new Thickness(5.0, 0.0, 5.0, 0.0);
                this.txt.VerticalAlignment = VerticalAlignment.Center;
                stackPanel.Children.Add((UIElement)this.txt);
            }
            stackPanel.Name = "stckButton";
            this.Content = (object)stackPanel;
            if (StartUpTrans.M_LAN.Equals("V"))
                return;
            this.ToolTip = (object)this.ToolTip2;
        }

        public string ImagePath
        {
            get
            {
                return (string)this.GetValue(ToolBarButton.ImagePathProperty);
            }
            set
            {
                this.SetValue(ToolBarButton.ImagePathProperty, (object)value);
            }
        }

        public string Text
        {
            get
            {
                return (string)this.GetValue(ToolBarButton.TextProperty);
            }
            set
            {
                this.SetValue(ToolBarButton.TextProperty, (object)value);
                this.txt.Text = this.Text;
            }
        }

        public static void OnImagePathChanged(
          DependencyObject sender,
          DependencyPropertyChangedEventArgs e)
        {

        }
    }
}
