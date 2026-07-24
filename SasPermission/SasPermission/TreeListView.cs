using System.Windows;
using System.Windows.Controls;

namespace SasPermission
{
    public class TreeListView : TreeView
    {
        public static readonly DependencyProperty AllowsColumnReorderProperty = DependencyProperty.Register(nameof(AllowsColumnReorder), typeof(bool), typeof(TreeListView), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));
        public static readonly DependencyProperty ColumnsProperty = DependencyProperty.Register(nameof(Columns), typeof(GridViewColumnCollection), typeof(TreeListView), (PropertyMetadata)new UIPropertyMetadata((PropertyChangedCallback)null));

        static TreeListView()
        {
            FrameworkElement.DefaultStyleKeyProperty.OverrideMetadata(typeof(TreeListView), (PropertyMetadata)new FrameworkPropertyMetadata((object)typeof(TreeListView)));
        }

        public TreeListView()
        {
            this.Columns = new GridViewColumnCollection();
        }

        public GridViewColumnCollection Columns
        {
            get
            {
                return (GridViewColumnCollection)this.GetValue(TreeListView.ColumnsProperty);
            }
            set
            {
                this.SetValue(TreeListView.ColumnsProperty, value);
            }
        }

        public bool AllowsColumnReorder
        {
            get
            {
                return (bool)this.GetValue(TreeListView.AllowsColumnReorderProperty);
            }
            set
            {
                this.SetValue(TreeListView.AllowsColumnReorderProperty, value);
            }
        }
    }
}
