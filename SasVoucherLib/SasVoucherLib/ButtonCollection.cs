using System.Windows.Controls;

namespace SasVoucherLib
{
    public class ButtonCollection : ItemsControl
    {
        public int Count
        {
            get
            {
                return this.Items.Count;
            }
        }

        public int Add(ToolBarButton item)
        {
            return this.Items.Add((object)item);
        }

        public ToolBarButton this[int i]
        {
            get
            {
                return this.Items[i] as ToolBarButton;
            }
            set
            {
                this.Items[i] = (object)value;
            }
        }
    }
}
