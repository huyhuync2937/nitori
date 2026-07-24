using SasControls;
using SasFormBrowes;
using System;
using System.Windows;
using System.Windows.Threading;

namespace COSXKSX.KSXK
{
    public partial class FrmWaiting : Form
    {
        public double pgValue
        {
            get
            {
                return this.PBar.Maximum;
            }
            set
            {
                this.PBar.Maximum = value;
            }
        }

        public FrmWaiting(double maximum)
        {
            this.InitializeComponent();
            this.ShowInTaskbar = false;
            this.PBar.Maximum = maximum;
            this.PBar.Value = 0.0;
            this.Topmost = true;
            SysFunc.LoadIcon((Window)this);
            this.ActiveMainForm = false;
        }

        public void SetMax(double maximum)
        {
            this.PBar.Maximum = maximum;
        }

        public void Set(double value)
        {
            Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Background, (Delegate)new Action(() =>
           {
               try
               {
                   this.PBar.Value = value;
               }
               catch (Exception ex)
               {
               }
           }));
        }
    }
}
