using SasControls;
using SasFormBrowes;
using System.Windows;
using System.Windows.Input;

namespace COSXLSX.CODMNVL
{
    public partial class FrmTimKiem : Form
    {
        public bool isOk = false;
        public string Ma_sp { get; private set; } = "";
        public string Ma_vt { get; private set; } = "";
        public string Ma_px { get; private set; } = "";

        public FrmTimKiem(string ma_sp, string ma_vt, string ma_px)
        {
            this.InitializeComponent();
            SysFunc.LoadIcon((Window)this);
            this.BindingSasObj = StartupBase.SasObj;
            this.txtMa_sp.Text = ma_sp;
            this.txtMa_vt.Text = ma_vt;
            this.txtMa_px.Text = ma_px;
        }

        private void Form_Loaded(object sender, RoutedEventArgs e)
        {
            this.txtMa_sp.SearchInit();
            this.txtMa_vt.SearchInit();
            this.txtMa_px.SearchInit();
            this.txtMa_sp.IsFocus = true;
        }

        private void ConfirmGridView_OnOk(object sender, RoutedEventArgs e)
        {
            if (Keyboard.FocusedElement != null && Keyboard.FocusedElement.GetType().Equals(typeof(TextBoxAutoComplete)))
            {
                TextBoxAutoComplete focusedElement = Keyboard.FocusedElement as TextBoxAutoComplete;
                if (focusedElement.ParentControl != null && !focusedElement.ParentControl.CheckLostFocus())
                    return;
            }
            this.Ma_sp = this.txtMa_sp.Text.Trim();
            this.Ma_vt = this.txtMa_vt.Text.Trim();
            this.Ma_px = this.txtMa_px.Text.Trim();
            this.isOk = true;
            this.Close();
        }

        private void ConfirmGridView_OnCancel(object sender, RoutedEventArgs e)
        {
            this.isOk = false;
            this.Close();
        }

        private void txtMa_sp_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtMa_sp.SearchInit();
            this.lblTenSp.Text = this.txtMa_sp.RowResult == null ? "" : (StartupBase.M_LAN.Equals("V") ? this.txtMa_sp.RowResult["ten_vt"].ToString() : this.txtMa_sp.RowResult["ten_vt2"].ToString());
        }

        private void txtMa_vt_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtMa_vt.SearchInit();
            this.lblTenVt.Text = this.txtMa_vt.RowResult == null ? "" : this.txtMa_vt.RowResult["Ten_Vt"].ToString();
        }

        private void txtMa_px_PreviewLostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            this.txtMa_px.SearchInit();
            this.lblTenPx.Text = this.txtMa_px.RowResult == null ? "" : this.txtMa_px.RowResult["ten_px"].ToString();
        }
    }
}
