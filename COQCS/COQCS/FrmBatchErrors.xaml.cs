using System.Collections.Generic;
using System.Windows;

namespace COQCS
{
    public sealed class CoqcsGroupErrorRow
    {
        public string LoaiMode { get; set; }
        public int Worker { get; set; }
        public string Error { get; set; }
    }

    // Màn hình hiện danh sách chi tiết các mã/nhóm chạy lỗi SAU KHI FrmBatchProgress đã đóng -
    // tách riêng khỏi message box tổng kết (ExMessageBox) vì FrmBatchProgress tự đóng sau 2s
    // nên không kịp đọc lưới lỗi bên trong nó, còn message box chỉ hiện được text, không xem/
    // cuộn được dạng lưới khi danh sách lỗi dài.
    public partial class FrmBatchErrors : Window
    {
        public FrmBatchErrors(string summary, IEnumerable<CoqcsBatchErrorRow> itemErrors, IEnumerable<CoqcsGroupErrorRow> groupErrors)
        {
            InitializeComponent();

            lblSummary.Text = summary;

            List<CoqcsBatchErrorRow> itemList = new List<CoqcsBatchErrorRow>(itemErrors);
            lvItemErrors.ItemsSource = itemList;
            gbItemErrors.Visibility = itemList.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

            List<CoqcsGroupErrorRow> groupList = new List<CoqcsGroupErrorRow>(groupErrors);
            lvGroupErrors.ItemsSource = groupList;
            gbGroupErrors.Visibility = groupList.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
