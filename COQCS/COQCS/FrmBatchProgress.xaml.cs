using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Threading;

namespace COQCS
{
    public sealed class CoqcsBatchErrorRow
    {
        public string MaVt { get; set; }
        public string Loai { get; set; }
        public int Worker { get; set; }
        public string Loi { get; set; }
    }

    // Màn hình tiến trình chạy QCS theo batch song song (dùng chung cho cả "ngang" và "dọc" -
    // xem StartUp.RunQcsBatchParallel): poll dbo.COQCS_BatchLog theo run_id (bảng do
    // dbo.COQCS_ngang_batch/COQCS_doc_batch ghi trực tiếp khi chạy từng mã) để hiển thị tiến độ
    // và mã lỗi theo thời gian thực, thay vì chỉ có 1 message box tổng kết sau khi chạy xong hết.
    public partial class FrmBatchProgress : Window
    {
        private readonly string _connectionString;
        private readonly Guid _runId;
        private readonly int _total;
        private readonly DispatcherTimer _timer;
        private readonly ObservableCollection<CoqcsBatchErrorRow> _errors = new ObservableCollection<CoqcsBatchErrorRow>();
        private readonly HashSet<long> _knownErrorIds = new HashSet<long>();

        private volatile bool _finished;
        private DispatcherTimer _autoCloseTimer;

        public int DoneCount { get; private set; }
        public int FailedCount { get; private set; }
        public IEnumerable<CoqcsBatchErrorRow> Errors { get { return _errors; } }

        public FrmBatchProgress(string connectionString, Guid runId, int total, string kindLabel)
        {
            InitializeComponent();

            _connectionString = connectionString;
            _runId = runId;
            _total = total;

            Title = string.Format("Tiến trình chạy QCS {0}", kindLabel);
            lvErrors.ItemsSource = _errors;
            pbProgress.Maximum = total > 0 ? total : 1;
            lblStatus.Text = string.Format("Đang chạy QCS {0}... (0/{1})", kindLabel, total);

            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            RefreshProgress();
        }

        private void RefreshProgress()
        {
            try
            {
                int done = 0, failed = 0;
                string current = null;

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(
                        "SELECT ISNULL(SUM(CASE WHEN ket_thuc IS NOT NULL THEN 1 ELSE 0 END), 0), " +
                        "ISNULL(SUM(CASE WHEN ket_thuc IS NOT NULL AND thanh_cong = 0 THEN 1 ELSE 0 END), 0) " +
                        "FROM dbo.COQCS_BatchLog WHERE run_id = @run_id", connection))
                    {
                        command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier).Value = _runId;

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                done = reader.GetInt32(0);
                                failed = reader.GetInt32(1);
                            }
                        }
                    }

                    using (SqlCommand command = new SqlCommand(
                        "SELECT TOP (1) ma_vt, loai FROM dbo.COQCS_BatchLog " +
                        "WHERE run_id = @run_id AND ket_thuc IS NULL ORDER BY bat_dau DESC", connection))
                    {
                        command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier).Value = _runId;

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                                current = string.Format("{0} [{1}]", reader.GetString(0), reader.GetString(1));
                        }
                    }

                    using (SqlCommand command = new SqlCommand(
                        "SELECT id, ma_vt, loai, worker, loi FROM dbo.COQCS_BatchLog " +
                        "WHERE run_id = @run_id AND ket_thuc IS NOT NULL AND thanh_cong = 0 ORDER BY id", connection))
                    {
                        command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier).Value = _runId;

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                long id = reader.GetInt64(0);
                                if (_knownErrorIds.Contains(id))
                                    continue;

                                _knownErrorIds.Add(id);
                                _errors.Add(new CoqcsBatchErrorRow
                                {
                                    MaVt = reader.GetString(1),
                                    Loai = reader.GetString(2),
                                    Worker = reader.GetInt32(3),
                                    Loi = reader.IsDBNull(4) ? "" : reader.GetString(4)
                                });
                            }
                        }
                    }
                }

                DoneCount = done;
                FailedCount = failed;

                pbProgress.Value = Math.Min(done, _total > 0 ? _total : done);
                lblStatus.Text = string.Format("Đã chạy {0}/{1} - {2} lỗi{3}", done, _total, failed, _finished ? " - Hoàn tất" : "");
                lblDetail.Text = _finished ? "" : (current == null ? "" : "Đang xử lý: " + current);
            }
            catch
            {
                // Bỏ qua lỗi polling tạm thời (deadlock/timeout) - lần tick sau sẽ tự cập nhật lại.
            }
        }

        public void MarkFinished()
        {
            _finished = true;
            _timer.Stop();
            RefreshProgress();
            btnClose.IsEnabled = true;

            // Tự đóng sau 2 giây để StartUp.RunQcsBatchParallel (đang block ở ShowDialog())
            // tiếp tục và hiện thông báo tổng kết, thay vì bắt người dùng phải tự bấm "Đóng" -
            // vẫn chờ 2s để kịp nhìn thấy trạng thái "Hoàn tất" trước khi cửa sổ biến mất.
            lblDetail.Text = "Tự đóng sau 2 giây...";
            _autoCloseTimer = new DispatcherTimer();
            _autoCloseTimer.Interval = TimeSpan.FromSeconds(2);
            _autoCloseTimer.Tick += delegate
            {
                _autoCloseTimer.Stop();
                Close();
            };
            _autoCloseTimer.Start();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!_finished)
                e.Cancel = true;
        }
    }
}
