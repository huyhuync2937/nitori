using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;

namespace COQCS
{
    public sealed class CoqcsWorkerResult
    {
        public string LoaiMode { get; private set; }
        public int Worker { get; private set; }
        public bool Success { get; private set; }
        public TimeSpan Elapsed { get; private set; }
        public string Error { get; private set; }

        public CoqcsWorkerResult(string loaiMode, int worker, bool success, TimeSpan elapsed, string error)
        {
            LoaiMode = loaiMode;
            Worker = worker;
            Success = success;
            Elapsed = elapsed;
            Error = error;
        }
    }

    // Chạy N worker song song (mỗi worker 1 SqlConnection riêng) cho CÙNG 1 @loai.
    // Mỗi worker gọi _storeProc (dbo.COQCS_ngang_batch hoặc dbo.COQCS_doc_batch - 2 proc có
    // cùng chữ ký tham số) với cùng @ds (toàn bộ danh sách ma_target của loai_mode đó) nhưng
    // @worker khác nhau - proc tự chia việc theo công thức (rn - 1) % @soWorker = @worker
    // đã có sẵn, nên driver không cần tự cắt danh sách.
    public sealed class CoqcsBatchRunner
    {
        private readonly string _connectionString;
        private readonly string _storeProc;

        public CoqcsBatchRunner(string connectionString, string storeProc)
        {
            _connectionString = connectionString;
            _storeProc = storeProc;
        }

        public List<CoqcsWorkerResult> Run(
            List<string> maVtList,
            string hdTuNg,
            string hdDenNg,
            string loai,
            int workerCount,
            Guid runId)
        {
            List<CoqcsWorkerResult> results = new List<CoqcsWorkerResult>();
            object lockObj = new object();
            List<Thread> threads = new List<Thread>();

            for (int w = 0; w < workerCount; w++)
            {
                int worker = w; // capture rõ ràng theo giá trị, tránh lỗi closure chia sẻ biến loop
                Thread thread = new Thread(delegate ()
                {
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    try
                    {
                        RunWorker(maVtList, hdTuNg, hdDenNg, loai, worker, workerCount, runId);

                        lock (lockObj)
                            results.Add(new CoqcsWorkerResult(loai, worker, true, stopwatch.Elapsed, null));
                    }
                    catch (Exception ex)
                    {
                        lock (lockObj)
                            results.Add(new CoqcsWorkerResult(loai, worker, false, stopwatch.Elapsed, ex.Message));
                    }
                });
                thread.IsBackground = true;
                threads.Add(thread);
            }

            foreach (Thread thread in threads)
                thread.Start();
            foreach (Thread thread in threads)
                thread.Join();

            return results;
        }

        private void RunWorker(
            List<string> maVtList,
            string hdTuNg,
            string hdDenNg,
            string loai,
            int worker,
            int workerCount,
            Guid runId)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(_storeProc, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 0;

                    SqlParameter ds = command.Parameters.Add("@ds", SqlDbType.Structured);
                    ds.TypeName = "dbo.udt_MaVt";
                    ds.Value = BuildMaVtTable(maVtList);

                    command.Parameters.Add("@hdTuNg", SqlDbType.VarChar, 8).Value = hdTuNg;
                    command.Parameters.Add("@hdDenNg", SqlDbType.VarChar, 8).Value = hdDenNg;
                    command.Parameters.Add("@loai", SqlDbType.VarChar, 20).Value = loai;
                    command.Parameters.Add("@worker", SqlDbType.Int).Value = worker;
                    command.Parameters.Add("@soWorker", SqlDbType.Int).Value = workerCount;
                    command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier).Value = runId;

                    command.ExecuteNonQuery();
                }
            }
        }

        private static DataTable BuildMaVtTable(List<string> maVtList)
        {
            DataTable table = new DataTable();
            table.Columns.Add("ma_vt", typeof(string));

            foreach (string maVt in maVtList)
                table.Rows.Add(maVt);

            return table;
        }
    }
}
