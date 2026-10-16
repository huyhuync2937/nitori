using Infragistics.Windows.DataPresenter;
using Infragistics.Windows.Editors;
using Microsoft.SqlServer.Server;
using SasControls;
using SasErrorLib;
using SasFormBrowes;
using SasFormReport;
using SasLib;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace COQCS
{
    public class StartUp : StartupBase
    {
        public static DataSet dsReport = new DataSet();
        public static string hdTuNgay_ = "  -  -  ";
        public static string hdDenNgay_ = "  -  -  ";
        public static string ctTuNgay_ = "  -  -  ";
        public static string ctDenNgay_ = "  -  -  ";
        public static string tableList = "v_COQCS";
        public static int kindStyleReport = -1;
        public static string g_strCondition = string.Empty;
        public static string g_strFilter = string.Empty;
        public static string g_ps_ck = "*";
        public static SasFormBrowes.FormBrowse oBrowse;
        public static SasFormBrowes.FormBrowse oBrowse_ct;
        public static FrmLoc _frmLoc;
        public static DataRow commandInfo;
        public static DateTime M_NGAY_KY1;
        public static DateTime M_NGAY_CT1;
        public static string M_MA_DVCS;
        public static string M_IP_TIEN;
        public static string M_IP_TIEN_NT;
        public static string M_IP_SL;
        public static string M_ma_nt0;
        public static object g_hdTuNg;
        public static object g_hdDenNg;
        public static object g_ctTuNg;
        public static object g_ctDenNg;
        public static object g_loaiNvl;
        public static object g_loai;
        public static object g_pic;

        public static string[] CanChangeValueFields;
        public static int m_user_id = 0;

        public override void Run()
        {
            StartupBase.Namespace = "COQCS";
            try
            {
                StartUp.CanChangeValueFields = "ngay_bd,ma_may,ma_ca,ngay_kt,ma_ca_kt,sl_cai_lenh,sl_kg_lenh,id,ma_khuon,dvt".Split(',');

                StartUp.M_ma_nt0 = StartupBase.SasObj.GetOption("M_MA_NT0").ToString();
                StartUp.M_IP_TIEN = StartupBase.SasObj.GetOption("M_IP_TIEN").ToString();
                StartUp.M_IP_SL = StartupBase.SasObj.GetOption("M_IP_SL").ToString();
                StartUp.M_IP_TIEN_NT = StartupBase.SasObj.GetOption("M_IP_TIEN_NT").ToString();
                StartUp.M_NGAY_KY1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_KY1");
                StartUp.M_NGAY_CT1 = (DateTime)StartupBase.SasObj.GetSysvar("M_NGAY_CT1");
                StartUp.M_MA_DVCS = StartupBase.SasObj.DmdvcsInfo.Rows[0][0].ToString();
                StartUp.commandInfo = SysFunc.GetCommandInfo(StartupBase.SasObj, StartupBase.Menu_Id);
                if (StartUp.commandInfo == null)
                    return;
                StartUp._frmLoc = new FrmLoc();
                StartUp._frmLoc.Title = SysFunc.Cat_Dau(StartupBase.M_LAN.Equals("V") ? StartUp.commandInfo["bar"].ToString() : StartUp.commandInfo["bar2"].ToString());
                StartUp._frmLoc.ShowDialog();
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }

        public static void CallGridReport( bool isFirstLoad,object hdTuNg,object hdDenNg,object txtPic,object txtStatus)
        {
            StartUp.g_hdTuNg = hdTuNg;
            StartUp.g_hdDenNg = hdDenNg;
            StartUp.g_pic = txtPic;
            StartUp.g_loai = txtStatus;


            if (isFirstLoad)
            {
                string loai = string.IsNullOrEmpty(txtStatus.ToString()) ? "" : txtStatus.ToString().Trim();

                // KHÔNG gọi qua COQCS_main/COQCS_doc_main nữa, cho cả dọc và ngang. Phân loại
                // 1 lần (persistent), rồi chạy song song nhiều worker gọi thẳng batch-proc tương
                // ứng cho từng nhóm loai_mode - dọc dùng đúng cơ chế như ngang theo yêu cầu.
                string hdTuNgStr = string.IsNullOrEmpty(hdTuNg.ToString()) ? "" : string.Format("{0:yyyyMMdd}", (object)(DateTime)hdTuNg);
                string hdDenNgStr = string.IsNullOrEmpty(hdDenNg.ToString()) ? "" : string.Format("{0:yyyyMMdd}", (object)(DateTime)hdDenNg);

                if (string.IsNullOrEmpty(hdTuNgStr) || string.IsNullOrEmpty(hdDenNgStr))
                {
                    int num = (int)ExMessageBox.Show(2304, StartupBase.SasObj, "Phải nhập đủ từ ngày và đến ngày khi chạy QCS!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                    return;
                }

                string pic = string.IsNullOrEmpty(txtPic.ToString()) ? "" : txtPic.ToString().Trim();

                if (loai == "2")
                    StartUp.RunQcsBatchParallel(StartUp.NgangLoaiModes, "dbo.COQCS_ngang_batch", "ngang", hdTuNgStr, hdDenNgStr, pic);
                else
                    StartUp.RunQcsBatchParallel(StartUp.DocLoaiModes, "dbo.COQCS_doc_batch", "dọc", hdTuNgStr, hdDenNgStr, pic);

                //StartUp.oBrowse.SetRowColorByTag("sx2", "0", Colors.Black, true);
                //StartUp.oBrowse.SetRowColorByTag("sx3", "1", Colors.Red, false);
            }
            //else
            //{
            //    StartUp.dsReport.Tables.Clear();
            //    StartUp.dsReport = StartupBase.SasObj.ExcuteReader(StartUp.cmd);
            //    DataTable dataTable = StartUp.dsReport.Tables[0].Copy();
            //    StartUp.dsReport.Tables[0].TableName = "tbDetail";
            //    StartUp.dsReport.Tables[0].TableName = "tbtong";
            //    dataTable.TableName = "tbDetail";
            //    StartUp.dsReport.Tables.Add(StartUp.CreateTableInfo().Copy());
            //    StartUp.oBrowse.frmBrw.oBrowse.DataSource = (IEnumerable)dataTable.DefaultView;
            //    StartUp.oBrowse.frmBrw.oBrowse.FieldLayouts[0].SummaryDefinitions.Clear();
            //    StartUp.oBrowse.UpdateSumaryFields();
            //}
           
            StartUp._frmLoc.Close();
        }

        // 6 nhóm loai_mode xử lý "ngang" bằng dbo.COQCS_ngang_batch/COQCS_ngang_all.
        // 'FC'/'NHAMAY_FC' là 2 loai_mode mới do COQCS_PhanLoai sinh ra khi vật tư vừa có BOM vừa có Forecast
        // (ksxsanphamf) - ưu tiên hơn 'NGANG_BOM'/'NHAMAY_B', phải chạy qua COQCS_ngang_all @mode='FC'/'NHAMAY_FC'.
        private static readonly string[] NgangLoaiModes = new string[] { "NGANG", "NGANG_BOM", "FC", "NHAMAY_B", "NHAMAY_KB", "NHAMAY_FC" };
        // 3 nhóm loai_mode xử lý "dọc" bằng dbo.COQCS_doc_batch/COQCS_doc_all (do COQCS_PhanLoai sinh ra
        // khi dinh_dang_qcs = 1) - cùng cơ chế song song như NgangLoaiModes ở trên.
        private static readonly string[] DocLoaiModes = new string[] { "DOC", "DOC_BOM", "DOC_FC" };
        // = 1 (không chạy song song giữa các worker trong 1 nhóm loai_mode). Đã từng tăng thử
        // lên 2 rồi 3 để test - trong lúc đó đã fix 2 bug có sẵn lộ ra do tăng concurrency:
        // (1) dbo.SetColumnDefaultValue (gọi qua INCD1_realtime_input) tạo constraint trên bảng
        // tạm bằng TÊN CỐ ĐỊNH không phân biệt session -> đụng tên nhau (1750/2714/1205) - đã
        // sửa thêm @@SPID vào tên constraint (xem Sql/Fix_SetColumnDefaultValue_SPID.sql).
        // (2) @retryDelay trong nhánh retry của COQCS_ngang_batch/COQCS_doc_batch khai báo kiểu
        // TIME - WAITFOR DELAY không nhận kiểu TIME - đã sửa thành DATETIME trực tiếp trên
        // server. Cả 2 fix trên vẫn còn hiệu lực trên DB dù đưa worker count về lại 1 ở đây -
        // không cần rollback, vì không gây hại gì khi chạy tuần tự (worker=1 hiếm khi chạm retry
        // path). Áp dụng cho cả dọc và ngang.
        private const int BATCH_GROUP_WORKER_COUNT = 1;

        // Dùng chung cho cả "ngang" (loaiModes=NgangLoaiModes, storeProc=COQCS_ngang_batch) và
        // "dọc" (loaiModes=DocLoaiModes, storeProc=COQCS_doc_batch) - 2 batch-proc có cùng chữ ký
        // tham số (xem CoqcsBatchRunner), chỉ khác @loai hợp lệ và proc con (ngang_all/doc_all).
        private static void RunQcsBatchParallel(string[] loaiModes, string storeProc, string kindLabel, string hdTuNgStr, string hdDenNgStr, string pic)
        {
            Guid runId = StartUp.ClassifyQcs(pic);

            // Lấy trước toàn bộ danh sách ma_target theo từng loai_mode để biết tổng số mã
            // cần chạy (hiển thị lên progress bar) trước khi bắt đầu chạy song song.
            Dictionary<string, List<string>> targetsByMode = new Dictionary<string, List<string>>();
            int total = 0;
            foreach (string loaiModeLoop in loaiModes)
            {
                List<string> maTargets = StartUp.GetMaTargetByLoaiMode(runId, loaiModeLoop);
                if (maTargets.Count == 0)
                    continue;

                targetsByMode[loaiModeLoop] = maTargets;
                total += maTargets.Count;
            }

            if (total == 0)
            {
                int numEmpty = (int)ExMessageBox.Show(2303, StartupBase.SasObj, string.Format("Không có mã vật tư nào cần chạy QCS {0}.", kindLabel), "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return;
            }

            List<CoqcsWorkerResult> allResults = new List<CoqcsWorkerResult>();

            FrmBatchProgress progressForm = new FrmBatchProgress(StartupBase.SasObj.M_ConnectString, runId, total, kindLabel);

            // Chạy toàn bộ việc điều phối trên 1 thread nền, để UI thread rảnh cho
            // FrmBatchProgress bơm DispatcherTimer cập nhật tiến độ real-time (poll
            // dbo.COQCS_BatchLog) trong lúc ShowDialog() đang block bên dưới.
            //
            // Chạy TUẦN TỰ từng nhóm loai_mode (không spawn hết N nhóm x nhiều worker song song
            // cùng lúc) - mỗi nhóm vẫn có BATCH_GROUP_WORKER_COUNT worker chạy song song bên trong
            // CoqcsBatchRunner.Run. Nhiều session đồng thời cùng gọi các proc dùng chung
            // (SetColumnDefaultValue, CopyColumnFromTable, INCD1_realtime_input...) gây tranh chấp
            // metadata tempdb (deadlock 1205, "Could not create constraint" 1750) dù batch-proc đã
            // có retry - chạy tuần tự từng nhóm để giảm hẳn tần suất tranh chấp.
            Thread orchestratorThread = new Thread(delegate ()
            {
                foreach (KeyValuePair<string, List<string>> kv in targetsByMode)
                {
                    string loaiMode = kv.Key;
                    List<string> maTargets = kv.Value;

                    CoqcsBatchRunner runner = new CoqcsBatchRunner(StartupBase.SasObj.M_ConnectString, storeProc);
                    List<CoqcsWorkerResult> groupResults = runner.Run(maTargets, hdTuNgStr, hdDenNgStr, loaiMode, BATCH_GROUP_WORKER_COUNT, runId);

                    allResults.AddRange(groupResults);
                }

                progressForm.Dispatcher.Invoke(new Action(delegate ()
                {
                    progressForm.MarkFinished();
                }));
            });
            orchestratorThread.IsBackground = true;
            orchestratorThread.Start();

            // Không gán Owner: FrmLoc đã bị Hide() ở ConfirmGridView_OnOk trước khi gọi tới đây,
            // nếu CenterOwner/Owner trỏ vào 1 window đang Hidden thì WPF có thể định vị progressForm
            // ra ngoài vùng hiển thị (kích thước owner coi như 0) -> form coi như "mất tích".
            progressForm.ShowDialog();

            // Lỗi cấp worker (cả proc batch throw ra C#, vd. mất cả lô trước khi kịp ghi dòng log
            // nào) - khác với lỗi cấp từng mã (progressForm.FailedCount, đọc từ COQCS_BatchLog,
            // đã hiển thị chi tiết trong lưới của progressForm lúc chạy).
            int soWorkerLoi = 0;
            foreach (CoqcsWorkerResult r in allResults)
                if (!r.Success)
                    soWorkerLoi++;

            string msg = string.Format("Đã chạy xong QCS {0} - {1}/{2} mã, {3} mã lỗi.", kindLabel,
                progressForm.DoneCount - progressForm.FailedCount, total, progressForm.FailedCount);

            int num = (int)ExMessageBox.Show(2303, StartupBase.SasObj, msg, "", MessageBoxButton.OK, MessageBoxImage.Asterisk);

            // FrmBatchProgress tự đóng sau 2s nên không kịp đọc lưới lỗi bên trong nó - mở riêng
            // FrmBatchErrors (không tự đóng) để xem/cuộn chi tiết từng mã lỗi + từng nhóm lỗi
            // ngay từ đầu, thay vì dồn hết vào 1 khối text trong message box như trước.
            if (progressForm.FailedCount > 0 || soWorkerLoi > 0)
            {
                List<CoqcsGroupErrorRow> groupErrors = new List<CoqcsGroupErrorRow>();
                foreach (CoqcsWorkerResult r in allResults)
                    if (!r.Success)
                        groupErrors.Add(new CoqcsGroupErrorRow { LoaiMode = r.LoaiMode, Worker = r.Worker, Error = r.Error });

                FrmBatchErrors errForm = new FrmBatchErrors(msg, progressForm.Errors, groupErrors);
                errForm.ShowDialog();
            }
        }

        // Tách khối "1. PHÂN LOẠI" của dbo.COQCS_main, ghi persistent vào dbo.COQCS_PhanLoaiKq
        // (thay vì #qcs_ds_part_no), trả về run_id để app đọc ra danh sách ma_target theo loai_mode.
        // Dùng chung cho cả dọc và ngang - COQCS_PhanLoai sinh ra loai_mode của cả 2 phía cùng lúc.
        private static Guid ClassifyQcs(string pic)
        {
            using (SqlConnection connection = new SqlConnection(StartupBase.SasObj.M_ConnectString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand("dbo.COQCS_PhanLoai", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 0;

                    command.Parameters.Add("@pic", SqlDbType.VarChar, 50).Value = string.IsNullOrEmpty(pic) ? (object)DBNull.Value : (object)pic;

                    SqlParameter runIdParam = command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier);
                    runIdParam.Direction = ParameterDirection.Output;

                    command.ExecuteNonQuery();

                    return (Guid)runIdParam.Value;
                }
            }
        }

        private static List<string> GetMaTargetByLoaiMode(Guid runId, string loaiMode)
        {
            List<string> list = new List<string>();

            using (SqlConnection connection = new SqlConnection(StartupBase.SasObj.M_ConnectString))
            {
                connection.Open();
                // DISTINCT: nhiều part_no khác nhau có thể cùng map về 1 code_mua (NHAMAY_*),
                // nên ma_target có thể bị trùng - nếu để trùng, @ds truyền vào COQCS_ngang_batch
                // sẽ có ma_vt lặp, làm CREATE UNIQUE CLUSTERED INDEX IX_DS thất bại ngay từ đầu
                // (lỗi 1750, mất trắng cả worker đó trước khi log được dòng nào).
                using (SqlCommand command = new SqlCommand("SELECT DISTINCT ma_target FROM dbo.COQCS_PhanLoaiKq WHERE run_id = @run_id AND loai_mode = @loai_mode", connection))
                {
                    command.Parameters.Add("@run_id", SqlDbType.UniqueIdentifier).Value = runId;
                    command.Parameters.Add("@loai_mode", SqlDbType.VarChar, 20).Value = loaiMode;

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                            list.Add(reader.GetString(0));
                    }
                }
            }

            return list;
        }

        public static void FrmBrw_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.None)
            {
                switch (e.Key)
                {
                    case Key.Space:
                        SelectOne();
                        break;
                   
                }
            }
            if (Keyboard.Modifiers != ModifierKeys.Control)
                return;
            switch (e.Key)
            {
                case Key.A:
                  SelectAll(true);
                    break;
                case Key.U:
                   SelectAll(false);
                    break;
            }
        }
        private static void SelectAll(bool tag)
        {
            if (StartUp.oBrowse.DataGrid.Records.Count == 0)
                return;
            foreach (DataRecord filteredInDataRecord in StartUp.oBrowse.DataGrid.RecordManager.GetFilteredInDataRecords())
                (filteredInDataRecord.DataItem as DataRowView)["chon"] = (object)tag;
        }

        private static void SelectOne()
        {
            if (StartUp.oBrowse.DataGrid.ActiveRecord == null)
                return;
            DataRecord activeRecord = StartUp.oBrowse.DataGrid.ActiveRecord as DataRecord;
            if (activeRecord == null || activeRecord.DataItem == null || activeRecord is FilterRecord)
                return;
            DataRowView dataItem = activeRecord.DataItem as DataRowView;
            bool flag = (bool)dataItem["chon"];
            dataItem["chon"] = (object)!flag;
        }


    
        public static void oBrowse_F7(object sender, EventArgs e)
        {
            ReportManager reportManager = new ReportManager(StartupBase.SasObj, StartUp.commandInfo["rep_file"].ToString(), StartUp.kindStyleReport);
            SysFunc.DSCopyWithFilter(StartUp.oBrowse.frmBrw.oBrowse, ref StartUp.dsReport, "tbDetail");
            reportManager.Preview(StartUp.dsReport);
            SysFunc.ResetFilter(ref StartUp.dsReport, "tbDetail");
        }

        public static void oBrowse_CTRL_R(object sender, EventArgs e)
        {
            try
            {
          
                bool kindReport = true;
            
                StartUp.CallGridReport(false, StartUp.g_hdTuNg, StartUp.g_hdDenNg, StartUp.g_pic, StartUp.g_loai);
            }
            catch (Exception ex)
            {
                ErrorLog.CatchMessage(ex);
            }
        }


        public static string fieldShow(int kindReport, int isDetail)
        {
            string str = string.Empty;
            string[] strArray1 = new string[2];
            switch (StartupBase.SasObj.GetOption("M_LAN").ToString().ToUpper())
            {
                case "V":
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray2 = StartUp.commandInfo["Vbrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray2[1] : strArray2[0];
                            break;
                        case 2:
                            string[] strArray3 = StartUp.commandInfo["Vbrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray3[1] : strArray3[0];
                            break;
                    }
                    break;
                default:
                    switch (kindReport)
                    {
                        case 1:
                            string[] strArray4 = StartUp.commandInfo["Ebrowse1"].ToString().Split('|');
                            str = isDetail != 0 ? strArray4[1] : strArray4[0];
                            break;
                        case 2:
                            string[] strArray5 = StartUp.commandInfo["Ebrowse2"].ToString().Split('|');
                            str = isDetail != 0 ? strArray5[1] : strArray5[0];
                            break;
                    }
                    break;
            }
            return str;
        }

    
    }
}
