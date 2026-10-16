using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using SasDataLib;
using SasDefine;
using SasControls;
using SasFormBrowes;
using SasVoucherLib;

namespace QCCNCCOLD
{
    // Metadata is kept in memory, not persisted as extra CT50 columns.
    internal static class CriterionMetadata
    {
        internal static DataTable LoadCatalog()
        {
            SqlCommand command = new SqlCommand(
                "SELECT ma_tieu_chi, ten_tieu_chi, muc_tieu, thang_diem, trong_so, trang_thai_dg " +
                "FROM dmtieuchi WHERE status = '1' AND trang_thai_dg = '2' " +
                "AND (ma_dvcs = @ma_dvcs OR ISNULL(ma_dvcs, '') = '') ORDER BY ma_tieu_chi");
            command.Parameters.Add("@ma_dvcs", SqlDbType.Char).Value = StartupBase.SasObj.M_ma_dvcs.Trim();
            return DataProvider.FillCommand(StartupBase.SasObj, command).Tables[0];
        }

        private static void EnsureStringColumn(DataTable table, string name)
        {
            if (!table.Columns.Contains(name)) table.Columns.Add(name, typeof(string));
        }

        internal static void ApplyToVoucher(DataTable details, DataTable catalog)
        {
            EnsureStringColumn(details, "muc_tieu");
            EnsureStringColumn(details, "thang_diem_tc");
            EnsureStringColumn(details, "trong_so_tc");
            Dictionary<string, DataRow> lookup = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow criterion in catalog.Rows)
                lookup[Convert.ToString(criterion["ma_tieu_chi"]).Trim()] = criterion;
            foreach (DataRow detail in details.Rows)
            {
                if (detail.RowState == DataRowState.Deleted) continue;
                DataRow criterion;
                if (lookup.TryGetValue(Convert.ToString(detail["ma_tieu_chi"]).Trim(), out criterion))
                {
                    detail["muc_tieu"] = criterion["muc_tieu"];
                    detail["thang_diem_tc"] = criterion["thang_diem"];
                    detail["trong_so_tc"] = criterion["trong_so"];
                }
                else
                {
                    detail["muc_tieu"] = DBNull.Value;
                    detail["thang_diem_tc"] = DBNull.Value;
                    detail["trong_so_tc"] = DBNull.Value;
                }
            }
        }

        internal static void PrepareReport(DataSet report)
        {
            DataTable catalog = LoadCatalog();
            DataTable details = report.Tables[1];
            ApplyToVoucher(details, catalog);
            EnsureStringColumn(details, "trong_so");
            // CT50 uses numeric thang_diem; only the print COPY uses the textual scale.
            if (details.Columns.Contains("thang_diem") && details.Columns["thang_diem"].DataType != typeof(string))
            {
                if (details.Columns.Contains("thang_diem_so"))
                    throw new InvalidOperationException("Report already contains thang_diem_so.");
                details.Columns["thang_diem"].ColumnName = "thang_diem_so";
            }
            EnsureStringColumn(details, "thang_diem");
            foreach (DataRow row in details.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                row["trong_so"] = row["trong_so_tc"];
                row["thang_diem"] = row["thang_diem_tc"];
            }
            catalog = catalog.Copy();
            catalog.TableName = "tbTieuChi";
            if (report.Tables.Contains(catalog.TableName)) report.Tables.Remove(catalog.TableName);
            report.Tables.Add(catalog);
        }
    }
}
