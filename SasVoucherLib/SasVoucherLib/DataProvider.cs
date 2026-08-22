using SasDataLib;
using SasErrorLib;
using SasLib;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using System;


namespace SasVoucherLib
{
    /// <summary>
    /// Lớp tiện ích dùng cho các thao tác dữ liệu trên các phiếu.
    /// </summary>
    public class DataProvider
    {
        /// <summary>Xóa các dòng thỏa điều kiện truyền vào.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="SqlTableName">Tên table chứa các rows sẽ xóa.</param>
        /// <param name="DeleteKey">Điều kiện xác định các rows sẽ bị xóa.</param>
        public static void DeleteRow(SasObject SasObj, string SqlTableName, string DeleteKey)
        {
            SqlCommand sqlcmd = new SqlCommand();
            string str = "Delete from " + SqlTableName + " where " + DeleteKey;
            sqlcmd.CommandText = str;
            SasObj.ExcuteNonQuery(sqlcmd);
        }

        /// <summary>Thêm nhiều dòng vào dữ liệu.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="SqlTableName">Tên table chứa các rows sẽ thêm.</param>
        /// <param name="LocalTable">Table chức các rows sẽ thêm.</param>
        /// <returns>Kết quả thực hiện.(True- thành công,False- thất bại).</returns>
        public static bool UpLoadDataTable(SasObject SasObj, string SqlTableName, DataTable LocalTable)
        {
            SqlCommand sqlCommand = new SqlCommand();
            string str = new string(' ', 32000);
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(SasObj, SqlTableName);
            string strFieldMapping = "";
            foreach (DataColumn column in (InternalDataCollectionBase)LocalTable.Columns)
            {
                if (((IEnumerable<DataRow>)sqlTableFieldList.Select("name='" + column.ColumnName + "'")).Count<DataRow>() != 0)
                    strFieldMapping = strFieldMapping + (string.IsNullOrEmpty(strFieldMapping) ? "" : ";") + string.Format("{0}={1}", (object)column.ColumnName, (object)column.ColumnName);
            }
            return SasObj.ExcuteBulkCopy(SqlTableName, LocalTable, strFieldMapping) != -1;
        }

        public static bool UpdateCtTable(
          SasObject SasObj,
          string SqlTableName,
          DataTable LocalTable,
          string stt_rec)
        {
            SqlCommand sqlCommand = new SqlCommand();
            string str = new string(' ', 32000);
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(SasObj, SqlTableName);
            string strFieldMapping = "";
            foreach (DataColumn column in (InternalDataCollectionBase)LocalTable.Columns)
            {
                if (((IEnumerable<DataRow>)sqlTableFieldList.Select("name='" + column.ColumnName + "'")).Count<DataRow>() != 0)
                    strFieldMapping = strFieldMapping + (string.IsNullOrEmpty(strFieldMapping) ? "" : ";") + string.Format("{0}={1}", (object)column.ColumnName, (object)column.ColumnName);
            }
            return  SasObj.UpdateCtTable(SqlTableName, LocalTable, strFieldMapping, stt_rec) != -1;
        }

        public static int UpdateCtTable(string sqlTable, DataTable tbCopy, string strFieldMapping, string Stt_rec, SasObject SasObj)
        {
            int num1 = 0;
            int index = SasObj.AvailableConnectionNo();
            SasObj.QConnectStatus[index] = 1;
            SqlTransaction sqlTransaction = (SqlTransaction)null;
            SqlBulkCopy BulkCmd = (SqlBulkCopy)null;
            try
            {
                sqlTransaction = SasObj.QConnection[index].BeginTransaction("UpdateCtTransaction");
                BulkCmd = new SqlBulkCopy(SasObj.QConnection[index], SqlBulkCopyOptions.Default, sqlTransaction);
                SqlCommand sqlCommand = new SqlCommand(string.Format("delete {0} where stt_rec=@stt_rec", sqlTable), SasObj.QConnection[index], sqlTransaction);
                sqlCommand.Parameters.Add("@stt_rec", SqlDbType.Char).Value = Stt_rec;
                string[] strArray1 = strFieldMapping.Split(';');
                BulkCmd.BatchSize = 500;
                BulkCmd.DestinationTableName = sqlTable;
                BulkCmd.BulkCopyTimeout = SasObj.nTimeOutExecute;
                foreach (string str in strArray1)
                {
                    char[] chArray = new char[1] { '=' };
                    string[] strArray2 = str.Split(chArray);
                    BulkCmd.ColumnMappings.Add(strArray2[1].Trim(), strArray2[0].Trim());
                }
                if (SasObj.isEnableSqlLog)
                    SasObj.SqlLog += string.Format("{0}{1}", SasObj.SqlLog != "" ? Environment.NewLine : "", SasObj.SqlString(tbCopy, BulkCmd));
                if (SasObj.QConnection[index].State == ConnectionState.Closed)
                    SasObj.QConnection[index].Open();
                sqlCommand.ExecuteNonQuery();
                BulkCmd.WriteToServer(tbCopy);
                sqlTransaction.Commit();
                SasObj.QConnection[index].Close();
                SasObj.QConnectStatus[index] = 0;
            }
            catch (SqlException ex1)
            {
                if (DataProvider.isConnectionCorrupt(ex1))
                {
                    int num2 = (int)MessageBox.Show(string.Format("SQL-EX - [{0}, {1}], loi duong truyen {2}", ex1.Number, ex1.ErrorCode, ex1.Message));
                    bool flag = false;
                    while (!flag)
                    {
                        Thread.Sleep(500);
                        try
                        {
                            if (SasObj.QConnection[index] != null)
                                SasObj.QConnection[index].Dispose();
                            SasObj.QConnection[index] = new SqlConnection(SasObj.M_ConnectString);
                            SasObj.QConnection[index].Open();
                            if (BulkCmd != null)
                            {
                                BulkCmd.Close();
                                BulkCmd = (SqlBulkCopy)null;
                            }
                            SqlCommand sqlCommand = new SqlCommand(string.Format("delete {0} where stt_rec=@stt_rec", sqlTable), SasObj.QConnection[index], sqlTransaction);
                            sqlCommand.Parameters.Clear();
                            sqlCommand.Parameters.Add("@stt_rec", SqlDbType.Char).Value = Stt_rec;
                            BulkCmd = new SqlBulkCopy(SasObj.QConnection[index], SqlBulkCopyOptions.Default, sqlTransaction);
                            BulkCmd.DestinationTableName = sqlTable;
                            BulkCmd.BatchSize = 500;
                            BulkCmd.BulkCopyTimeout = SasObj.nTimeOutExecute;
                            string str1 = strFieldMapping;
                            char[] chArray1 = new char[1] { ';' };
                            foreach (string str2 in str1.Split(chArray1))
                            {
                                char[] chArray2 = new char[1] { '=' };
                                string[] strArray = str2.Split(chArray2);
                                BulkCmd.ColumnMappings.Add(strArray[1], strArray[0]);
                            }
                            if (SasObj.QConnection[index].State == ConnectionState.Closed)
                                SasObj.QConnection[index].Open();
                            sqlCommand.ExecuteNonQuery();
                            BulkCmd.WriteToServer(tbCopy);
                            sqlTransaction.Commit();
                            SasObj.QConnection[index].Close();
                            SasObj.QConnectStatus[index] = 0;
                            flag = true;
                        }
                        catch (SqlException ex2)
                        {
                        }
                        catch (Exception ex2)
                        {
                        }
                    }
                    return num1;
                }
                int num3 = (int)MessageBox.Show(string.Format("SQL-EX - [{0}, {1}], loi chung chung {2}", ex1.Number, ex1.ErrorCode, ex1.Message));
                sqlTransaction.Rollback();
                SasObj.QConnection[index].Close();
                if (SasObj.QConnection[index] != null)
                    SasObj.QConnection[index].Dispose();
                SasObj.QConnection[index] = new SqlConnection(SasObj.M_ConnectString);
                SasObj.QConnection[index].Open();
                SasObj.QConnectStatus[index] = 0;
                num1 = -1;
                int num4 = (int)MessageBox.Show(ex1.Message);
            }
            catch (Exception ex)
            {
                int num2 = (int)MessageBox.Show(string.Format("EX-EX {0}-{1}{2}", ex.Message, Environment.NewLine, strFieldMapping));
                sqlTransaction.Rollback();
                SasObj.QConnection[index].Close();
                if (SasObj.QConnection[index] != null)
                    SasObj.QConnection[index].Dispose();
                SasObj.QConnection[index] = new SqlConnection(SasObj.M_ConnectString);
                SasObj.QConnection[index].Open();
                SasObj.QConnectStatus[index] = 0;
                num1 = -1;
                int num3 = (int)MessageBox.Show(ex.Message);
            }
            return num1;
        }

        public static bool isConnectionCorrupt(SqlException sqlex)
        {
            int number = sqlex.Number;
            return number <= 64 ? number == 53 || number == 64 : number == 121 || number == 10051 || number == 10054;
        }
        /// <summary>Update các dòng trong dữ liệu.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="SqlTableName">Tên table chứa các rows sẽ update.</param>
        /// <param name="SqlPriColumn">Danh sách các cột khóa chính.</param>
        /// <param name="LocalTable">Table chức các rows sẽ sửa.</param>
        /// <param name="DefaultNonUpdatedFields">Danh sách các cột không update.</param>
        /// <returns>Kết quả thực hiện.(True- thành công,False- thất bại).</returns>
        public static bool UpdateDataTable(
          SasObject SasObj,
          string SqlTableName,
          string SqlPriColumn,
          DataTable LocalTable,
          string DefaultNonUpdatedFields)
        {
            string str1 = DefaultNonUpdatedFields;
            SqlCommand sqlcmd = new SqlCommand();
            string str2 = "Select 1 from " + SqlTableName;
            if (string.IsNullOrEmpty(SqlPriColumn))
                return false;
            string[] strArray1 = SqlPriColumn.Split(';');
            DataTable sqlTableFieldList = ListFunc.GetSqlTableFieldList(SasObj, SqlTableName);
            foreach (DataColumn column in (InternalDataCollectionBase)LocalTable.Columns)
            {
                if (((IEnumerable<DataRow>)sqlTableFieldList.Select("name='" + column.ColumnName + "'")).Count<DataRow>() == 0)
                    str1 = str1 + (string.IsNullOrEmpty(str1) ? "" : ";") + column.ColumnName;
            }
            string str3 = "update " + SqlTableName + " set ";
            int num = 0;
            string[] strArray2 = str1.ToUpper().Split(';');
            foreach (DataColumn column in (InternalDataCollectionBase)LocalTable.Columns)
            {
                if (!((IEnumerable<string>)strArray2).Contains<string>(column.ColumnName.ToUpper()))
                {
                    str3 = str3 + (num == 0 ? "" : ",") + column.ColumnName + " = @" + column.ColumnName;
                    DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + column.ColumnName + "'");
                    sqlcmd.Parameters.Add("@" + column.ColumnName, ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString()));
                    ++num;
                }
            }
            string str4 = str3 + " where ";
            for (int index = 0; index < strArray1.Length; ++index)
            {
                str4 = str4 + (index == 0 ? "" : " and ") + strArray1[index] + " = @" + strArray1[index];
                DataRow[] dataRowArray = sqlTableFieldList.Select("name='" + strArray1[index] + "'");
                sqlcmd.Parameters.Add("@" + strArray1[index], ListFunc.GetSqlDBType(dataRowArray[0]["datatype"].ToString()));
            }
            foreach (DataRow row in (InternalDataCollectionBase)LocalTable.Rows)
            {
                foreach (SqlParameter parameter in (DbParameterCollection)sqlcmd.Parameters)
                {
                    string index = parameter.ParameterName.Substring(1);
                    parameter.Value = row[index];
                }
                sqlcmd.CommandText = str4;                
                try
                {
                    SasObj.ExcuteNonQuery(sqlcmd);
                    //List<SqlParameter> sqlParameter1 = new List<SqlParameter>();
                    //foreach (SqlParameter para in sqlcmd.Parameters)
                    //{
                    //    sqlParameter1.Add(para);
                    //}
                    //SqlGenerator gn = new SqlGenerator();
                    //string str = gn.CreateExecutableSqlStatement("UpdateCommandInfo", CommandType.Text, sqlParameter1.ToArray());
                    //System.Windows.Forms.Clipboard.SetText(str);
                }
                catch (SqlException ex)
                {
                    ErrorLog.CatchMessage(ex);
                    return false;
                }
            }
            return true;
        }

        /// <summary>Tạo mới stt_rec cho một phiếu.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="cMa_ct">Mã chứng từ của phiếu.</param>
        /// <param name="Ws_Id">Mã trạm của phiếu.</param>
        /// <returns>stt_rec được sinh mới.</returns>
        public static string NewTrans(SasObject SasObj, string cMa_ct, string Ws_Id)
        {
            string str = "";
            string format = "exec {0} @Ma_ct, @Ws_Id";
            SqlCommand sqlcmd = new SqlCommand(StartUpTrans.Process_Store == null || StartUpTrans.Process_Store.Length <= 9 ? string.Format(format, (object)"GetSttRec") : string.Format(format, (object)StartUpTrans.Process_Store[9]));
            sqlcmd.Parameters.Add("@Ma_ct", SqlDbType.Char).Value = (object)cMa_ct;
            sqlcmd.Parameters.Add("@Ws_Id", SqlDbType.Char).Value = (object)Ws_Id;
            DataTable dataTable = new DataTable();
            try
            {
                dataTable = SasObj.ExcuteReader(sqlcmd).Tables[0];
            }
            catch (SqlException ex)
            {
                ErrorLog.CatchMessage(ex);
            }
            if (dataTable.Rows.Count > 0)
                str = dataTable.Rows[0]["Stt_rec"].ToString();
            return str;
        }

        /// <summary>Phương thức lấy dữ liệu.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="cmd">SqlCommand dùng để lấy dữ liệu.</param>
        /// <returns>Dataset chứa kết quả lấy dữ liệu.</returns>
        public static DataSet FillCommand(SasObject SasObj, SqlCommand cmd)
        {
            return (DataSet)SasObj.ExcuteReader(cmd);
        }

        /// <summary>Tạo mới stt_rec cho một phiếu.</summary>
        /// <param name="SasObj">SasObject dùng để xử lý dữ liệu.</param>
        /// <param name="cMa_ct">Mã chứng từ của phiếu.</param>
        /// <param name="Ws_Id">Mã trạm của phiếu.</param>
        /// <returns>stt_rec được sinh mới.</returns>
        public static string NewTrans2(SasObject SasObj, string cMa_ct, string Ws_Id)
        {
            string str = "";
            str = Ws_Id + System.Guid.NewGuid().ToString().Replace("-", "").ToUpper() +cMa_ct;
            return str.Substring(0,36);
        }
    }
}
