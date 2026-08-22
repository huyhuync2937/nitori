using SasControls;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Linq;

namespace SasIeCt
{
    public class C_GetDataExcel
    {
        public static string StrBrowse = "";

        public static string StrBrowseFieldNull = "";

        public static string TCNV3String = "ÊÈèÉÌéÐÒÕúóÓÔíÝáãìõâêòµ\u00b8¶·¹\u00a8»¾¼½Æ©ÇË®ÎÏÑªÖ×ØÜÞßäôù«åæç¬ëîïñ­øö÷ýûüþ¡¢§£¤¥¦";

        public static string UnicodeString = "ấẩốẫèộéềếỳúểễớíỏóỡừõờũàáảãạăằắẳẵặâầậđẻẽẹêệìỉĩịòọụựôồổỗơởợùủưứửữýỷỹỵĂÂĐÊÔƠƯ";

        public static DataSet GetData(ImportInfo info, string convertFont)
        {
            DataSet dtImport = new DataSet();
            try
            {
                DataTable dataFromExcel = GetDataFromExcel(info.FileName, info.FieldNotNull);
                DataTable structTable = GetStructTable(info.TableTemplate);
                if (dataFromExcel == null || structTable == null)
                {
                    return null;
                }
                dtImport.Tables.Add(dataFromExcel.Copy());
                dtImport.Tables.Add(structTable.Copy());
                ConverDateTime(ref dtImport);
                ConverNumeric(ref dtImport);
                if (!(convertFont.Trim() == "2"))
                {
                    return dtImport;
                }
                ConverFont(ref dtImport);
                return dtImport;
            }
            catch (Exception ex)
            {
                if (StartUp.waiting != null)
                {
                    StartUp.waiting.Close();
                }
                MessageBox.Show("Lỗi cột column");
                ExMessageBox.Show(140, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return null;
            }
        }

        public void setupPackage()
        {
        }

        private static readonly XNamespace NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static readonly XNamespace NsOfficeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static DataTable GetDataFromExcel(string _fileName, string field_not_null)
        {
            DataTable dataTable;
            try
            {
                dataTable = ReadExcelFirstSheet(_fileName);
            }
            catch (Exception ex3)
            {
                if (StartUp.waiting != null)
                {
                    StartUp.waiting.Close();
                }
                ExMessageBox.Show(150, StartupBase.SasObj, "[" + ex3.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return null;
            }
            if (dataTable == null)
            {
                return null;
            }
            DataTable tbExcel = XoaDongTrang(dataTable);
            XoaCotTrang(ref tbExcel);
            tbExcel.TableName = "DataExcel";
            return SetColumnName(tbExcel, field_not_null);
        }

        // Không phải mọi file mang đuôi .xlsx đều đúng chuẩn OOXML/zip - nhiều file do
        // ERP/report khác xuất ra thực chất là "Excel XML Spreadsheet" (SpreadsheetML) hoặc
        // định dạng .xls cũ, chỉ đổi đuôi. Excel tự nhận diện và mở được (kèm cảnh báo), nhưng
        // Package.Open thì không. Nên thử lần lượt: OOXML/zip (nhanh, không đoán kiểu dữ liệu)
        // -> SpreadsheetML XML -> OLEDB (khoan dung nhất, dùng làm phương án cuối).
        private static DataTable ReadExcelFirstSheet(string fileName)
        {
            try
            {
                return ReadExcelFirstSheetOpc(fileName);
            }
            catch
            {
            }
            try
            {
                return ReadExcelFirstSheetSpreadsheetMl(fileName);
            }
            catch
            {
            }
            return ReadExcelFirstSheetOleDb(fileName);
        }

        private static DataTable ReadExcelFirstSheetOpc(string fileName)
        {
            using (Package package = Package.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                PackagePart workbookPart = package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative));
                XDocument workbookDoc = LoadXml(workbookPart);
                XElement firstSheetEl = workbookDoc.Root.Element(NsMain + "sheets").Elements(NsMain + "sheet").FirstOrDefault();
                if (firstSheetEl == null)
                {
                    return null;
                }
                string rId = firstSheetEl.Attribute(NsOfficeRel + "id").Value;
                PackageRelationship sheetRel = workbookPart.GetRelationship(rId);
                Uri sheetUri = PackUriHelper.ResolvePartUri(workbookPart.Uri, sheetRel.TargetUri);
                PackagePart sheetPart = package.GetPart(sheetUri);
                string[] sharedStrings = ReadSharedStrings(package, workbookPart);
                bool[] dateStyles = ReadCellStyleIsDate(package, workbookPart);
                XDocument sheetDoc = LoadXml(sheetPart);
                List<XElement> rowElements = sheetDoc.Root.Element(NsMain + "sheetData").Elements(NsMain + "row").ToList();
                if (rowElements.Count == 0)
                {
                    return null;
                }
                Dictionary<int, string> headers = new Dictionary<int, string>();
                int maxCol = 0;
                foreach (XElement cell in rowElements[0].Elements(NsMain + "c"))
                {
                    int colIndex = ColumnLetterToIndex(GetColumnLetter(cell.Attribute("r").Value));
                    headers[colIndex] = GetCellValue(cell, sharedStrings, dateStyles);
                    if (colIndex > maxCol)
                    {
                        maxCol = colIndex;
                    }
                }
                DataTable dataTable = new DataTable("DataExcel");
                dataTable.Columns.Add("Stt_(stt):IV", typeof(int));
                dataTable.Columns[0].AutoIncrement = true;
                dataTable.Columns[0].AutoIncrementSeed = 2L;
                dataTable.Columns[0].AutoIncrementStep = 1L;
                for (int col = 1; col <= maxCol; col++)
                {
                    string headerName = headers.ContainsKey(col) ? headers[col] : ("F" + col);
                    dataTable.Columns.Add(headerName, typeof(string));
                }
                for (int r = 1; r < rowElements.Count; r++)
                {
                    DataRow row = dataTable.NewRow();
                    foreach (XElement cell in rowElements[r].Elements(NsMain + "c"))
                    {

                        int colIndex = ColumnLetterToIndex(GetColumnLetter(cell.Attribute("r").Value));
                        if (colIndex >= 1 && colIndex <= maxCol)
                        {
                            row[colIndex] = GetCellValue(cell, sharedStrings, dateStyles);
                        }
                    }
                    dataTable.Rows.Add(row);
                }
                return dataTable;
            }
        }

        private static readonly XNamespace NsSpreadsheetMl = "urn:schemas-microsoft-com:office:spreadsheet";

        private static DataTable ReadExcelFirstSheetSpreadsheetMl(string fileName)
        {
            XDocument doc = XDocument.Load(fileName);
            XElement worksheetEl = doc.Root.Elements(NsSpreadsheetMl + "Worksheet").FirstOrDefault();
            XElement tableEl = worksheetEl?.Element(NsSpreadsheetMl + "Table");
            if (tableEl == null)
            {
                return null;
            }
            List<XElement> rowElements = tableEl.Elements(NsSpreadsheetMl + "Row").ToList();
            if (rowElements.Count == 0)
            {
                return null;
            }
            Dictionary<int, string> headers = new Dictionary<int, string>();
            int maxCol = 0;
            foreach (KeyValuePair<int, string> cell in GetSpreadsheetMlRowCells(rowElements[0]))
            {
                headers[cell.Key] = cell.Value;
                if (cell.Key > maxCol)
                {
                    maxCol = cell.Key;
                }
            }
            DataTable dataTable = new DataTable("DataExcel");
            dataTable.Columns.Add("Stt_(stt):IV", typeof(int));
            dataTable.Columns[0].AutoIncrement = true;
            dataTable.Columns[0].AutoIncrementSeed = 2L;
            dataTable.Columns[0].AutoIncrementStep = 1L;
            for (int col = 1; col <= maxCol; col++)
            {
                string headerName = headers.ContainsKey(col) ? headers[col] : ("F" + col);
                dataTable.Columns.Add(headerName, typeof(string));
            }
            for (int r = 1; r < rowElements.Count; r++)
            {
                DataRow row = dataTable.NewRow();
                foreach (KeyValuePair<int, string> cell in GetSpreadsheetMlRowCells(rowElements[r]))
                {
                    if (cell.Key >= 1 && cell.Key <= maxCol)
                    {
                        row[cell.Key] = cell.Value;
                    }
                }
                dataTable.Rows.Add(row);
            }
            return dataTable;
        }

        // Cell "ss:Index" đánh dấu vị trí cột thực khi cột trước đó bị bỏ trống (không ghi
        // <Cell>), nên phải cộng dồn thủ công thay vì dùng thứ tự phần tử.
        private static IEnumerable<KeyValuePair<int, string>> GetSpreadsheetMlRowCells(XElement rowEl)
        {
            int currentIndex = 0;
            foreach (XElement cellEl in rowEl.Elements(NsSpreadsheetMl + "Cell"))
            {
                XAttribute indexAttr = cellEl.Attribute(NsSpreadsheetMl + "Index");
                currentIndex = (indexAttr != null) ? int.Parse(indexAttr.Value, CultureInfo.InvariantCulture) : (currentIndex + 1);
                XElement dataEl = cellEl.Element(NsSpreadsheetMl + "Data");
                yield return new KeyValuePair<int, string>(currentIndex, (dataEl != null) ? dataEl.Value : "");
            }
        }

        // Phương án cuối cùng: dùng driver OLEDB (khoan dung nhất với các file lệch chuẩn),
        // chỉ dùng khi cả hai cách đọc XML ở trên đều thất bại.
        private static DataTable ReadExcelFirstSheetOleDb(string fileName)
        {
            string excelFormat = string.Equals(Path.GetExtension(fileName), ".xls", StringComparison.OrdinalIgnoreCase) ? "Excel 8.0" : "Excel 12.0 Xml";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"" + excelFormat + ";HDR=Yes;IMEX=1\"";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                if (schemaTable == null || schemaTable.Rows.Count == 0)
                {
                    return null;
                }
                string sheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();
                DataTable rawTable = new DataTable();
                using (OleDbCommand command = new OleDbCommand("SELECT * FROM [" + sheetName + "]", connection))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                {
                    adapter.Fill(rawTable);
                }
                DataTable dataTable = new DataTable("DataExcel");
                dataTable.Columns.Add("Stt_(stt):IV", typeof(int));
                dataTable.Columns[0].AutoIncrement = true;
                dataTable.Columns[0].AutoIncrementSeed = 2L;
                dataTable.Columns[0].AutoIncrementStep = 1L;
                foreach (DataColumn column in rawTable.Columns)
                {
                    dataTable.Columns.Add(column.ColumnName, typeof(string));
                }
                foreach (DataRow rawRow in rawTable.Rows)
                {
                    DataRow row = dataTable.NewRow();
                    for (int i = 0; i < rawTable.Columns.Count; i++)
                    {
                        row[i + 1] = rawRow[i]?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(row);
                }
                return dataTable;
            }
        }

        private static XDocument LoadXml(PackagePart part)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
            using (System.Xml.XmlReader xmlReader = System.Xml.XmlReader.Create(stream))
            {
                return XDocument.Load(xmlReader);
            }
        }

        private static string[] ReadSharedStrings(Package package, PackagePart workbookPart)
        {
            PackageRelationship rel = workbookPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings").FirstOrDefault();
            if (rel == null)
            {
                return new string[0];
            }
            Uri uri = PackUriHelper.ResolvePartUri(workbookPart.Uri, rel.TargetUri);
            if (!package.PartExists(uri))
            {
                return new string[0];
            }
            XDocument doc = LoadXml(package.GetPart(uri));
            List<string> list = new List<string>();
            foreach (XElement si in doc.Root.Elements(NsMain + "si"))
            {
                XElement tEl = si.Element(NsMain + "t");
                string text = (tEl != null) ? tEl.Value : string.Concat(si.Elements(NsMain + "r").Select((XElement r) => (string)r.Element(NsMain + "t")));
                list.Add(text);
            }
            return list.ToArray();
        }

        // Excel lưu ô ngày dưới dạng số serial (số ngày kể từ 1899-12-30), định dạng hiển thị
        // nằm ở styles.xml (cellXfs -> numFmtId), không nằm trong sheet.xml. Phải đọc styles.xml
        // để biết style index nào ứng với numFmt kiểu ngày thì mới convert đúng, nếu không số
        // serial (vd 45885) sẽ bị đưa thẳng ra ngoài thay vì chuỗi ngày.
        private static readonly int[] BuiltInDateNumFmtIds = new int[] { 14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58 };

        private static bool[] ReadCellStyleIsDate(Package package, PackagePart workbookPart)
        {
            PackageRelationship rel = workbookPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles").FirstOrDefault();
            if (rel == null)
            {
                return new bool[0];
            }
            Uri uri = PackUriHelper.ResolvePartUri(workbookPart.Uri, rel.TargetUri);
            if (!package.PartExists(uri))
            {
                return new bool[0];
            }
            XDocument doc = LoadXml(package.GetPart(uri));
            Dictionary<int, string> customFormats = new Dictionary<int, string>();
            XElement numFmtsEl = doc.Root.Element(NsMain + "numFmts");
            if (numFmtsEl != null)
            {
                foreach (XElement numFmtEl in numFmtsEl.Elements(NsMain + "numFmt"))
                {
                    int fmtId = int.Parse(numFmtEl.Attribute("numFmtId").Value, CultureInfo.InvariantCulture);
                    customFormats[fmtId] = (string)numFmtEl.Attribute("formatCode") ?? "";
                }
            }
            XElement cellXfsEl = doc.Root.Element(NsMain + "cellXfs");
            if (cellXfsEl == null)
            {
                return new bool[0];
            }
            List<XElement> xfs = cellXfsEl.Elements(NsMain + "xf").ToList();
            bool[] result = new bool[xfs.Count];
            for (int i = 0; i < xfs.Count; i++)
            {
                XAttribute numFmtIdAttr = xfs[i].Attribute("numFmtId");
                int numFmtId = (numFmtIdAttr != null) ? int.Parse(numFmtIdAttr.Value, CultureInfo.InvariantCulture) : 0;
                result[i] = IsDateNumFmt(numFmtId, customFormats);
            }
            return result;
        }

        private static bool IsDateNumFmt(int numFmtId, Dictionary<int, string> customFormats)
        {
            if (BuiltInDateNumFmtIds.Contains(numFmtId))
            {
                return true;
            }
            if (numFmtId < 164)
            {
                return false;
            }
            string code;
            if (!customFormats.TryGetValue(numFmtId, out code) || string.IsNullOrEmpty(code))
            {
                return false;
            }
            string stripped = Regex.Replace(code, "\"[^\"]*\"", "");
            stripped = Regex.Replace(stripped, "\\[[^\\]]*\\]", "");
            return Regex.IsMatch(stripped, "[dmyhs]", RegexOptions.IgnoreCase);
        }

        private static string GetColumnLetter(string cellRef)
        {
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i]))
            {
                i++;
            }
            return cellRef.Substring(0, i);
        }

        private static int ColumnLetterToIndex(string columnLetter)
        {
            int index = 0;
            foreach (char c in columnLetter)
            {
                index = index * 26 + (c - 'A' + 1);
            }
            return index;
        }

        private static readonly DateTime ExcelEpoch = new DateTime(1899, 12, 30);

        private static string GetCellValue(XElement cell, string[] sharedStrings, bool[] dateStyles)
        {
            string t = (string)cell.Attribute("t");
            if (t == "inlineStr")
            {
                XElement isEl = cell.Element(NsMain + "is");
                return (isEl != null) ? string.Concat(isEl.Descendants(NsMain + "t").Select((XElement x) => x.Value)) : "";
            }
            XElement vEl = cell.Element(NsMain + "v");
            if (vEl == null)
            {
                return "";
            }
            if (t == "s")
            {
                int idx;
                if (int.TryParse(vEl.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx) && idx >= 0 && idx < sharedStrings.Length)
                {
                    return sharedStrings[idx];
                }
                return "";
            }
            if (string.IsNullOrEmpty(t) || t == "n")
            {
                XAttribute styleAttr = cell.Attribute("s");
                int styleIndex = (styleAttr != null) ? int.Parse(styleAttr.Value, CultureInfo.InvariantCulture) : 0;
                double serial;
                if (styleIndex >= 0 && styleIndex < dateStyles.Length && dateStyles[styleIndex]
                    && double.TryParse(vEl.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out serial))
                {
                    return ExcelEpoch.AddDays(serial).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                }
                return NormalizeNumericCellValue(vEl.Value);
            }
            return vEl.Value;
        }

        // Excel lưu số trong sheet.xml theo kiểu round-trip của double, nên số thập phân nhỏ
        // (vd 0.068) bị ghi dưới dạng ký hiệu khoa học ("6.8000000000000005E-2") và số bình
        // thường cũng dính nhiễu double (vd "8160.0000000000009" thay vì "8160"). Chuỗi "E..."
        // khi insert thẳng vào cột numeric/decimal SQL Server sẽ lỗi convert varchar -> numeric
        // (implicit conversion không hiểu ký hiệu khoa học). Excel bản thân chỉ giữ 15 chữ số có
        // nghĩa, nên làm tròn theo đúng độ chính xác đó (G15) để vừa bỏ nhiễu vừa tránh ký hiệu E.
        private static string NormalizeNumericCellValue(string raw)
        {
            double d;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
            {
                return raw;
            }
            string formatted = d.ToString("G15", CultureInfo.InvariantCulture);
            if (formatted.IndexOf('E') < 0)
            {
                return formatted;
            }
            decimal dec;
            return decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out dec)
                ? dec.ToString(CultureInfo.InvariantCulture)
                : formatted;
        }

        public static DataTable XoaDongTrang(DataTable tbExcel)
        {
            bool flag = false;
            DataTable dataTable = tbExcel.Clone();
            for (int i = 0; i < tbExcel.Rows.Count; i++)
            {
                flag = false;
                for (int j = 1; j < tbExcel.Columns.Count; j++)
                {
                    if (tbExcel.Rows[i][j].ToString().Trim() != "")
                    {
                        flag = true;
                        break;
                    }
                }
                if (flag)
                {
                    dataTable.ImportRow(tbExcel.Rows[i]);
                }
            }
            return dataTable;
        }

        public static void XoaCotTrang(ref DataTable tbExcel)
        {
            string text = "";
            int result = 0;
            for (int i = 0; i < tbExcel.Columns.Count; i++)
            {
                text = tbExcel.Columns[i].ColumnName.Trim();
                if (text.IndexOf("F") == 0 && int.TryParse(text.Substring(1, text.Length - 1), out result))
                {
                    tbExcel.Columns.Remove(text);
                    i = -1;
                }
            }
        }

        private static DataTable SetColumnName(DataTable tb, string field_not_null)
        {
            try
            {
                StrBrowse = "";
                StrBrowseFieldNull = "stt:H=Dòng";
                string text = "";
                string[] source = field_not_null.Split(';');
                for (int i = 0; i < tb.Columns.Count; i++)
                {
                    string text2 = tb.Columns[i].ColumnName.ToString();
                    int num = text2.LastIndexOf('(');
                    int num2 = text2.LastIndexOf(')');
                    if (num == -1 || num2 == -1)
                    {
                        if (StartUp.waiting != null)
                        {
                            StartUp.waiting.Close();
                        }
                        ExMessageBox.Show(160, StartupBase.SasObj, $"Tên cột << [{text2}] >> trong file  excel không đúng định dạng!", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                        return null;
                    }
                    string text3 = text2.Substring(num + 1, num2 - num - 1).Trim();
                    string text4 = text3 + ":H=" + text2.Substring(0, num - 1).Trim() + text2.Substring(num2 + 1, text2.Length - num2 - 1).Trim();
                    text = text + text4 + ";";
                    tb.Columns[i].ColumnName = text3;
                    if (tb.Columns[i].DataType == typeof(DateTime))
                    {
                        text4 += ":D";
                    }
                    if (source.Contains(text3))
                    {
                        StrBrowseFieldNull = StrBrowseFieldNull + ";" + text4;
                    }
                }
                StrBrowse = text.Substring(0, text.Length - 1);
                StrBrowse = StrBrowse.Replace("#", ".");
                StrBrowseFieldNull = StrBrowseFieldNull.Replace("#", ".");
                return tb;
            }
            catch (Exception ex)
            {
                ExMessageBox.Show(165, StartupBase.SasObj, "[" + ex.Message + "]", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
                return null;
            }
        }

        public static DataTable GetStructTable(string tableName)
        {
            if (tableName == null || tableName == "")
            {
                return null;
            }
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandText += $"SELECT * FROM information_schema.columns WHERE table_name like '{tableName.Trim()}'";
            DataSet dataSet = StartupBase.SasObj.ExcuteReader(sqlCommand);
            if (dataSet == null || dataSet.Tables.Count == 0)
            {
                return null;
            }
            dataSet.Tables[0].TableName = "StrucImex";
            return dataSet.Tables[0];
        }

        public static void ConverFont(ref DataSet dtImport)
        {
            for (int i = 0; i < dtImport.Tables["DataExcel"].Rows.Count; i++)
            {
                for (int j = 0; j < dtImport.Tables["DataExcel"].Columns.Count; j++)
                {
                    DataRow[] array = dtImport.Tables["StrucImex"].Select("column_name = '" + dtImport.Tables["DataExcel"].Columns[j] + "'");
                    if (array.Length <= 0)
                    {
                        continue;
                    }
                    int result = 0;
                    if (int.TryParse(array[0]["character_maximum_length"].ToString(), out result) && result != -1)
                    {
                        string value = ConvertTcvn3ToUnicode(dtImport.Tables["DataExcel"].Rows[i][j].ToString().Trim());
                        if (!string.IsNullOrEmpty(value))
                        {
                            dtImport.Tables["DataExcel"].Rows[i][j] = value;
                        }
                    }
                }
            }
        }

        private static void ConverDateTime(ref DataSet dtImport)
        {
            for (int i = 0; i < dtImport.Tables["DataExcel"].Rows.Count; i++)
            {
                for (int j = 0; j < dtImport.Tables["DataExcel"].Columns.Count; j++)
                {
                    DataRow[] array = dtImport.Tables["StrucImex"].Select("column_name = '" + dtImport.Tables["DataExcel"].Columns[j] + "'");
                    if (array.Length > 0 && array[0]["data_type"].ToString().Trim() == "smalldatetime" && !dtImport.Tables["DataExcel"].Rows[i][j].GetType().FullName.Equals("System.DateTime"))
                    {
                        string[] array2 = dtImport.Tables["DataExcel"].Rows[i][j].ToString().Replace(" ", "").Replace("/", "-")
                            .Split('-');
                        if (array2.Length == 3)
                        {
                            dtImport.Tables["DataExcel"].Rows[i][j] = array2[2].Substring(0, 4) + ((array2[1].Length > 1) ? array2[1] : ("0" + array2[1])) + ((array2[0].Length > 1) ? array2[0] : ("0" + array2[0]));
                        }
                    }
                }
            }
        }

        // Ô Excel để trống ở cột numeric/decimal đọc ra là chuỗi rỗng "", không phải NULL.
        // Insert thẳng "" vào cột numeric bên SQL Server bị lỗi "Error converting data
        // type nvarchar to numeric" (khác NULL, "" không phải literal số hợp lệ). Ở đây
        // quy ô trống về "0" và loại dấu phẩy ngăn cách hàng nghìn trước khi ghi xuống DB.
        private static readonly string[] NumericDataTypes = new string[] { "numeric", "decimal", "float", "money", "smallmoney", "int", "bigint", "smallint", "tinyint", "real" };

        private static void ConverNumeric(ref DataSet dtImport)
        {
            for (int i = 0; i < dtImport.Tables["DataExcel"].Rows.Count; i++)
            {
                for (int j = 0; j < dtImport.Tables["DataExcel"].Columns.Count; j++)
                {
                    DataRow[] array = dtImport.Tables["StrucImex"].Select("column_name = '" + dtImport.Tables["DataExcel"].Columns[j] + "'");
                    if (array.Length <= 0 || !NumericDataTypes.Contains(array[0]["data_type"].ToString().Trim()))
                    {
                        continue;
                    }
                    string value = dtImport.Tables["DataExcel"].Rows[i][j].ToString().Trim();
                    dtImport.Tables["DataExcel"].Rows[i][j] = string.IsNullOrEmpty(value) ? "0" : value.Replace(",", "");
                }
            }
        }

        public static List<char> GetAllChar(string s)
        {
            List<char> list = new List<char>();
            int length = s.Length;
            for (int i = 0; i < length; i++)
            {
                if (!list.Contains(s[i]) && TCNV3String.Contains(s[i]))
                {
                    list.Add(s[i]);
                }
            }
            int count = list.Count;
            for (int j = 0; j < count - 1; j++)
            {
                for (int k = j + 1; k < count; k++)
                {
                    if (TCNV3String.IndexOf(list[j]) > TCNV3String.IndexOf(list[k]))
                    {
                        char value = list[j];
                        list[j] = list[k];
                        list[k] = value;
                    }
                }
            }
            return list;
        }

        public static string ConvertTcvn3ToUnicode(string input)
        {
            input = input.Trim();
            List<char> allChar = GetAllChar(input);
            int num = allChar.Count();
            for (int i = 0; i < num; i++)
            {
                int num2 = TCNV3String.IndexOf(allChar[i]);
                if (num2 >= 0)
                {
                    input = input.Replace(TCNV3String[num2], UnicodeString[num2]);
                }
            }
            return input;
        }
    }
}
