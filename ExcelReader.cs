using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using ExcelDataReader;

namespace XlsToAccessImporter
{
    internal sealed class ExcelReader
    {
        static ExcelReader()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        public List<ExcelSheetInfo> GetSheets(string excelPath)
        {
            List<ExcelSheetInfo> sheets = new List<ExcelSheetInfo>();
            using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = CreateReader(stream, excelPath))
            {
                DataSet dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = delegate
                    {
                        return new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        };
                    }
                });

                for (int i = 0; i < dataSet.Tables.Count; i++)
                {
                    DataTable table = dataSet.Tables[i];
                    sheets.Add(new ExcelSheetInfo
                    {
                        DisplayName = table.TableName,
                        QueryName = table.TableName,
                        RangeQueryName = table.TableName
                    });
                }
            }

            return sheets;
        }

        public DataTable ReadSheetPreview(string excelPath, string queryName, int maxRows)
        {
            using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = CreateReader(stream, excelPath))
            {
                MoveToSheet(reader, queryName);

                DataTable table = null;
                int rowCount = 0;
                bool headerConsumed = false;
                while (reader.Read())
                {
                    if (!headerConsumed)
                    {
                        table = CreateTableFromHeader(reader);
                        headerConsumed = true;
                        continue;
                    }

                    if (table == null)
                    {
                        continue;
                    }

                    DataRow row = table.NewRow();
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        row[i] = reader.GetValue(i) ?? DBNull.Value;
                    }

                    table.Rows.Add(row);
                    rowCount++;
                    if (rowCount >= maxRows)
                    {
                        break;
                    }
                }

                return table ?? new DataTable(queryName);
            }
        }

        public IEnumerable<DataTable> ReadSheetInBatches(string excelPath, string queryName, int batchSize, Func<bool> shouldCancel, Action<int> onRowsRead)
        {
            using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = CreateReader(stream, excelPath))
            {
                MoveToSheet(reader, queryName);

                DataTable currentBatch = null;
                int totalRead = 0;
                bool headerConsumed = false;
                while (reader.Read())
                {
                    if (shouldCancel())
                    {
                        yield break;
                    }

                    if (!headerConsumed)
                    {
                        currentBatch = CreateTableFromHeader(reader);
                        headerConsumed = true;
                        continue;
                    }

                    if (currentBatch == null)
                    {
                        continue;
                    }

                    DataRow row = currentBatch.NewRow();
                    for (int i = 0; i < currentBatch.Columns.Count; i++)
                    {
                        row[i] = reader.GetValue(i) ?? DBNull.Value;
                    }

                    currentBatch.Rows.Add(row);
                    totalRead++;

                    if (currentBatch.Rows.Count >= batchSize)
                    {
                        onRowsRead(totalRead);
                        yield return currentBatch;
                        currentBatch = CloneStructure(currentBatch);
                    }
                }

                if (currentBatch != null && currentBatch.Rows.Count > 0)
                {
                    onRowsRead(totalRead);
                    yield return currentBatch;
                }
            }
        }

        public int CountRows(string excelPath, string queryName)
        {
            using (FileStream stream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (IExcelDataReader reader = CreateReader(stream, excelPath))
            {
                MoveToSheet(reader, queryName);
                int count = -1;
                while (reader.Read())
                {
                    count++;
                }

                return Math.Max(0, count);
            }
        }

        private static IExcelDataReader CreateReader(Stream stream, string excelPath)
        {
            string extension = Path.GetExtension(excelPath) ?? string.Empty;
            if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return ExcelReaderFactory.CreateBinaryReader(stream);
            }

            return ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        private static void MoveToSheet(IExcelDataReader reader, string queryName)
        {
            do
            {
                if (string.Equals(reader.Name, queryName, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }
            while (reader.NextResult());

            throw new InvalidOperationException("未找到工作表: " + queryName);
        }

        private static DataTable CreateTableFromHeader(IExcelDataReader reader)
        {
            DataTable table = new DataTable(reader.Name);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                string columnName = Convert.ToString(reader.GetValue(i), CultureInfo.CurrentCulture);
                if (string.IsNullOrWhiteSpace(columnName))
                {
                    columnName = "F" + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                table.Columns.Add(columnName);
            }

            return table;
        }

        private static DataTable CloneStructure(DataTable source)
        {
            DataTable table = new DataTable(source.TableName);
            for (int i = 0; i < source.Columns.Count; i++)
            {
                table.Columns.Add(source.Columns[i].ColumnName);
            }

            return table;
        }
    }
}
