using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace XlsToAccessImporter
{
    internal sealed class ExcelReader
    {
        public List<ExcelSheetInfo> GetSheets(string excelPath)
        {
            List<ExcelSheetInfo> sheets = new List<ExcelSheetInfo>();
            string connectionString = OleDbUtility.BuildExcelConnectionString(excelPath, true);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable == null)
                {
                    return sheets;
                }

                foreach (DataRow row in schemaTable.Rows)
                {
                    object tableNameValue = row["TABLE_NAME"];
                    string tableName = tableNameValue == null ? string.Empty : tableNameValue.ToString();
                    if (!IsWorksheetName(tableName))
                    {
                        continue;
                    }

                    string normalized = NormalizeQueryName(tableName);
                    sheets.Add(new ExcelSheetInfo
                    {
                        DisplayName = normalized.Substring(0, normalized.Length - 1),
                        QueryName = normalized,
                        RangeQueryName = NameSanitizer.EscapeIdentifier(normalized)
                    });
                }
            }

            return sheets;
        }

        public DataTable ReadSheetPreview(string excelPath, string queryName, int maxRows)
        {
            string connectionString = OleDbUtility.BuildExcelConnectionString(excelPath, true);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(BuildSelectSql(queryName), connection))
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    DataTable table = new DataTable();
                    if (reader == null)
                    {
                        return table;
                    }

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        table.Columns.Add(reader.GetName(i));
                    }

                    int rowCount = 0;
                    while (reader.Read() && rowCount < maxRows)
                    {
                        DataRow row = table.NewRow();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[i] = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                        }

                        table.Rows.Add(row);
                        rowCount++;
                    }

                    return table;
                }
            }
        }

        public IEnumerable<DataTable> ReadSheetInBatches(string excelPath, string queryName, int batchSize, Func<bool> shouldCancel, Action<int> onTotalDiscovered)
        {
            string connectionString = OleDbUtility.BuildExcelConnectionString(excelPath, true);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(BuildSelectSql(queryName), connection))
                using (OleDbDataReader reader = command.ExecuteReader())
                {
                    if (reader == null)
                    {
                        yield break;
                    }

                    DataTable batch = CreateBatchTable(reader);
                    int totalRows = 0;
                    while (reader.Read())
                    {
                        if (shouldCancel())
                        {
                            yield break;
                        }

                        DataRow row = batch.NewRow();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[i] = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                        }

                        batch.Rows.Add(row);
                        totalRows++;

                        if (batch.Rows.Count >= batchSize)
                        {
                            onTotalDiscovered(totalRows);
                            yield return batch;
                            batch = CreateBatchTable(reader);
                        }
                    }

                    if (batch.Rows.Count > 0)
                    {
                        onTotalDiscovered(totalRows);
                        yield return batch;
                    }
                }
            }
        }

        public int CountRows(string excelPath, string queryName)
        {
            string connectionString = OleDbUtility.BuildExcelConnectionString(excelPath, true);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand("SELECT COUNT(*) FROM " + NameSanitizer.EscapeIdentifier(queryName), connection))
                {
                    object count = command.ExecuteScalar();
                    return count == null || count == DBNull.Value ? 0 : Convert.ToInt32(count, CultureInfo.InvariantCulture);
                }
            }
        }

        private static bool IsWorksheetName(string tableName)
        {
            return tableName.EndsWith("$", StringComparison.OrdinalIgnoreCase) || tableName.EndsWith("$'", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeQueryName(string tableName)
        {
            string normalized = tableName.Trim('\'');
            if (!normalized.EndsWith("$", StringComparison.OrdinalIgnoreCase))
            {
                normalized += "$";
            }

            return normalized;
        }

        private static DataTable CreateBatchTable(IDataRecord reader)
        {
            DataTable table = new DataTable();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i));
            }

            return table;
        }

        private static string BuildSelectSql(string queryName)
        {
            return "SELECT * FROM " + NameSanitizer.EscapeIdentifier(queryName);
        }
    }
}
