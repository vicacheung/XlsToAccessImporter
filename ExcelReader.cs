using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

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
                    if (tableName.EndsWith("$", StringComparison.OrdinalIgnoreCase) || tableName.EndsWith("$'", StringComparison.OrdinalIgnoreCase))
                    {
                        string display = tableName.Trim('\'');
                        if (display.EndsWith("$", StringComparison.OrdinalIgnoreCase))
                        {
                            display = display.Substring(0, display.Length - 1);
                        }

                        sheets.Add(new ExcelSheetInfo
                        {
                            DisplayName = display,
                            QueryName = tableName
                        });
                    }
                }
            }

            return sheets;
        }

        public DataTable ReadSheet(string excelPath, string queryName)
        {
            string connectionString = OleDbUtility.BuildExcelConnectionString(excelPath, true);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + queryName.Replace("]", "]]" ) + "]", connection))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    return table;
                }
            }
        }
    }
}
