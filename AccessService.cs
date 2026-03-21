using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace XlsToAccessImporter
{
    internal sealed class AccessService
    {
        public bool TableExists(string accessConnectionString, string tableName)
        {
            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                connection.Open();
                DataTable schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, tableName, "TABLE" });
                return schema != null && schema.Rows.Count > 0;
            }
        }

        public string GetUniqueTableName(string accessConnectionString, string requestedName)
        {
            string sanitized = NameSanitizer.SanitizeTableName(requestedName);
            if (!TableExists(accessConnectionString, sanitized))
            {
                return sanitized;
            }

            string candidate = sanitized + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            if (!TableExists(accessConnectionString, candidate))
            {
                return candidate;
            }

            int suffix = 2;
            while (TableExists(accessConnectionString, candidate + "_" + suffix.ToString(CultureInfo.InvariantCulture)))
            {
                suffix++;
            }

            return candidate + "_" + suffix.ToString(CultureInfo.InvariantCulture);
        }

        public void CreateTable(string accessConnectionString, string tableName, IList<ColumnDefinition> columns, SchemaInferenceService schemaInferenceService, bool forceAllText)
        {
            List<string> definitions = new List<string>();
            for (int i = 0; i < columns.Count; i++)
            {
                ColumnDefinition column = columns[i];
                ColumnTypeOption type = schemaInferenceService.ResolveType(column, forceAllText);
                string accessType = type == ColumnTypeOption.DateTime ? "DATETIME" : "MEMO";
                definitions.Add(NameSanitizer.EscapeIdentifier(column.FinalName) + " " + accessType);
            }

            string sql = "CREATE TABLE " + NameSanitizer.EscapeIdentifier(tableName) + " (" + string.Join(", ", definitions.ToArray()) + ")";
            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void InsertRow(string accessConnectionString, string tableName, IList<ColumnDefinition> columns, IList<object> values)
        {
            List<string> fieldNames = new List<string>();
            List<string> placeholders = new List<string>();

            for (int i = 0; i < columns.Count; i++)
            {
                fieldNames.Add(NameSanitizer.EscapeIdentifier(columns[i].FinalName));
                placeholders.Add("?");
            }

            string sql = "INSERT INTO " + NameSanitizer.EscapeIdentifier(tableName) + " (" + string.Join(", ", fieldNames.ToArray()) + ") VALUES (" + string.Join(", ", placeholders.ToArray()) + ")";
            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    for (int i = 0; i < values.Count; i++)
                    {
                        command.Parameters.AddWithValue("@p" + i.ToString(CultureInfo.InvariantCulture), values[i] ?? DBNull.Value);
                    }

                    command.ExecuteNonQuery();
                }
            }
        }

        public void InsertBatch(string accessConnectionString, string tableName, IList<ColumnDefinition> columns, IList<IList<object>> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return;
            }

            List<string> fieldNames = new List<string>();
            List<string> placeholders = new List<string>();
            for (int i = 0; i < columns.Count; i++)
            {
                fieldNames.Add(NameSanitizer.EscapeIdentifier(columns[i].FinalName));
                placeholders.Add("?");
            }

            string sql = "INSERT INTO " + NameSanitizer.EscapeIdentifier(tableName) +
                " (" + string.Join(", ", fieldNames.ToArray()) + ") VALUES (" + string.Join(", ", placeholders.ToArray()) + ")";

            using (OleDbConnection connection = new OleDbConnection(accessConnectionString))
            {
                connection.Open();
                using (OleDbTransaction transaction = connection.BeginTransaction())
                using (OleDbCommand command = new OleDbCommand(sql, connection, transaction))
                {
                    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                    {
                        command.Parameters.Clear();
                        IList<object> row = rows[rowIndex];
                        for (int i = 0; i < row.Count; i++)
                        {
                            command.Parameters.AddWithValue("@p" + i.ToString(CultureInfo.InvariantCulture), row[i] ?? DBNull.Value);
                        }

                        command.ExecuteNonQuery();
                    }

                    transaction.Commit();
                }
            }
        }
    }
}
