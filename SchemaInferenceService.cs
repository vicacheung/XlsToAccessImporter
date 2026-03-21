using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace XlsToAccessImporter
{
    internal sealed class SchemaInferenceService
    {
        public List<ColumnDefinition> BuildColumnDefinitions(DataTable table)
        {
            List<ColumnDefinition> columns = new List<ColumnDefinition>();
            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                DataColumn column = table.Columns[i];
                ColumnTypeOption inferredType = InferDateTimeColumn(table, i) ? ColumnTypeOption.DateTime : ColumnTypeOption.Text;
                columns.Add(new ColumnDefinition
                {
                    OriginalName = column.ColumnName,
                    FinalName = NameSanitizer.SanitizeColumnName(column.ColumnName, i + 1, usedNames),
                    InferredType = inferredType,
                    SelectedType = ColumnTypeOption.Auto,
                    SourceOrdinal = i
                });
            }

            return columns;
        }

        public ColumnTypeOption ResolveType(ColumnDefinition column, bool forceAllText)
        {
            if (forceAllText)
            {
                return ColumnTypeOption.Text;
            }

            if (column.SelectedType == ColumnTypeOption.Auto)
            {
                return column.InferredType;
            }

            return column.SelectedType;
        }

        private static bool InferDateTimeColumn(DataTable table, int columnIndex)
        {
            int successCount = 0;
            int failCount = 0;
            int checkedCount = 0;
            for (int i = 0; i < table.Rows.Count && checkedCount < 50; i++)
            {
                object value = table.Rows[i][columnIndex];
                string text = value == null || value == DBNull.Value ? string.Empty : Convert.ToString(value, CultureInfo.CurrentCulture);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                checkedCount++;
                DateTime parsed;
                if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out parsed) ||
                    DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out parsed))
                {
                    successCount++;
                }
                else
                {
                    failCount++;
                }
            }

            if (checkedCount == 0)
            {
                return false;
            }

            return successCount >= 3 && failCount == 0;
        }
    }
}
