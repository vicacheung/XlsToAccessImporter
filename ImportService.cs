using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace XlsToAccessImporter
{
    internal sealed class ImportService
    {
        private readonly SchemaInferenceService _schemaInferenceService;
        private readonly AccessService _accessService;

        public ImportService(SchemaInferenceService schemaInferenceService, AccessService accessService)
        {
            _schemaInferenceService = schemaInferenceService;
            _accessService = accessService;
        }

        public ImportResult Execute(ImportOptions options)
        {
            string accessConnectionString = OleDbUtility.BuildAccessConnectionString(options.AccessPath);
            string finalTableName = _accessService.GetUniqueTableName(accessConnectionString, options.RequestedTableName);
            _accessService.CreateTable(accessConnectionString, finalTableName, options.Columns, _schemaInferenceService, options.ForceAllText);

            ImportResult result = new ImportResult();
            result.FinalTableName = finalTableName;

            for (int rowIndex = 0; rowIndex < options.SourceTable.Rows.Count; rowIndex++)
            {
                DataRow row = options.SourceTable.Rows[rowIndex];
                if (IsBlankRow(row))
                {
                    continue;
                }

                try
                {
                    List<object> values = ConvertRowValues(row, options.Columns, options.ForceAllText, rowIndex + 2);
                    _accessService.InsertRow(accessConnectionString, finalTableName, options.Columns, values);
                    result.InsertedRows++;
                }
                catch (Exception ex)
                {
                    result.FailedRows++;
                    result.Messages.Add("第 " + (rowIndex + 2).ToString(CultureInfo.InvariantCulture) + " 行失败: " + ex.Message);
                }
            }

            return result;
        }

        private List<object> ConvertRowValues(DataRow row, IList<ColumnDefinition> columns, bool forceAllText, int excelRowNumber)
        {
            List<object> values = new List<object>();
            for (int i = 0; i < columns.Count; i++)
            {
                ColumnDefinition column = columns[i];
                object rawValue = row[column.SourceOrdinal];
                string text = rawValue == null || rawValue == DBNull.Value ? string.Empty : Convert.ToString(rawValue, CultureInfo.CurrentCulture);
                ColumnTypeOption resolvedType = _schemaInferenceService.ResolveType(column, forceAllText);

                if (string.IsNullOrWhiteSpace(text))
                {
                    values.Add(DBNull.Value);
                    continue;
                }

                if (resolvedType == ColumnTypeOption.DateTime)
                {
                    DateTime dateValue;
                    if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out dateValue) ||
                        DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dateValue))
                    {
                        values.Add(dateValue);
                    }
                    else
                    {
                        throw new InvalidOperationException("列 '" + column.FinalName + "' 的值 '" + text + "' 不是有效日期时间。请改成文本列或勾选全文本存入。Excel 行号: " + excelRowNumber.ToString(CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    values.Add(text.Trim());
                }
            }

            return values;
        }

        private static bool IsBlankRow(DataRow row)
        {
            for (int i = 0; i < row.Table.Columns.Count; i++)
            {
                object value = row[i];
                if (value != null && value != DBNull.Value && !string.IsNullOrWhiteSpace(Convert.ToString(value, CultureInfo.CurrentCulture)))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
