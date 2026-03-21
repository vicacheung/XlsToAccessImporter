using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading;

namespace XlsToAccessImporter
{
    internal sealed class ImportService
    {
        private readonly ExcelReader _excelReader;
        private readonly SchemaInferenceService _schemaInferenceService;
        private readonly AccessService _accessService;
        private readonly Action<ImportProgress> _progressReporter;

        public ImportService(ExcelReader excelReader, SchemaInferenceService schemaInferenceService, AccessService accessService, Action<ImportProgress> progressReporter)
        {
            _excelReader = excelReader;
            _schemaInferenceService = schemaInferenceService;
            _accessService = accessService;
            _progressReporter = progressReporter ?? delegate { };
        }

        public ImportResult Execute(ImportOptions options)
        {
            CancellationToken cancellationToken = options.CancellationToken;
            _progressReporter(new ImportProgress { Stage = ImportStage.Preparing, Message = "正在准备导入..." });
            cancellationToken.ThrowIfCancellationRequested();

            _progressReporter(new ImportProgress { Stage = ImportStage.ReadingExcel, Message = "正在统计 Excel 数据行数..." });
            int totalRows = _excelReader.CountRows(options.ExcelPath, options.SheetQueryName);
            cancellationToken.ThrowIfCancellationRequested();

            string accessConnectionString = OleDbUtility.BuildAccessConnectionString(options.AccessPath);
            string finalTableName = _accessService.GetUniqueTableName(accessConnectionString, options.RequestedTableName);

            _progressReporter(new ImportProgress { Stage = ImportStage.CreatingTable, Message = "正在创建目标表..." });
            _accessService.CreateTable(accessConnectionString, finalTableName, options.Columns, _schemaInferenceService, options.ForceAllText);

            ImportResult result = new ImportResult();
            result.FinalTableName = finalTableName;

            int processedRows = 0;
            foreach (DataTable batch in _excelReader.ReadSheetInBatches(
                options.ExcelPath,
                options.SheetQueryName,
                options.BatchSize,
                delegate { return cancellationToken.IsCancellationRequested; },
                delegate(int discoveredRows)
                {
                    _progressReporter(new ImportProgress
                    {
                        Stage = ImportStage.ReadingExcel,
                        Message = "正在读取 Excel 数据... 已读取 " + discoveredRows + " 行",
                        TotalRows = totalRows,
                        ProcessedRows = processedRows,
                        SuccessRows = result.InsertedRows,
                        FailedRows = result.FailedRows
                    });
                }))
            {
                cancellationToken.ThrowIfCancellationRequested();
                List<IList<object>> rowsToInsert = new List<IList<object>>();
                List<string> batchErrors = new List<string>();

                for (int rowIndex = 0; rowIndex < batch.Rows.Count; rowIndex++)
                {
                    DataRow row = batch.Rows[rowIndex];
                    int excelRowNumber = processedRows + rowIndex + 2;
                    if (IsBlankRow(row))
                    {
                        continue;
                    }

                    try
                    {
                        rowsToInsert.Add(ConvertRowValues(row, options.Columns, options.ForceAllText, excelRowNumber));
                    }
                    catch (Exception ex)
                    {
                        result.FailedRows++;
                        batchErrors.Add("第 " + excelRowNumber.ToString(CultureInfo.InvariantCulture) + " 行失败: " + ex.Message);
                    }
                }

                if (rowsToInsert.Count > 0)
                {
                    _accessService.InsertBatch(accessConnectionString, finalTableName, options.Columns, rowsToInsert);
                    result.InsertedRows += rowsToInsert.Count;
                }

                if (batchErrors.Count > 0)
                {
                    result.Messages.AddRange(batchErrors);
                }

                processedRows += batch.Rows.Count;
                _progressReporter(new ImportProgress
                {
                    Stage = ImportStage.ImportingData,
                    ProcessedRows = processedRows,
                    TotalRows = totalRows,
                    SuccessRows = result.InsertedRows,
                    FailedRows = result.FailedRows
                });
            }

            _progressReporter(new ImportProgress
            {
                Stage = ImportStage.Completed,
                TotalRows = totalRows,
                ProcessedRows = processedRows,
                SuccessRows = result.InsertedRows,
                FailedRows = result.FailedRows,
                Message = "导入完成"
            });

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
