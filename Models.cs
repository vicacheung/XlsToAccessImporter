using System;
using System.Collections.Generic;
using System.Data;

namespace XlsToAccessImporter
{
    internal enum ColumnTypeOption
    {
        Auto,
        Text,
        DateTime
    }

    internal sealed class ExcelSheetInfo
    {
        public string DisplayName { get; set; }
        public string QueryName { get; set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }

    internal sealed class ColumnDefinition
    {
        public string OriginalName { get; set; }
        public string FinalName { get; set; }
        public ColumnTypeOption InferredType { get; set; }
        public ColumnTypeOption SelectedType { get; set; }
        public int SourceOrdinal { get; set; }
    }

    internal sealed class ImportOptions
    {
        public string ExcelPath { get; set; }
        public string AccessPath { get; set; }
        public string SheetQueryName { get; set; }
        public string RequestedTableName { get; set; }
        public bool ForceAllText { get; set; }
        public List<ColumnDefinition> Columns { get; set; }
        public DataTable SourceTable { get; set; }
    }

    internal sealed class ImportResult
    {
        public string FinalTableName { get; set; }
        public int InsertedRows { get; set; }
        public int FailedRows { get; set; }
        public List<string> Messages { get; private set; }

        public ImportResult()
        {
            Messages = new List<string>();
        }
    }

    internal sealed class ProviderNotFoundException : Exception
    {
        public ProviderNotFoundException(string message) : base(message)
        {
        }
    }
}
