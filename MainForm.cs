using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace XlsToAccessImporter
{
    internal sealed class MainForm : Form
    {
        private readonly ExcelReader _excelReader = new ExcelReader();
        private readonly SchemaInferenceService _schemaInferenceService = new SchemaInferenceService();
        private readonly AccessService _accessService = new AccessService();

        private TextBox _txtExcelPath;
        private TextBox _txtAccessPath;
        private ComboBox _cmbSheets;
        private TextBox _txtTableName;
        private CheckBox _chkForceAllText;
        private DataGridView _gridColumns;
        private DataGridView _gridPreview;
        private TextBox _txtLog;
        private Button _btnImport;
        private Button _btnCancel;
        private ProgressBar _progressBar;
        private Label _lblStatus;
        private NumericUpDown _numBatchSize;
        private DataTable _currentSheet;
        private List<ColumnDefinition> _columnDefinitions;
        private CancellationTokenSource _importCancellation;

        public MainForm()
        {
            Text = "XLS 导入 Access 工具";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1100, 760);
            Size = new Size(1220, 820);
            BuildUi();
        }

        private void BuildUi()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(12);
            root.RowCount = 5;
            root.ColumnCount = 1;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 25));

            root.Controls.Add(BuildFilePanel(), 0, 0);
            root.Controls.Add(BuildOptionsPanel(), 0, 1);
            root.Controls.Add(BuildColumnsPanel(), 0, 2);
            root.Controls.Add(BuildPreviewPanel(), 0, 3);
            root.Controls.Add(BuildLogPanel(), 0, 4);

            Controls.Add(root);
        }

        private Control BuildFilePanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.AutoSize = true;
            panel.ColumnCount = 4;
            panel.RowCount = 2;
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            panel.Padding = new Padding(0, 0, 0, 8);

            panel.Controls.Add(new Label { Text = "Excel 文件", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 0);
            _txtExcelPath = new TextBox { Dock = DockStyle.Fill };
            panel.Controls.Add(_txtExcelPath, 1, 0);
            Button btnExcel = new Button { Text = "选择 Excel", Dock = DockStyle.Fill };
            btnExcel.Click += OnChooseExcel;
            panel.Controls.Add(btnExcel, 2, 0);
            panel.Controls.Add(new Label { Text = string.Empty, Dock = DockStyle.Fill }, 3, 0);

            panel.Controls.Add(new Label { Text = "Access 文件", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 1);
            _txtAccessPath = new TextBox { Dock = DockStyle.Fill };
            panel.Controls.Add(_txtAccessPath, 1, 1);
            Button btnAccess = new Button { Text = "打开已有", Dock = DockStyle.Fill };
            btnAccess.Click += OnChooseAccess;
            panel.Controls.Add(btnAccess, 2, 1);
            Button btnNewMdb = new Button { Text = "新建 MDB", Dock = DockStyle.Fill };
            btnNewMdb.Click += OnChooseNewMdb;
            panel.Controls.Add(btnNewMdb, 3, 1);

            return panel;
        }

        private Control BuildOptionsPanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.AutoSize = true;
            panel.ColumnCount = 9;
            panel.RowCount = 2;
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            panel.Padding = new Padding(0, 0, 0, 8);
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            panel.Controls.Add(new Label { Text = "工作表", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 0, 0);
            _cmbSheets = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbSheets.SelectedIndexChanged += OnSheetChanged;
            panel.Controls.Add(_cmbSheets, 1, 0);

            panel.Controls.Add(new Label { Text = "表名", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 2, 0);
            _txtTableName = new TextBox { Dock = DockStyle.Fill };
            panel.Controls.Add(_txtTableName, 3, 0);

            _chkForceAllText = new CheckBox { Text = "全文本存入", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            _chkForceAllText.CheckedChanged += delegate { RefreshColumnTypeHints(); };
            panel.Controls.Add(_chkForceAllText, 4, 0);

            panel.Controls.Add(new Label { Text = "批量", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill }, 5, 0);
            _numBatchSize = new NumericUpDown { Dock = DockStyle.Fill, Minimum = 100, Maximum = 5000, Increment = 100, Value = 1000, ThousandsSeparator = true };
            panel.Controls.Add(_numBatchSize, 6, 0);

            _btnImport = new Button { Text = "开始导入", Dock = DockStyle.Fill };
            _btnImport.Click += OnImport;
            panel.Controls.Add(_btnImport, 7, 0);

            _btnCancel = new Button { Text = "取消导入", Dock = DockStyle.Fill, Enabled = false };
            _btnCancel.Click += OnCancelImport;
            panel.Controls.Add(_btnCancel, 8, 0);

            _progressBar = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Visible = false };
            panel.SetColumnSpan(_progressBar, 2);
            panel.Controls.Add(_progressBar, 7, 1);

            _lblStatus = new Label { Text = "状态: 待就绪", Dock = DockStyle.Fill, AutoEllipsis = true };
            panel.SetColumnSpan(_lblStatus, 7);
            panel.Controls.Add(_lblStatus, 0, 1);

            return panel;
        }

        private Control BuildColumnsPanel()
        {
            GroupBox group = new GroupBox();
            group.Text = "字段设置";
            group.Dock = DockStyle.Fill;

            _gridColumns = new DataGridView();
            _gridColumns.Dock = DockStyle.Fill;
            _gridColumns.AllowUserToAddRows = false;
            _gridColumns.AllowUserToDeleteRows = false;
            _gridColumns.AutoGenerateColumns = false;
            _gridColumns.RowHeadersVisible = false;
            _gridColumns.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            _gridColumns.Columns.Add(new DataGridViewTextBoxColumn { Name = "OriginalName", HeaderText = "原列名", DataPropertyName = "OriginalName", ReadOnly = true, Width = 220 });
            _gridColumns.Columns.Add(new DataGridViewTextBoxColumn { Name = "FinalName", HeaderText = "最终列名", DataPropertyName = "FinalName", Width = 220 });
            _gridColumns.Columns.Add(new DataGridViewTextBoxColumn { Name = "InferredTypeText", HeaderText = "自动判断", DataPropertyName = "InferredType", ReadOnly = true, Width = 120 });

            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.Name = "SelectedType";
            typeColumn.HeaderText = "手动类型";
            typeColumn.DataPropertyName = "SelectedType";
            typeColumn.Width = 120;
            typeColumn.DataSource = Enum.GetValues(typeof(ColumnTypeOption));
            _gridColumns.Columns.Add(typeColumn);

            group.Controls.Add(_gridColumns);
            return group;
        }

        private Control BuildPreviewPanel()
        {
            GroupBox group = new GroupBox();
            group.Text = "数据预览";
            group.Dock = DockStyle.Fill;

            _gridPreview = new DataGridView();
            _gridPreview.Dock = DockStyle.Fill;
            _gridPreview.AllowUserToAddRows = false;
            _gridPreview.AllowUserToDeleteRows = false;
            _gridPreview.ReadOnly = true;
            _gridPreview.RowHeadersVisible = false;
            _gridPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            group.Controls.Add(_gridPreview);
            return group;
        }

        private Control BuildLogPanel()
        {
            GroupBox group = new GroupBox();
            group.Text = "导入日志";
            group.Dock = DockStyle.Fill;

            _txtLog = new TextBox();
            _txtLog.Dock = DockStyle.Fill;
            _txtLog.Multiline = true;
            _txtLog.ScrollBars = ScrollBars.Both;
            _txtLog.ReadOnly = true;

            group.Controls.Add(_txtLog);
            return group;
        }

        private void OnChooseExcel(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel 文件 (*.xls;*.xlsx)|*.xls;*.xlsx";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _txtExcelPath.Text = dialog.FileName;
                LoadSheets();
            }
        }

        private void OnChooseAccess(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Access 文件 (*.accdb;*.mdb)|*.accdb;*.mdb";
                dialog.Title = "选择已有 Access 文件";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _txtAccessPath.Text = dialog.FileName;
                AppendLog("已选择已有 Access 文件: " + dialog.FileName);
            }
        }

        private void OnChooseNewMdb(object sender, EventArgs e)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "Access 2003 数据库 (*.mdb)|*.mdb";
                dialog.DefaultExt = "mdb";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = false;
                dialog.Title = "新建 MDB 文件";
                dialog.FileName = BuildSuggestedMdbFileName();

                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _txtAccessPath.Text = dialog.FileName;
                AppendLog("已设置新的 MDB 文件路径: " + dialog.FileName);
            }
        }

        private void LoadSheets()
        {
            try
            {
                UpdateStatus(ImportStage.ReadingExcel, "正在读取工作表列表...");
                ToggleBusy(true, true);
                _cmbSheets.Items.Clear();
                _gridPreview.DataSource = null;
                _gridColumns.DataSource = null;
                _currentSheet = null;
                _columnDefinitions = null;

                List<ExcelSheetInfo> sheets = _excelReader.GetSheets(_txtExcelPath.Text.Trim());
                foreach (ExcelSheetInfo sheet in sheets)
                {
                    _cmbSheets.Items.Add(sheet);
                }

                if (_cmbSheets.Items.Count > 0)
                {
                    _cmbSheets.SelectedIndex = 0;
                }

                AppendLog("已加载工作表: " + sheets.Count);
                UpdateStatus(ImportStage.Completed, "工作表列表读取完成");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "读取 Excel 失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog("读取 Excel 失败: " + ex.Message);
                UpdateStatus(ImportStage.None, "读取 Excel 失败");
            }
            finally
            {
                ToggleBusy(false, false);
            }
        }

        private void OnSheetChanged(object sender, EventArgs e)
        {
            ExcelSheetInfo selectedSheet = _cmbSheets.SelectedItem as ExcelSheetInfo;
            if (selectedSheet == null || string.IsNullOrWhiteSpace(_txtExcelPath.Text))
            {
                return;
            }

            try
            {
                UpdateStatus(ImportStage.ReadingExcel, "正在读取工作表数据...");
                ToggleBusy(true, true);
                _currentSheet = _excelReader.ReadSheetPreview(_txtExcelPath.Text.Trim(), selectedSheet.QueryName, 200);
                UpdateStatus(ImportStage.AnalyzingColumns, "正在分析字段类型...");
                _gridPreview.DataSource = CreatePreviewTable(_currentSheet, 100);
                _columnDefinitions = _schemaInferenceService.BuildColumnDefinitions(_currentSheet);
                _gridColumns.DataSource = new BindingSource { DataSource = _columnDefinitions };
                if (string.IsNullOrWhiteSpace(_txtTableName.Text))
                {
                    _txtTableName.Text = selectedSheet.DisplayName;
                }

                SuggestMdbPathIfNeeded(selectedSheet.DisplayName);

                RefreshColumnTypeHints();
                AppendLog("已读取工作表预览: " + selectedSheet.DisplayName + "，预览行数: " + _currentSheet.Rows.Count);
                UpdateStatus(ImportStage.Completed, "工作表预览已加载，可开始导入");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "读取工作表失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog("读取工作表失败: " + ex.Message);
                UpdateStatus(ImportStage.None, "读取工作表失败");
            }
            finally
            {
                ToggleBusy(false, false);
            }
        }

        private void OnImport(object sender, EventArgs e)
        {
            try
            {
                ValidateBeforeImport();
                if (_importCancellation != null)
                {
                    throw new InvalidOperationException("当前已有导入任务正在执行。");
                }

                _importCancellation = new CancellationTokenSource();
                ToggleBusy(true, true);
                UpdateStatus(ImportStage.Preparing, "正在准备导入...");

                ImportOptions options = new ImportOptions
                {
                    ExcelPath = _txtExcelPath.Text.Trim(),
                    AccessPath = EnsureAccessFile(),
                    SheetQueryName = ((_cmbSheets.SelectedItem as ExcelSheetInfo) ?? new ExcelSheetInfo()).QueryName,
                    RequestedTableName = _txtTableName.Text.Trim(),
                    ForceAllText = _chkForceAllText.Checked,
                    BatchSize = Convert.ToInt32(_numBatchSize.Value),
                    Columns = ReadColumnsFromGrid(),
                    CancellationToken = _importCancellation.Token
                };

                System.Threading.ThreadPool.QueueUserWorkItem(delegate
                {
                    ImportResult result = null;
                    Exception error = null;
                    try
                    {
                        ImportService importService = new ImportService(_excelReader, _schemaInferenceService, _accessService, ReportProgress);
                        result = importService.Execute(options);
                    }
                    catch (OperationCanceledException)
                    {
                        error = new InvalidOperationException("导入已取消。");
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                    }

                    BeginInvoke(new Action(delegate
                    {
                        if (error != null)
                        {
                            string title = error.Message == "导入已取消。" ? "已取消" : "导入失败";
                            MessageBox.Show(this, error.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            AppendLog((error.Message == "导入已取消。" ? "导入取消: " : "导入失败: ") + error.Message);
                            UpdateStatus(ImportStage.None, error.Message == "导入已取消。" ? "导入已取消" : "导入失败");
                        }
                        else if (result != null)
                        {
                            AppendLog("导入完成，最终表名: " + result.FinalTableName);
                            AppendLog("成功: " + result.InsertedRows + "，失败: " + result.FailedRows);
                            for (int i = 0; i < result.Messages.Count; i++)
                            {
                                AppendLog(result.Messages[i]);
                            }

                            MessageBox.Show(this,
                                "导入完成。" + Environment.NewLine +
                                "表名: " + result.FinalTableName + Environment.NewLine +
                                "成功: " + result.InsertedRows + Environment.NewLine +
                                "失败: " + result.FailedRows,
                                "完成",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }

                        ToggleBusy(false, false);
                        _importCancellation.Dispose();
                        _importCancellation = null;
                    }));
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog("导入失败: " + ex.Message);
                UpdateStatus(ImportStage.None, "导入失败");
                ToggleBusy(false, false);
                if (_importCancellation != null)
                {
                    _importCancellation.Dispose();
                    _importCancellation = null;
                }
            }
        }

        private void OnCancelImport(object sender, EventArgs e)
        {
            if (_importCancellation == null)
            {
                return;
            }

            _importCancellation.Cancel();
            UpdateStatus(ImportStage.None, "正在取消导入...");
            AppendLog("已请求取消当前导入任务。");
        }

        private string EnsureAccessFile()
        {
            string path = _txtAccessPath.Text.Trim();
            bool exists = File.Exists(path);
            string ensuredPath = OleDbUtility.EnsureAccessDatabase(path);
            if (!exists && File.Exists(ensuredPath))
            {
                AppendLog("已自动创建空白 MDB 文件: " + ensuredPath);
            }

            return ensuredPath;
        }

        private void ValidateBeforeImport()
        {
            if (string.IsNullOrWhiteSpace(_txtExcelPath.Text) || !File.Exists(_txtExcelPath.Text.Trim()))
            {
                throw new InvalidOperationException("请选择有效的 Excel 文件。");
            }

            if (string.IsNullOrWhiteSpace(_txtAccessPath.Text))
            {
                throw new InvalidOperationException("请选择已有 Access 文件，或点击“新建 MDB”指定一个新的数据库文件。");
            }

            if (_cmbSheets.SelectedItem == null)
            {
                throw new InvalidOperationException("请选择工作表。");
            }

            if (_currentSheet == null)
            {
                throw new InvalidOperationException("当前工作表尚未加载成功。");
            }

            if (_columnDefinitions == null || _columnDefinitions.Count == 0)
            {
                throw new InvalidOperationException("当前工作表没有可导入的字段。");
            }

            if (string.IsNullOrWhiteSpace(_txtTableName.Text))
            {
                throw new InvalidOperationException("请输入目标表名。");
            }
        }

        private List<ColumnDefinition> ReadColumnsFromGrid()
        {
            List<ColumnDefinition> columns = _gridColumns.DataSource is BindingSource
                ? (List<ColumnDefinition>)((BindingSource)_gridColumns.DataSource).DataSource
                : _columnDefinitions;

            if (columns == null)
            {
                throw new InvalidOperationException("字段配置为空。");
            }

            HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Count; i++)
            {
                ColumnDefinition column = columns[i];
                column.FinalName = NameSanitizer.SanitizeColumnName(column.FinalName, i + 1, usedNames);
            }

            _gridColumns.Refresh();
            return columns;
        }

        private void RefreshColumnTypeHints()
        {
            if (_columnDefinitions == null)
            {
                return;
            }

            _gridColumns.Refresh();
            AppendLog(_chkForceAllText.Checked ? "已开启全文本存入。" : "已按自动规则识别日期时间列，其余列按文本处理。");
        }

        private string BuildSuggestedMdbFileName()
        {
            string baseName = "import_result";
            if (_cmbSheets.SelectedItem is ExcelSheetInfo)
            {
                baseName = ((ExcelSheetInfo)_cmbSheets.SelectedItem).DisplayName;
            }
            else if (!string.IsNullOrWhiteSpace(_txtExcelPath.Text))
            {
                baseName = Path.GetFileNameWithoutExtension(_txtExcelPath.Text.Trim());
            }

            string sanitized = NameSanitizer.SanitizeTableName(baseName).Replace(' ', '_');
            return sanitized + ".mdb";
        }

        private void SuggestMdbPathIfNeeded(string sheetName)
        {
            string currentPath = _txtAccessPath.Text.Trim();
            if (!string.IsNullOrWhiteSpace(currentPath))
            {
                return;
            }

            string excelPath = _txtExcelPath.Text.Trim();
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                return;
            }

            string directory = Path.GetDirectoryName(excelPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return;
            }

            string fileName = NameSanitizer.SanitizeTableName(sheetName).Replace(' ', '_') + ".mdb";
            _txtAccessPath.Text = Path.Combine(directory, fileName);
            AppendLog("已自动建议 MDB 路径: " + _txtAccessPath.Text + "。可直接导入，或点“新建 MDB”另选位置。");
        }

        private static DataTable CreatePreviewTable(DataTable sourceTable, int maxRows)
        {
            DataTable preview = sourceTable.Clone();
            for (int i = 0; i < sourceTable.Rows.Count && i < maxRows; i++)
            {
                preview.ImportRow(sourceTable.Rows[i]);
            }

            return preview;
        }

        private void AppendLog(string message)
        {
            _txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }

        private void ToggleBusy(bool busy, bool showMarquee)
        {
            UseWaitCursor = busy;
            _btnImport.Enabled = !busy;
            _btnCancel.Enabled = busy;
            _numBatchSize.Enabled = !busy;
            _cmbSheets.Enabled = !busy;
            _progressBar.Visible = busy;
            _progressBar.Style = busy && showMarquee ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
        }

        private void UpdateStatus(ImportStage stage, string message)
        {
            string text = string.IsNullOrWhiteSpace(message) ? string.Empty : message;
            _lblStatus.Text = "状态: " + text;
        }

        private void ReportProgress(ImportProgress progress)
        {
            BeginInvoke(new Action(delegate
            {
                switch (progress.Stage)
                {
                    case ImportStage.ReadingExcel:
                        UpdateStatus(progress.Stage, string.IsNullOrWhiteSpace(progress.Message) ? "正在读取 Excel 数据..." : progress.Message);
                        ToggleBusy(true, true);
                        break;
                    case ImportStage.Preparing:
                        UpdateStatus(progress.Stage, string.IsNullOrWhiteSpace(progress.Message) ? "正在准备导入..." : progress.Message);
                        ToggleBusy(true, true);
                        break;
                    case ImportStage.CreatingTable:
                        UpdateStatus(progress.Stage, string.IsNullOrWhiteSpace(progress.Message) ? "正在创建目标表..." : progress.Message);
                        ToggleBusy(true, true);
                        break;
                    case ImportStage.ImportingData:
                        ToggleBusy(true, false);
                        UpdateStatus(progress.Stage,
                            "正在导入数据: 已处理 " + progress.ProcessedRows + " / " + progress.TotalRows +
                            "，成功 " + progress.SuccessRows + "，失败 " + progress.FailedRows);
                        break;
                    case ImportStage.Completed:
                        UpdateStatus(progress.Stage,
                            "导入完成: 已处理 " + progress.ProcessedRows + " / " + progress.TotalRows +
                            "，成功 " + progress.SuccessRows + "，失败 " + progress.FailedRows);
                        ToggleBusy(false, false);
                        break;
                }
            }));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_importCancellation != null)
            {
                _importCancellation.Cancel();
            }

            base.OnFormClosing(e);
        }
    }
}
