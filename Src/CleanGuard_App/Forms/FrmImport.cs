using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmImport : Form
    {
        private readonly Button _btnDownloadTemplate = new Button();
        private readonly Button _btnImportFile = new Button();
        private readonly Button _btnExportErrors = new Button();
        private readonly Button _btnCopyErrors = new Button();
        private readonly Button _btnRefreshLogs = new Button();
        private readonly TextBox _txtResult = new TextBox();
        private readonly Label _lblFailurePreview = new Label();
        private readonly DataGridView _gridFailurePreview = new DataGridView();
        private readonly DataGridView _gridLogs = new DataGridView();

        private ImportResult _lastResult;

        public FrmImport()
        {
            Text = "数据导入向导";
            Width = 980;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadImportLogs();
            UpdateActionState();
        }

        private void InitializeLayout()
        {
            _btnDownloadTemplate.Text = "下载模板(CSV/XLSX)";
            _btnDownloadTemplate.SetBounds(20, 20, 170, 30);
            _btnDownloadTemplate.Click += (s, e) => DownloadTemplate();
            Controls.Add(_btnDownloadTemplate);

            _btnImportFile.Text = "选择文件并导入";
            _btnImportFile.SetBounds(200, 20, 140, 30);
            _btnImportFile.Click += (s, e) => ImportFile();
            Controls.Add(_btnImportFile);

            _btnExportErrors.Text = "导出回填模板";
            _btnExportErrors.SetBounds(350, 20, 130, 30);
            _btnExportErrors.Click += (s, e) => ExportErrors();
            Controls.Add(_btnExportErrors);

            _btnCopyErrors.Text = "复制错误信息";
            _btnCopyErrors.SetBounds(490, 20, 130, 30);
            _btnCopyErrors.Click += (s, e) => CopyErrors();
            Controls.Add(_btnCopyErrors);

            _btnRefreshLogs.Text = "刷新导入日志";
            _btnRefreshLogs.SetBounds(630, 20, 130, 30);
            _btnRefreshLogs.Click += (s, e) => LoadImportLogs();
            Controls.Add(_btnRefreshLogs);

            _txtResult.SetBounds(20, 70, 920, 130);
            _txtResult.Multiline = true;
            _txtResult.ReadOnly = true;
            _txtResult.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_txtResult);

            _lblFailurePreview.Text = "失败数据预览（前20条，可直接按原列修正后回填）";
            _lblFailurePreview.SetBounds(20, 210, 450, 24);
            Controls.Add(_lblFailurePreview);

            _gridFailurePreview.SetBounds(20, 235, 920, 150);
            _gridFailurePreview.ReadOnly = true;
            _gridFailurePreview.AllowUserToAddRows = false;
            _gridFailurePreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_gridFailurePreview);

            _gridLogs.SetBounds(20, 395, 920, 165);
            _gridLogs.ReadOnly = true;
            _gridLogs.AllowUserToAddRows = false;
            _gridLogs.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_gridLogs);
        }

        private void DownloadTemplate()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel 文件|*.xlsx|CSV 文件|*.csv";
                dialog.FileName = "CleanGuard_导入模板.xlsx";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                string warning;
                string actualPath = ImportHelper.ExportTemplateWithFallback(dialog.FileName, out warning);
                string message = string.IsNullOrWhiteSpace(warning)
                    ? "模板下载成功。"
                    : warning + "\n文件路径：" + actualPath;
                MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel 文件|*.xlsx|CSV 文件|*.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _lastResult = ImportHelper.ImportFromFile(dialog.FileName);
                ShowImportResult(_lastResult);
                LoadImportLogs();
                UpdateActionState();
                DialogResult = DialogResult.OK;
            }
        }

        private void ExportErrors()
        {
            if (_lastResult == null || !_lastResult.Errors.Any())
            {
                MessageBox.Show("当前没有可导出的回填数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel 文件|*.xlsx|CSV 文件|*.csv";
                dialog.FileName = "ImportRefillTemplate.xlsx";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _lastResult.ExportErrors(dialog.FileName);
                MessageBox.Show("回填模板已导出，可修正后再次导入。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CopyErrors()
        {
            if (_lastResult == null || !_lastResult.Errors.Any())
            {
                MessageBox.Show("当前没有可复制的错误信息。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string text = string.Join(Environment.NewLine, _lastResult.Errors);
            Clipboard.SetText(text);
            MessageBox.Show("错误信息已复制到剪贴板。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void UpdateActionState()
        {
            bool hasErrors = _lastResult != null && _lastResult.Errors.Any();
            _btnExportErrors.Enabled = hasErrors;
            _btnCopyErrors.Enabled = hasErrors;
        }

        private void ShowImportResult(ImportResult result)
        {
            var lines = new[]
            {
                "导入完成",
                "文件: " + Path.GetFileName(result.SourceFile ?? string.Empty),
                "成功: " + result.SuccessCount,
                "失败: " + result.FailedCount,
                "时间: " + result.ImportTime.ToString("yyyy-MM-dd HH:mm:ss")
            }.ToList();

            if (result.Errors.Any())
            {
                lines.Add("---- 失败明细（前10条）----");
                lines.AddRange(result.Errors.Take(10));
                lines.Add("提示：可点击“导出回填模板”进行批量修正后再次导入。");
            }

            _txtResult.Text = string.Join(Environment.NewLine, lines);
            BindFailurePreview(result);
        }

        private void BindFailurePreview(ImportResult result)
        {
            if (result == null || result.FailedRows == null || result.FailedRows.Count == 0)
            {
                _gridFailurePreview.DataSource = null;
                return;
            }

            var table = new DataTable();
            table.Columns.Add("源行号");
            table.Columns.Add("工号");
            table.Columns.Add("姓名");
            table.Columns.Add("工序");
            table.Columns.Add("1F衣柜");
            table.Columns.Add("1F鞋柜");
            table.Columns.Add("2F衣柜");
            table.Columns.Add("2F鞋柜");
            table.Columns.Add("错误信息");

            int count = Math.Min(20, result.FailedRows.Count);
            for (int i = 0; i < count; i++)
            {
                var row = result.FailedRows[i];
                table.Rows.Add(
                    row.RowNumber,
                    row.EmpNo,
                    row.Name,
                    row.Process,
                    row.Locker1FClothes,
                    row.Locker1FShoe,
                    row.Locker2FClothes,
                    row.Locker2FShoe,
                    row.Error);
            }

            _gridFailurePreview.DataSource = table;
        }

        private void LoadImportLogs()
        {
            DataTable logs = SQLiteHelper.QuerySystemLogs("Import", 50);
            _gridLogs.DataSource = logs;
        }
    }
}
