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
        private readonly Button _btnImportCsv = new Button();
        private readonly Button _btnExportErrors = new Button();
        private readonly Button _btnRefreshLogs = new Button();
        private readonly TextBox _txtResult = new TextBox();
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
        }

        private void InitializeLayout()
        {
            _btnDownloadTemplate.Text = "下载CSV模板";
            _btnDownloadTemplate.SetBounds(20, 20, 130, 30);
            _btnDownloadTemplate.Click += (s, e) => DownloadTemplate();
            Controls.Add(_btnDownloadTemplate);

            _btnImportCsv.Text = "选择CSV并导入";
            _btnImportCsv.SetBounds(160, 20, 140, 30);
            _btnImportCsv.Click += (s, e) => ImportCsv();
            Controls.Add(_btnImportCsv);

            _btnExportErrors.Text = "导出失败明细";
            _btnExportErrors.SetBounds(310, 20, 130, 30);
            _btnExportErrors.Click += (s, e) => ExportErrors();
            Controls.Add(_btnExportErrors);

            _btnRefreshLogs.Text = "刷新导入日志";
            _btnRefreshLogs.SetBounds(450, 20, 130, 30);
            _btnRefreshLogs.Click += (s, e) => LoadImportLogs();
            Controls.Add(_btnRefreshLogs);

            _txtResult.SetBounds(20, 70, 920, 130);
            _txtResult.Multiline = true;
            _txtResult.ReadOnly = true;
            _txtResult.ScrollBars = ScrollBars.Vertical;
            Controls.Add(_txtResult);

            _gridLogs.SetBounds(20, 220, 920, 340);
            _gridLogs.ReadOnly = true;
            _gridLogs.AllowUserToAddRows = false;
            _gridLogs.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_gridLogs);
        }

        private void DownloadTemplate()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                dialog.FileName = "CleanGuard_导入模板.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                ImportHelper.ExportTemplateCsv(dialog.FileName);
                MessageBox.Show("模板下载成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportCsv()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _lastResult = ImportHelper.ImportFromCsv(dialog.FileName);
                ShowImportResult(_lastResult);
                LoadImportLogs();
                DialogResult = DialogResult.OK;
            }
        }

        private void ExportErrors()
        {
            if (_lastResult == null || !_lastResult.Errors.Any())
            {
                MessageBox.Show("当前没有可导出的失败明细。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                dialog.FileName = "ImportErrors.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _lastResult.ExportErrorsCsv(dialog.FileName);
                MessageBox.Show("失败明细已导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            }

            _txtResult.Text = string.Join(Environment.NewLine, lines);
        }

        private void LoadImportLogs()
        {
            DataTable logs = SQLiteHelper.QuerySystemLogs("Import", 50);
            _gridLogs.DataSource = logs;
        }
    }
}
