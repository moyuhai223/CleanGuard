using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmProcessAudit : Form
    {
        private readonly NumericUpDown _numLimit = new NumericUpDown();
        private readonly TextBox _txtKeyword = new TextBox();
        private readonly Button _btnRefresh = new Button();
        private readonly Button _btnExport = new Button();
        private readonly DataGridView _grid = new DataGridView();

        private DataTable _rawTable;

        public FrmProcessAudit()
        {
            Text = "工序字典审计视图";
            Width = 980;
            Height = 580;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadLogs();
        }

        private void InitializeLayout()
        {
            Controls.Add(new Label { Text = "显示条数", Left = 20, Top = 24, Width = 60 });

            _numLimit.SetBounds(85, 20, 80, 28);
            _numLimit.Minimum = 10;
            _numLimit.Maximum = 500;
            _numLimit.Value = 100;
            Controls.Add(_numLimit);

            Controls.Add(new Label { Text = "关键字", Left = 180, Top = 24, Width = 50 });
            _txtKeyword.SetBounds(235, 20, 220, 28);
            _txtKeyword.TextChanged += (s, e) => ApplyFilter();
            Controls.Add(_txtKeyword);

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(470, 20, 80, 28);
            _btnRefresh.Click += (s, e) => LoadLogs();
            Controls.Add(_btnRefresh);

            _btnExport.Text = "导出CSV";
            _btnExport.SetBounds(560, 20, 90, 28);
            _btnExport.Click += (s, e) => ExportCsv();
            Controls.Add(_btnExport);

            _grid.SetBounds(20, 65, 920, 460);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);
        }

        private void LoadLogs()
        {
            _rawTable = SQLiteHelper.QueryProcessAuditLogs((int)_numLimit.Value);
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_rawTable == null)
            {
                return;
            }

            string keyword = (_txtKeyword.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                _grid.DataSource = _rawTable;
                return;
            }

            DataView view = _rawTable.DefaultView;
            string escaped = keyword.Replace("'", "''");
            view.RowFilter = string.Format("内容 LIKE '%{0}%'", escaped);
            _grid.DataSource = view;
        }

        private void ExportCsv()
        {
            DataTable table = null;
            var view = _grid.DataSource as DataView;
            if (view != null)
            {
                table = view.ToTable();
            }
            else
            {
                table = _grid.DataSource as DataTable;
            }

            if (table == null || table.Rows.Count == 0)
            {
                MessageBox.Show("当前无可导出的审计数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                dialog.FileName = "ProcessAudit.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var sb = new StringBuilder();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(EscapeCsv(table.Columns[i].ColumnName));
                }
                sb.AppendLine();

                foreach (DataRow row in table.Rows)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        if (i > 0) sb.Append(',');
                        sb.Append(EscapeCsv(Convert.ToString(row[i])));
                    }
                    sb.AppendLine();
                }

                File.WriteAllText(dialog.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("审计数据导出成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string EscapeCsv(string value)
        {
            string text = value ?? string.Empty;
            text = text.Replace("\"", "\"\"");
            return "\"" + text + "\"";
        }
    }
}
