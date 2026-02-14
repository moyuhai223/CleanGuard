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
        private readonly ComboBox _cmbOpType = new ComboBox();
        private readonly CheckBox _chkDateRange = new CheckBox();
        private readonly DateTimePicker _dtFrom = new DateTimePicker();
        private readonly DateTimePicker _dtTo = new DateTimePicker();
        private readonly TextBox _txtKeyword = new TextBox();
        private readonly Button _btnRefresh = new Button();
        private readonly Button _btnClearFilters = new Button();
        private readonly Button _btnExport = new Button();
        private readonly DataGridView _grid = new DataGridView();
        private readonly Label _lblStats = new Label();

        private DataTable _rawTable;

        public FrmProcessAudit()
        {
            Text = "工序字典审计视图";
            Width = 1160;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            ToggleDateRange();
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

            Controls.Add(new Label { Text = "操作类型", Left = 180, Top = 24, Width = 60 });
            _cmbOpType.SetBounds(245, 20, 130, 28);
            _cmbOpType.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbOpType.Items.AddRange(new object[] { "全部", "新增", "删除", "重命名", "批量导入" });
            _cmbOpType.SelectedIndex = 0;
            _cmbOpType.SelectedIndexChanged += (s, e) => ApplyFilter();
            Controls.Add(_cmbOpType);

            _chkDateRange.Text = "启用日期";
            _chkDateRange.SetBounds(390, 22, 80, 24);
            _chkDateRange.CheckedChanged += (s, e) => ToggleDateRange();
            Controls.Add(_chkDateRange);

            Controls.Add(new Label { Text = "起始日期", Left = 470, Top = 24, Width = 60 });
            _dtFrom.SetBounds(535, 20, 110, 28);
            _dtFrom.Format = DateTimePickerFormat.Custom;
            _dtFrom.CustomFormat = "yyyy-MM-dd";
            _dtFrom.ValueChanged += (s, e) => ApplyFilter();
            Controls.Add(_dtFrom);

            Controls.Add(new Label { Text = "结束日期", Left = 655, Top = 24, Width = 60 });
            _dtTo.SetBounds(720, 20, 110, 28);
            _dtTo.Format = DateTimePickerFormat.Custom;
            _dtTo.CustomFormat = "yyyy-MM-dd";
            _dtTo.ValueChanged += (s, e) => ApplyFilter();
            Controls.Add(_dtTo);

            Controls.Add(new Label { Text = "关键字", Left = 840, Top = 24, Width = 50 });
            _txtKeyword.SetBounds(895, 20, 120, 28);
            _txtKeyword.TextChanged += (s, e) => ApplyFilter();
            Controls.Add(_txtKeyword);

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(20, 530, 70, 28);
            _btnRefresh.Click += (s, e) => LoadLogs();
            Controls.Add(_btnRefresh);

            _btnClearFilters.Text = "清除筛选";
            _btnClearFilters.SetBounds(100, 530, 90, 28);
            _btnClearFilters.Click += (s, e) => ResetFilters();
            Controls.Add(_btnClearFilters);

            _btnExport.Text = "导出CSV";
            _btnExport.SetBounds(200, 530, 90, 28);
            _btnExport.Click += (s, e) => ExportCsv();
            Controls.Add(_btnExport);

            _grid.SetBounds(20, 90, 1120, 430);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);

            _lblStats.SetBounds(20, 66, 1120, 20);
            _lblStats.Text = "当前结果：0";
            Controls.Add(_lblStats);
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

            DataView view = _rawTable.DefaultView;
            var conditions = new StringBuilder();

            string opType = Convert.ToString(_cmbOpType.SelectedItem ?? "全部");
            if (!string.Equals(opType, "全部", StringComparison.Ordinal))
            {
                if (conditions.Length > 0) conditions.Append(" AND ");
                conditions.Append(BuildOpTypeCondition(opType));
            }

            if (_chkDateRange.Checked)
            {
                DateTime fromDate = _dtFrom.Value.Date;
                DateTime toDate = _dtTo.Value.Date.AddDays(1);
                if (fromDate < toDate)
                {
                    if (conditions.Length > 0) conditions.Append(" AND ");
                    conditions.AppendFormat("时间 >= '{0}' AND 时间 < '{1}'", fromDate.ToString("yyyy-MM-dd"), toDate.ToString("yyyy-MM-dd"));
                }
            }

            string keyword = (_txtKeyword.Text ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(keyword))
            {
                if (conditions.Length > 0) conditions.Append(" AND ");
                string escaped = keyword.Replace("'", "''");
                conditions.AppendFormat("内容 LIKE '%{0}%'", escaped);
            }

            view.RowFilter = conditions.ToString();
            _grid.DataSource = view;
            UpdateStats(view);
        }


        private void ToggleDateRange()
        {
            _dtFrom.Enabled = _chkDateRange.Checked;
            _dtTo.Enabled = _chkDateRange.Checked;
            ApplyFilter();
        }

        private void ResetFilters()
        {
            _cmbOpType.SelectedIndex = 0;
            _chkDateRange.Checked = false;
            _txtKeyword.Clear();
            ApplyFilter();
        }


        private void UpdateStats(DataView view)
        {
            int total = view == null ? 0 : view.Count;
            int added = CountByPrefix(view, "新增工序字典:");
            int deleted = CountByPrefix(view, "删除工序字典:");
            int renamed = CountByPrefix(view, "重命名工序字典:");
            int imported = CountByPrefix(view, "工序批量导入完成");
            _lblStats.Text = string.Format("当前结果：{0}（新增 {1}，删除 {2}，重命名 {3}，批量导入 {4}）", total, added, deleted, renamed, imported);
        }

        private static int CountByPrefix(DataView view, string prefix)
        {
            if (view == null || string.IsNullOrWhiteSpace(prefix))
            {
                return 0;
            }

            int count = 0;
            foreach (DataRowView row in view)
            {
                string message = Convert.ToString(row["内容"]);
                if (!string.IsNullOrWhiteSpace(message) && message.StartsWith(prefix, StringComparison.Ordinal))
                {
                    count++;
                }
            }

            return count;
        }

        private static string BuildOpTypeCondition(string opType)
        {
            if (opType == "新增") return "内容 LIKE '新增工序字典:%'";
            if (opType == "删除") return "内容 LIKE '删除工序字典:%'";
            if (opType == "重命名") return "内容 LIKE '重命名工序字典:%'";
            if (opType == "批量导入") return "内容 LIKE '工序批量导入完成%'";
            return string.Empty;
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
