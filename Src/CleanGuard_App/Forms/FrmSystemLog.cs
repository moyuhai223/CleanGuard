using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmSystemLog : Form
    {
        private readonly ComboBox _cmbType = new ComboBox();
        private readonly NumericUpDown _numLimit = new NumericUpDown();
        private readonly Button _btnQuery = new Button();
        private readonly Button _btnExport = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmSystemLog()
        {
            Text = "系统日志";
            Width = 960;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadLogs();
        }

        private void InitializeLayout()
        {
            _cmbType.SetBounds(20, 20, 140, 28);
            _cmbType.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbType.Items.AddRange(new object[] { "全部", "Import", "Backup", "Print", "Employee" });
            _cmbType.SelectedIndex = 0;
            Controls.Add(_cmbType);

            _numLimit.SetBounds(180, 20, 80, 28);
            _numLimit.Minimum = 10;
            _numLimit.Maximum = 500;
            _numLimit.Value = 100;
            Controls.Add(_numLimit);

            _btnQuery.Text = "查询";
            _btnQuery.SetBounds(280, 20, 80, 28);
            _btnQuery.Click += (s, e) => LoadLogs();
            Controls.Add(_btnQuery);

            _btnExport.Text = "导出CSV";
            _btnExport.SetBounds(370, 20, 100, 28);
            _btnExport.Click += (s, e) => ExportCsv();
            Controls.Add(_btnExport);

            _grid.SetBounds(20, 60, 900, 500);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);
        }

        private void LoadLogs()
        {
            string type = _cmbType.SelectedItem != null && _cmbType.SelectedItem.ToString() != "全部"
                ? _cmbType.SelectedItem.ToString()
                : string.Empty;
            int limit = Convert.ToInt32(_numLimit.Value);
            DataTable dt = SQLiteHelper.QuerySystemLogs(type, limit);
            _grid.DataSource = dt;
        }

        private void ExportCsv()
        {
            var dt = _grid.DataSource as DataTable;
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("当前无可导出的日志。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                dialog.FileName = "SystemLogs.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var sb = new StringBuilder();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append('"').Append(dt.Columns[i].ColumnName.Replace("\"", "\"\"")).Append('"');
                }
                sb.AppendLine();

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (i > 0) sb.Append(',');
                        var cell = Convert.ToString(row[i]) ?? string.Empty;
                        sb.Append('"').Append(cell.Replace("\"", "\"\"")).Append('"');
                    }
                    sb.AppendLine();
                }

                File.WriteAllText(dialog.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("导出成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
