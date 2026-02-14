using System;
using System.Data;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmProcessAudit : Form
    {
        private readonly NumericUpDown _numLimit = new NumericUpDown();
        private readonly Button _btnRefresh = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmProcessAudit()
        {
            Text = "工序字典审计视图";
            Width = 900;
            Height = 560;
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

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(180, 20, 80, 28);
            _btnRefresh.Click += (s, e) => LoadLogs();
            Controls.Add(_btnRefresh);

            _grid.SetBounds(20, 65, 840, 440);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);
        }

        private void LoadLogs()
        {
            DataTable table = SQLiteHelper.QueryProcessAuditLogs((int)_numLimit.Value);
            _grid.DataSource = table;
        }
    }
}
