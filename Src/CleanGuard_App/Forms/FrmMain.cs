using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmMain : Form
    {
        private readonly TextBox _txtSearch = new TextBox();
        private readonly Button _btnSearch = new Button();
        private readonly Button _btnLockerMap = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmMain()
        {
            Text = "CleanGuard 劳保与更衣柜管理系统";
            Width = 1200;
            Height = 700;
            StartPosition = FormStartPosition.CenterScreen;

            InitializeLayout();
            LoadEmployeeData();
        }

        private void InitializeLayout()
        {
            _txtSearch.SetBounds(20, 20, 300, 30);
            Controls.Add(_txtSearch);

            _btnSearch.Text = "搜索";
            _btnSearch.SetBounds(330, 20, 80, 30);
            _btnSearch.Click += (s, e) => LoadEmployeeData(_txtSearch.Text.Trim());
            Controls.Add(_btnSearch);

            _btnLockerMap.Text = "柜位分布图";
            _btnLockerMap.SetBounds(420, 20, 120, 30);
            _btnLockerMap.Click += (s, e) => ShowLockerHeatmapPlaceholder();
            Controls.Add(_btnLockerMap);

            _grid.SetBounds(20, 70, 1140, 560);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);
        }

        private void LoadEmployeeData(string keyword = "")
        {
            DataTable table = SQLiteHelper.QueryEmployees(keyword);
            _grid.DataSource = table;

            if (_grid.Columns.Count > 0)
            {
                _grid.Columns[0].Frozen = true;
                _grid.Columns[0].DefaultCellStyle.BackColor = Color.AliceBlue;
            }
        }

        private void ShowLockerHeatmapPlaceholder()
        {
            MessageBox.Show("V1 开发阶段：柜位分布图将在后续版本接入图表组件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
