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
        private readonly Button _btnAdd = new Button();
        private readonly Button _btnResign = new Button();
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
            _txtSearch.SetBounds(20, 20, 250, 30);
            Controls.Add(_txtSearch);

            _btnSearch.Text = "搜索";
            _btnSearch.SetBounds(280, 20, 80, 30);
            _btnSearch.Click += (s, e) => LoadEmployeeData(_txtSearch.Text.Trim());
            Controls.Add(_btnSearch);

            _btnAdd.Text = "新增员工";
            _btnAdd.SetBounds(370, 20, 100, 30);
            _btnAdd.Click += (s, e) => OpenEditor();
            Controls.Add(_btnAdd);

            _btnResign.Text = "办理离职";
            _btnResign.SetBounds(480, 20, 100, 30);
            _btnResign.Click += (s, e) => ResignSelectedEmployee();
            Controls.Add(_btnResign);

            _btnLockerMap.Text = "柜位分布图";
            _btnLockerMap.SetBounds(590, 20, 120, 30);
            _btnLockerMap.Click += (s, e) => ShowLockerHeatmapPlaceholder();
            Controls.Add(_btnLockerMap);

            _grid.SetBounds(20, 70, 1140, 560);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
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

        private void OpenEditor()
        {
            using (var editor = new FrmEditor())
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    LoadEmployeeData(_txtSearch.Text.Trim());
                }
            }
        }

        private void ResignSelectedEmployee()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择一名员工。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string empNo = Convert.ToString(_grid.SelectedRows[0].Cells["工号"].Value);
            var info = SQLiteHelper.GetEmployeeLockerInfo(empNo);
            if (info == null)
            {
                MessageBox.Show("未找到该员工。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string message =
                $"确认 {info.Name} 离职吗？这将释放以下资源：\n" +
                $"1F衣柜: {DisplayLocker(info.Locker1FClothes)}\n" +
                $"1F鞋柜: {DisplayLocker(info.Locker1FShoe)}\n" +
                $"2F衣柜: {DisplayLocker(info.Locker2FClothes)}\n" +
                $"2F鞋柜: {DisplayLocker(info.Locker2FShoe)}";

            var result = MessageBox.Show(message, "离职确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SQLiteHelper.MarkEmployeeResigned(empNo);
                LoadEmployeeData(_txtSearch.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "离职失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string DisplayLocker(string locker)
        {
            return string.IsNullOrWhiteSpace(locker) ? "(空)" : locker;
        }

        private void ShowLockerHeatmapPlaceholder()
        {
            var summary = SQLiteHelper.GetLockerSummary();
            string message =
                "当前柜位占用统计（简版）\n\n" +
                $"1F 衣柜：{summary.OneFClothesOccupied}/{summary.OneFClothesTotal}\n" +
                $"1F 鞋柜：{summary.OneFShoeOccupied}/{summary.OneFShoeTotal}\n" +
                $"2F 衣柜：{summary.TwoFClothesOccupied}/{summary.TwoFClothesTotal}\n" +
                $"2F 鞋柜：{summary.TwoFShoeOccupied}/{summary.TwoFShoeTotal}";

            MessageBox.Show(message, "柜位分布图（统计版）", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
