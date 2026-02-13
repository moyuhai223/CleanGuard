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
        private readonly Button _btnEdit = new Button();
        private readonly Button _btnResign = new Button();
        private readonly Button _btnRestore = new Button();
        private readonly Button _btnPrintLabel = new Button();
        private readonly Button _btnLockerMap = new Button();
        private readonly Button _btnImport = new Button();
        private readonly Button _btnLogs = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmMain()
        {
            Text = "CleanGuard 劳保与更衣柜管理系统";
            Width = 1280;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            InitializeLayout();
            LoadEmployeeData();
        }

        private void InitializeLayout()
        {
            _txtSearch.SetBounds(20, 20, 200, 30);
            Controls.Add(_txtSearch);

            _btnSearch.Text = "搜索";
            _btnSearch.SetBounds(230, 20, 70, 30);
            _btnSearch.Click += (s, e) => LoadEmployeeData(_txtSearch.Text.Trim());
            Controls.Add(_btnSearch);

            _btnAdd.Text = "新增员工";
            _btnAdd.SetBounds(310, 20, 90, 30);
            _btnAdd.Click += (s, e) => OpenEditor();
            Controls.Add(_btnAdd);

            _btnEdit.Text = "编辑员工";
            _btnEdit.SetBounds(410, 20, 90, 30);
            _btnEdit.Click += (s, e) => OpenEditorForSelected();
            Controls.Add(_btnEdit);

            _btnResign.Text = "办理离职";
            _btnResign.SetBounds(510, 20, 90, 30);
            _btnResign.Click += (s, e) => ResignSelectedEmployee();
            Controls.Add(_btnResign);

            _btnRestore.Text = "办理复职";
            _btnRestore.SetBounds(610, 20, 90, 30);
            _btnRestore.Click += (s, e) => RestoreSelectedEmployee();
            Controls.Add(_btnRestore);

            _btnPrintLabel.Text = "打印标签";
            _btnPrintLabel.SetBounds(710, 20, 90, 30);
            _btnPrintLabel.Click += (s, e) => PrintSelectedEmployeeLabel();
            Controls.Add(_btnPrintLabel);

            _btnImport.Text = "数据导入";
            _btnImport.SetBounds(810, 20, 100, 30);
            _btnImport.Click += (s, e) => OpenImportForm();
            Controls.Add(_btnImport);

            _btnLockerMap.Text = "柜位分布图";
            _btnLockerMap.SetBounds(920, 20, 120, 30);
            _btnLockerMap.Click += (s, e) => OpenLockerChart();
            Controls.Add(_btnLockerMap);

            _btnLogs.Text = "系统日志";
            _btnLogs.SetBounds(1050, 20, 90, 30);
            _btnLogs.Click += (s, e) => OpenSystemLogs();
            Controls.Add(_btnLogs);

            _grid.SetBounds(20, 70, 1220, 590);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.CellDoubleClick += (s, e) => OpenEditorForSelected();
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

        private void OpenEditorForSelected()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择一名员工。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string empNo = Convert.ToString(_grid.SelectedRows[0].Cells["工号"].Value);
            if (string.IsNullOrWhiteSpace(empNo))
            {
                MessageBox.Show("无法识别当前行工号。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var editor = new FrmEditor(empNo))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    LoadEmployeeData(_txtSearch.Text.Trim());
                }
            }
        }

        private void OpenImportForm()
        {
            using (var form = new FrmImport())
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    LoadEmployeeData(_txtSearch.Text.Trim());
                }
            }
        }

        private void PrintSelectedEmployeeLabel()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择一名员工。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var row = _grid.SelectedRows[0];
            string empNo = Convert.ToString(row.Cells["工号"].Value);
            string name = Convert.ToString(row.Cells["姓名"].Value);
            string process = Convert.ToString(row.Cells["工序"].Value);
            string locker2F = Convert.ToString(row.Cells["2F衣柜"].Value);

            if (string.IsNullOrWhiteSpace(empNo) || string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("当前选中数据不完整，无法打印。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Printer.ShowLabelPreview(empNo, name, process, locker2F);
                SQLiteHelper.WriteSystemLog("Print", $"打印员工标签: {empNo}-{name}, 二维码内容={Printer.BuildQrPayload(empNo, name, locker2F)}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打印预览失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RestoreSelectedEmployee()
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

            if (info.Status == 1)
            {
                MessageBox.Show("该员工当前已是在职状态。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show($"确认将 {info.Name} 办理复职吗？", "复职确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SQLiteHelper.RestoreEmployee(empNo);
                LoadEmployeeData(_txtSearch.Text.Trim());
                MessageBox.Show("复职办理成功。可通过编辑员工重新分配柜位。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "复职失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void OpenLockerChart()
        {
            using (var form = new FrmLockerChart())
            {
                form.ShowDialog(this);
            }
        }

        private void OpenSystemLogs()
        {
            using (var form = new FrmSystemLog())
            {
                form.ShowDialog(this);
            }
        }
    }
}
