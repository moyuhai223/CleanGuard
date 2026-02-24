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
        private readonly Button _btnDelete = new Button();
        private readonly Button _btnPrintLabel = new Button();
        private readonly Button _btnLockerMap = new Button();
        private readonly Button _btnImport = new Button();
        private readonly Button _btnLogs = new Button();
        private readonly Button _btnProcessDict = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmMain()
        {
            Text = "CleanGuard 劳保与更衣柜管理系统";
            Width = 1280;
            Height = 720;
            StartPosition = FormStartPosition.CenterScreen;

            UiTheme.ApplyFormStyle(this);

            InitializeLayout();
            LoadEmployeeData();
        }

        private void InitializeLayout()
        {
            _txtSearch.SetBounds(20, 20, 280, 32);
            Controls.Add(_txtSearch);

            _btnSearch.Text = "搜索";
            _btnSearch.SetBounds(310, 20, 90, 32);
            _btnSearch.Click += (s, e) => LoadEmployeeData(_txtSearch.Text.Trim());
            Controls.Add(_btnSearch);
            UiTheme.StylePrimaryButton(_btnSearch);

            var grpDailyOps = CreateActionGroup("步骤1：日常员工操作", 20, 65, 580, 70, Color.FromArgb(69, 126, 245));
            Controls.Add(grpDailyOps);

            _btnAdd.Text = "新增员工";
            _btnAdd.Click += (s, e) => OpenEditor();
            grpDailyOps.Controls.Add(_btnAdd);
            UiTheme.StylePrimaryButton(_btnAdd);

            _btnEdit.Text = "编辑员工";
            _btnEdit.Click += (s, e) => OpenEditorForSelected();
            grpDailyOps.Controls.Add(_btnEdit);
            UiTheme.StylePrimaryButton(_btnEdit);

            _btnRestore.Text = "办理复职";
            _btnRestore.Click += (s, e) => RestoreSelectedEmployee();
            grpDailyOps.Controls.Add(_btnRestore);
            UiTheme.StylePrimaryButton(_btnRestore);

            _btnPrintLabel.Text = "打印标签";
            _btnPrintLabel.Click += (s, e) => PrintSelectedEmployeeLabel();
            grpDailyOps.Controls.Add(_btnPrintLabel);
            UiTheme.StylePrimaryButton(_btnPrintLabel);

            LayoutActionButtons(grpDailyOps, 24, 108, 32);

            var grpDangerOps = CreateActionGroup("步骤2：高风险操作（请谨慎）", 610, 65, 290, 70, Color.FromArgb(200, 80, 80));
            Controls.Add(grpDangerOps);

            _btnResign.Text = "办理离职";
            _btnResign.Click += (s, e) => ResignSelectedEmployee();
            grpDangerOps.Controls.Add(_btnResign);
            UiTheme.StyleWarningButton(_btnResign);

            _btnDelete.Text = "删除员工";
            _btnDelete.Click += (s, e) => DeleteSelectedEmployee();
            grpDangerOps.Controls.Add(_btnDelete);
            UiTheme.StyleWarningButton(_btnDelete);

            LayoutActionButtons(grpDangerOps, 24, 122, 32);

            var grpModuleNav = CreateActionGroup("步骤3：进入业务模块", 910, 65, 370, 70, Color.FromArgb(69, 126, 245));
            Controls.Add(grpModuleNav);

            _btnImport.Text = "数据导入";
            _btnImport.Click += (s, e) => OpenImportForm();
            grpModuleNav.Controls.Add(_btnImport);
            UiTheme.StylePrimaryButton(_btnImport);

            _btnLockerMap.Text = "柜位分布图";
            _btnLockerMap.Click += (s, e) => OpenLockerChart();
            grpModuleNav.Controls.Add(_btnLockerMap);
            UiTheme.StylePrimaryButton(_btnLockerMap);

            _btnLogs.Text = "系统日志";
            _btnLogs.Click += (s, e) => OpenSystemLogs();
            grpModuleNav.Controls.Add(_btnLogs);
            UiTheme.StylePrimaryButton(_btnLogs);

            _btnProcessDict.Text = "工序字典";
            _btnProcessDict.Click += (s, e) => OpenProcessDict();
            grpModuleNav.Controls.Add(_btnProcessDict);
            UiTheme.StylePrimaryButton(_btnProcessDict);

            LayoutActionButtons(grpModuleNav, 24, 82, 32);

            _grid.SetBounds(20, 145, 1260, 515);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.CellDoubleClick += (s, e) => OpenEditorForSelected();
            Controls.Add(_grid);
            UiTheme.StyleDataGrid(_grid);
        }

        private static GroupBox CreateActionGroup(string title, int left, int top, int width, int height, Color titleColor)
        {
            var group = new GroupBox
            {
                Text = title,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                ForeColor = titleColor
            };

            return group;
        }

        private static void LayoutActionButtons(GroupBox group, int top, int width, int height)
        {
            int left = 12;
            foreach (Control control in group.Controls)
            {
                var button = control as Button;
                if (button == null)
                {
                    continue;
                }

                button.SetBounds(left, top, width, height);
                left += width + 10;
            }
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
                var mode = MessageBox.Show("选择打印方式：\n是 = 打印预览\n否 = 直接打印\n取消 = 取消", "打印方式", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (mode == DialogResult.Cancel)
                {
                    return;
                }

                bool printed = false;
                if (mode == DialogResult.Yes)
                {
                    Printer.ShowLabelPreview(empNo, name, process, locker2F);
                    printed = true;
                }
                else
                {
                    printed = Printer.PrintLabelDirect(this, empNo, name, process, locker2F);
                }

                if (printed)
                {
                    SQLiteHelper.WriteSystemLog("Print", $"打印员工标签: {empNo}-{name}, 二维码内容={Printer.BuildQrPayload(empNo, name, locker2F)}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打印失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void DeleteSelectedEmployee()
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
                $"确认彻底删除员工 {info.Name} ({info.EmpNo}) 吗？\n" +
                "该操作不可恢复，并将释放以下资源：\n" +
                $"1F衣柜: {DisplayLocker(info.Locker1FClothes)}\n" +
                $"1F鞋柜: {DisplayLocker(info.Locker1FShoe)}\n" +
                $"2F衣柜: {DisplayLocker(info.Locker2FClothes)}\n" +
                $"2F鞋柜: {DisplayLocker(info.Locker2FShoe)}";

            var result = MessageBox.Show("【高风险操作】\n" + message + "\n\n请再次确认：此操作会永久删除员工数据。", "删除确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SQLiteHelper.DeleteEmployee(empNo);
                LoadEmployeeData(_txtSearch.Text.Trim());
                MessageBox.Show("员工删除成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            var result = MessageBox.Show("【高风险操作】\n" + message + "\n\n请确认是否继续办理离职。", "离职确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
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


        private void OpenProcessDict()
        {
            using (var form = new FrmProcessManage())
            {
                form.ShowDialog(this);
            }

            LoadEmployeeData(_txtSearch.Text.Trim());
        }
    }
}
