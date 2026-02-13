using System;
using System.Data;
using System.Drawing;
using System.Linq;
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
        private readonly Button _btnDownloadTemplate = new Button();
        private readonly Button _btnImportCsv = new Button();
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
            _txtSearch.SetBounds(20, 20, 220, 30);
            Controls.Add(_txtSearch);

            _btnSearch.Text = "搜索";
            _btnSearch.SetBounds(250, 20, 70, 30);
            _btnSearch.Click += (s, e) => LoadEmployeeData(_txtSearch.Text.Trim());
            Controls.Add(_btnSearch);

            _btnAdd.Text = "新增员工";
            _btnAdd.SetBounds(330, 20, 95, 30);
            _btnAdd.Click += (s, e) => OpenEditor();
            Controls.Add(_btnAdd);

            _btnResign.Text = "办理离职";
            _btnResign.SetBounds(435, 20, 95, 30);
            _btnResign.Click += (s, e) => ResignSelectedEmployee();
            Controls.Add(_btnResign);

            _btnDownloadTemplate.Text = "下载导入模板";
            _btnDownloadTemplate.SetBounds(540, 20, 110, 30);
            _btnDownloadTemplate.Click += (s, e) => DownloadImportTemplate();
            Controls.Add(_btnDownloadTemplate);

            _btnImportCsv.Text = "导入数据";
            _btnImportCsv.SetBounds(660, 20, 90, 30);
            _btnImportCsv.Click += (s, e) => ImportFromCsv();
            Controls.Add(_btnImportCsv);

            _btnLockerMap.Text = "柜位分布图";
            _btnLockerMap.SetBounds(760, 20, 120, 30);
            _btnLockerMap.Click += (s, e) => ShowLockerHeatmapPlaceholder();
            Controls.Add(_btnLockerMap);

            _grid.SetBounds(20, 70, 1220, 590);
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

        private void DownloadImportTemplate()
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

        private void ImportFromCsv()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var result = ImportHelper.ImportFromCsv(dialog.FileName);
                LoadEmployeeData(_txtSearch.Text.Trim());

                string msg = $"导入完成。\n成功: {result.SuccessCount}\n失败: {result.FailedCount}";
                if (result.Errors.Any())
                {
                    msg += "\n\n失败明细（最多显示前5条）：\n" + string.Join("\n", result.Errors.Take(5));
                }

                MessageBox.Show(msg, "导入结果", MessageBoxButtons.OK,
                    result.FailedCount == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
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
