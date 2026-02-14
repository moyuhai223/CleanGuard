using System;
using System.Data;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmProcessManage : Form
    {
        private readonly DataGridView _grid = new DataGridView();
        private readonly TextBox _txtProcess = new TextBox();
        private readonly Button _btnAdd = new Button();
        private readonly Button _btnDelete = new Button();
        private readonly Button _btnRename = new Button();
        private readonly Button _btnRefresh = new Button();
        private readonly Button _btnImport = new Button();
        private readonly Button _btnAudit = new Button();
        private readonly Button _btnLocker = new Button();

        public FrmProcessManage()
        {
            Text = "工序字典维护";
            Width = 680;
            Height = 520;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadProcesses();
        }

        private void InitializeLayout()
        {
            Controls.Add(new Label { Text = "工序名称", Left = 20, Top = 25, Width = 60 });
            _txtProcess.SetBounds(90, 20, 180, 28);
            Controls.Add(_txtProcess);

            _btnAdd.Text = "新增工序";
            _btnAdd.SetBounds(290, 20, 90, 28);
            _btnAdd.Click += (s, e) => AddProcess();
            Controls.Add(_btnAdd);

            _btnDelete.Text = "删除选中";
            _btnDelete.SetBounds(390, 20, 90, 28);
            _btnDelete.Click += (s, e) => DeleteSelected();
            Controls.Add(_btnDelete);

            _btnRename.Text = "重命名";
            _btnRename.SetBounds(490, 20, 80, 28);
            _btnRename.Click += (s, e) => RenameSelected();
            Controls.Add(_btnRename);

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(580, 20, 60, 28);
            _btnRefresh.Click += (s, e) => LoadProcesses();
            Controls.Add(_btnRefresh);

            _btnImport.Text = "批量导入";
            _btnImport.SetBounds(20, 460, 90, 28);
            _btnImport.Click += (s, e) => ImportCsv();
            Controls.Add(_btnImport);

            _btnAudit.Text = "审计视图";
            _btnAudit.SetBounds(120, 460, 90, 28);
            _btnAudit.Click += (s, e) => OpenAudit();
            Controls.Add(_btnAudit);

            _btnLocker.Text = "柜位异常";
            _btnLocker.SetBounds(220, 460, 90, 28);
            _btnLocker.Click += (s, e) => OpenLockerManage();
            Controls.Add(_btnLocker);

            _grid.SetBounds(20, 65, 620, 385);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.SelectionChanged += (s, e) => FillSelectedToInput();
            Controls.Add(_grid);
        }

        private void LoadProcesses()
        {
            DataTable table = SQLiteHelper.QueryProcessesTable();
            _grid.DataSource = table;
        }



        private void ImportCsv()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv|文本文件|*.txt";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    var result = SQLiteHelper.ImportProcessesFromCsv(dialog.FileName);
                    LoadProcesses();
                    var message = string.Format("导入完成。成功 {0}，跳过 {1}，失败 {2}", result.SuccessCount, result.SkippedCount, result.FailedCount);
                    if (result.Errors.Count > 0)
                    {
                        message += Environment.NewLine + Environment.NewLine +
                                   "失败明细（前5条）：" + Environment.NewLine +
                                   string.Join(Environment.NewLine, result.Errors.GetRange(0, Math.Min(5, result.Errors.Count)).ToArray());
                    }

                    MessageBox.Show(message, "导入结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }


        private void OpenLockerManage()
        {
            using (var form = new FrmLockerManage())
            {
                form.ShowDialog(this);
            }
        }

        private void OpenAudit()
        {
            using (var form = new FrmProcessAudit())
            {
                form.ShowDialog(this);
            }
        }

        private void FillSelectedToInput()
        {
            if (_grid.SelectedRows.Count <= 0)
            {
                return;
            }

            string name = Convert.ToString(_grid.SelectedRows[0].Cells["工序名称"].Value);
            if (!string.IsNullOrWhiteSpace(name))
            {
                _txtProcess.Text = name;
            }
        }

        private void RenameSelected()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择要重命名的工序。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string oldName = Convert.ToString(_grid.SelectedRows[0].Cells["工序名称"].Value);
            string newName = (_txtProcess.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(oldName))
            {
                MessageBox.Show("未识别到原工序名称。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(newName))
            {
                MessageBox.Show("请输入新的工序名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(
                "确认将工序“" + oldName + "”重命名为“" + newName + "”？\n已分配该工序的员工将自动同步更新。",
                "重命名确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SQLiteHelper.RenameProcess(oldName, newName);
                LoadProcesses();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "重命名失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void AddProcess()
        {
            try
            {
                SQLiteHelper.AddProcess(_txtProcess.Text);
                _txtProcess.Clear();
                LoadProcesses();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "新增失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DeleteSelected()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择要删除的工序。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string name = Convert.ToString(_grid.SelectedRows[0].Cells["工序名称"].Value);
            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show("未识别到工序名称。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var confirm = MessageBox.Show("确认删除工序：" + name + "？", "删除确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SQLiteHelper.DeleteProcess(name);
                LoadProcesses();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
