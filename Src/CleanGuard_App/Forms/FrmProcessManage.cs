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
        private readonly Button _btnRefresh = new Button();

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

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(490, 20, 70, 28);
            _btnRefresh.Click += (s, e) => LoadProcesses();
            Controls.Add(_btnRefresh);

            _grid.SetBounds(20, 65, 620, 390);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            Controls.Add(_grid);
        }

        private void LoadProcesses()
        {
            DataTable table = SQLiteHelper.QueryProcessesTable();
            _grid.DataSource = table;
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
