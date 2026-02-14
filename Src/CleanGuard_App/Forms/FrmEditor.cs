using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmEditor : Form
    {
        private readonly TextBox _txtEmpNo = new TextBox();
        private readonly TextBox _txtName = new TextBox();
        private readonly ComboBox _cmbProcess = new ComboBox();
        private readonly ComboBox _cmb1FClothes = new ComboBox();
        private readonly ComboBox _cmb1FShoe = new ComboBox();
        private readonly ComboBox _cmb2FClothes = new ComboBox();
        private readonly ComboBox _cmb2FShoe = new ComboBox();
        private readonly Button _btnSave = new Button();

        private readonly Dictionary<string, DataGridView> _itemGrids = new Dictionary<string, DataGridView>();
        private readonly string[] _categories = { "无尘服", "安全鞋", "帆布鞋", "洁净帽" };
        private readonly string _editingEmpNo;

        public FrmEditor(string editingEmpNo = null)
        {
            _editingEmpNo = editingEmpNo;
            Text = string.IsNullOrWhiteSpace(_editingEmpNo) ? "员工录入" : "员工编辑";
            Width = 930;
            Height = 730;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadData();
        }

        private void InitializeLayout()
        {
            var grpBase = new GroupBox { Text = "基本信息", Left = 20, Top = 20, Width = 870, Height = 130 };

            var lblEmpNo = new Label { Text = "工号", Left = 20, Top = 35, Width = 80 };
            _txtEmpNo.SetBounds(100, 30, 180, 28);

            var lblName = new Label { Text = "姓名", Left = 320, Top = 35, Width = 80 };
            _txtName.SetBounds(400, 30, 180, 28);

            var lblProcess = new Label { Text = "工序", Left = 20, Top = 80, Width = 80 };
            _cmbProcess.SetBounds(100, 75, 180, 28);
            _cmbProcess.DropDownStyle = ComboBoxStyle.DropDown;
            _cmbProcess.Items.AddRange(SQLiteHelper.GetProcesses());

            grpBase.Controls.Add(lblEmpNo);
            grpBase.Controls.Add(_txtEmpNo);
            grpBase.Controls.Add(lblName);
            grpBase.Controls.Add(_txtName);
            grpBase.Controls.Add(lblProcess);
            grpBase.Controls.Add(_cmbProcess);
            Controls.Add(grpBase);

            var grp1F = new GroupBox { Text = "一楼柜位配置 (1F)", Left = 20, Top = 170, Width = 870, Height = 100 };
            grp1F.Controls.Add(new Label { Text = "衣柜", Left = 20, Top = 42, Width = 50 });
            _cmb1FClothes.SetBounds(70, 38, 180, 28);
            grp1F.Controls.Add(_cmb1FClothes);

            grp1F.Controls.Add(new Label { Text = "鞋柜", Left = 290, Top = 42, Width = 50 });
            _cmb1FShoe.SetBounds(340, 38, 180, 28);
            grp1F.Controls.Add(_cmb1FShoe);
            Controls.Add(grp1F);

            var grp2F = new GroupBox { Text = "二楼柜位配置 (2F)", Left = 20, Top = 290, Width = 870, Height = 100 };
            grp2F.Controls.Add(new Label { Text = "衣柜", Left = 20, Top = 42, Width = 50 });
            _cmb2FClothes.SetBounds(70, 38, 180, 28);
            grp2F.Controls.Add(_cmb2FClothes);

            grp2F.Controls.Add(new Label { Text = "鞋柜", Left = 290, Top = 42, Width = 50 });
            _cmb2FShoe.SetBounds(340, 38, 180, 28);
            grp2F.Controls.Add(_cmb2FShoe);
            Controls.Add(grp2F);

            var tabItems = new TabControl { Left = 20, Top = 410, Width = 870, Height = 230 };
            foreach (string category in _categories)
            {
                var tab = new TabPage(category);
                BuildDynamicItemPanel(tab, category);
                tabItems.TabPages.Add(tab);
            }
            Controls.Add(tabItems);

            _btnSave.Text = "保存";
            _btnSave.SetBounds(790, 650, 100, 30);
            _btnSave.Click += (s, e) => SaveEmployee();
            Controls.Add(_btnSave);
        }

        private void BuildDynamicItemPanel(TabPage tab, string category)
        {
            var btnAdd = new Button { Text = "新增一行", Left = 20, Top = 12, Width = 90, Height = 26 };
            var btnRemove = new Button { Text = "删除选中", Left = 120, Top = 12, Width = 90, Height = 26 };

            var grid = new DataGridView
            {
                Left = 20,
                Top = 45,
                Width = 810,
                Height = 145,
                AllowUserToAddRows = false,
                ReadOnly = false,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "序号", HeaderText = "序号", ReadOnly = true, FillWeight = 15 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "尺码", HeaderText = "尺码", FillWeight = 25 });
            if (NeedCodeColumn(category))
            {
                grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "编码", HeaderText = "编码", FillWeight = 30 });
            }
            if (NeedConditionColumn(category))
            {
                var conditionColumn = new DataGridViewComboBoxColumn
                {
                    Name = "新旧",
                    HeaderText = "新旧",
                    FillWeight = 20,
                    FlatStyle = FlatStyle.Flat
                };
                conditionColumn.Items.AddRange("新", "旧");
                grid.Columns.Add(conditionColumn);
            }
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "领用日期", HeaderText = "领用日期(yyyy-MM-dd)", FillWeight = 30 });

            btnAdd.Click += (s, e) => AddItemRow(grid, category, string.Empty, string.Empty, string.Empty, DateTime.Now.ToString("yyyy-MM-dd"));
            btnRemove.Click += (s, e) => RemoveSelectedRow(grid);

            tab.Controls.Add(btnAdd);
            tab.Controls.Add(btnRemove);
            tab.Controls.Add(grid);
            _itemGrids[category] = grid;
        }

        private static void AddItemRow(DataGridView grid, string category, string size, string itemCode, string itemCondition, string issueDate)
        {
            int index = grid.Rows.Count + 1;
            var values = new List<object> { index, size ?? string.Empty };
            if (NeedCodeColumn(category))
            {
                values.Add(itemCode ?? string.Empty);
            }
            if (NeedConditionColumn(category))
            {
                values.Add(string.IsNullOrWhiteSpace(itemCondition) ? "新" : itemCondition);
            }
            values.Add(issueDate ?? string.Empty);
            grid.Rows.Add(values.ToArray());
        }

        private static void RemoveSelectedRow(DataGridView grid)
        {
            if (grid.SelectedRows.Count == 0)
            {
                return;
            }

            grid.Rows.RemoveAt(grid.SelectedRows[0].Index);
            RebuildIndexes(grid);
        }

        private static void RebuildIndexes(DataGridView grid)
        {
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                grid.Rows[i].Cells["序号"].Value = i + 1;
            }
        }

        private void LoadData()
        {
            if (string.IsNullOrWhiteSpace(_editingEmpNo))
            {
                LoadLockerOptions(null);
                return;
            }

            var model = SQLiteHelper.GetEmployeeEditModel(_editingEmpNo);
            if (model == null)
            {
                MessageBox.Show("未找到员工信息。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }

            _txtEmpNo.Text = model.EmpNo;
            _txtEmpNo.ReadOnly = true;
            _txtName.Text = model.Name;
            _cmbProcess.Text = model.Process;

            LoadLockerOptions(model);
            SetComboValue(_cmb1FClothes, model.Locker1FClothes);
            SetComboValue(_cmb1FShoe, model.Locker1FShoe);
            SetComboValue(_cmb2FClothes, model.Locker2FClothes);
            SetComboValue(_cmb2FShoe, model.Locker2FShoe);

            LoadItemData(model.EmpNo);
        }

        private void LoadItemData(string empNo)
        {
            var items = SQLiteHelper.GetEmployeeItems(empNo);
            foreach (var category in _categories)
            {
                DataGridView grid = _itemGrids[category];
                grid.Rows.Clear();

                int slot = 1;
                foreach (var item in items)
                {
                    if (!string.Equals(item.Category, category, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    AddItemRow(grid, category, item.Size, item.ItemCode, item.ItemCondition, item.IssueDate);
                    grid.Rows[grid.Rows.Count - 1].Cells["序号"].Value = slot;
                    slot++;
                }
            }
        }

        private void LoadLockerOptions(EmployeeEditModel model)
        {
            BindLockerCombo(_cmb1FClothes, "1F", "衣柜", model != null ? model.Locker1FClothes : null);
            BindLockerCombo(_cmb1FShoe, "1F", "鞋柜", model != null ? model.Locker1FShoe : null);
            BindLockerCombo(_cmb2FClothes, "2F", "衣柜", model != null ? model.Locker2FClothes : null);
            BindLockerCombo(_cmb2FShoe, "2F", "鞋柜", model != null ? model.Locker2FShoe : null);
        }

        private void SaveEmployee()
        {
            try
            {
                string empNo = _txtEmpNo.Text;
                if (string.IsNullOrWhiteSpace(_editingEmpNo))
                {
                    SQLiteHelper.AddEmployee(
                        empNo,
                        _txtName.Text,
                        _cmbProcess.Text,
                        _cmb1FClothes.Text,
                        _cmb1FShoe.Text,
                        _cmb2FClothes.Text,
                        _cmb2FShoe.Text);
                }
                else
                {
                    empNo = _editingEmpNo;
                    SQLiteHelper.UpdateEmployee(
                        empNo,
                        _txtName.Text,
                        _cmbProcess.Text,
                        _cmb1FClothes.Text,
                        _cmb1FShoe.Text,
                        _cmb2FClothes.Text,
                        _cmb2FShoe.Text);
                }

                SQLiteHelper.ReplaceEmployeeItems(empNo, BuildItemInputs());

                MessageBox.Show("保存成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private List<EmployeeItemInput> BuildItemInputs()
        {
            var list = new List<EmployeeItemInput>();
            foreach (var category in _categories)
            {
                DataGridView grid = _itemGrids[category];
                int slot = 1;
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    string size = Convert.ToString(row.Cells["尺码"].Value);
                    string itemCode = NeedCodeColumn(category) ? Convert.ToString(row.Cells["编码"].Value) : string.Empty;
                    string itemCondition = NeedConditionColumn(category) ? Convert.ToString(row.Cells["新旧"].Value) : string.Empty;
                    string issueDate = Convert.ToString(row.Cells["领用日期"].Value);
                    if (!string.IsNullOrWhiteSpace(issueDate))
                    {
                        DateTime date;
                        if (!DateTime.TryParse(issueDate, out date))
                        {
                            throw new InvalidOperationException(string.Format("{0} 第 {1} 行领用日期格式错误，请使用 yyyy-MM-dd。", category, slot));
                        }

                        issueDate = date.ToString("yyyy-MM-dd");
                    }

                    if (string.IsNullOrWhiteSpace(size) && string.IsNullOrWhiteSpace(itemCode) && string.IsNullOrWhiteSpace(itemCondition) && string.IsNullOrWhiteSpace(issueDate))
                    {
                        continue;
                    }

                    list.Add(new EmployeeItemInput
                    {
                        Category = category,
                        SlotIndex = slot,
                        Size = string.IsNullOrWhiteSpace(size) ? null : size.Trim(),
                        ItemCode = string.IsNullOrWhiteSpace(itemCode) ? null : itemCode.Trim(),
                        ItemCondition = string.IsNullOrWhiteSpace(itemCondition) ? null : itemCondition.Trim(),
                        IssueDate = string.IsNullOrWhiteSpace(issueDate) ? null : issueDate
                    });
                    slot++;
                }
            }

            return list;
        }

        private static bool NeedCodeColumn(string category)
        {
            return string.Equals(category, "无尘服", StringComparison.Ordinal) ||
                   string.Equals(category, "洁净帽", StringComparison.Ordinal);
        }

        private static bool NeedConditionColumn(string category)
        {
            return string.Equals(category, "安全鞋", StringComparison.Ordinal) ||
                   string.Equals(category, "帆布鞋", StringComparison.Ordinal);
        }

        private static void BindLockerCombo(ComboBox comboBox, string location, string type, string selectedLocker)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add(string.Empty);
            comboBox.Items.AddRange(SQLiteHelper.GetAvailableLockersIncluding(location, type, selectedLocker));
            comboBox.SelectedIndex = 0;
        }

        private static void SetComboValue(ComboBox comboBox, string value)
        {
            comboBox.Text = string.IsNullOrWhiteSpace(value) ? string.Empty : value;
        }
    }
}
