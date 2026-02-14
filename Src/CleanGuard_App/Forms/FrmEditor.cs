using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
        private readonly Button _btnLimitSettings = new Button();

        private readonly Dictionary<string, DataGridView> _itemGrids = new Dictionary<string, DataGridView>();
        private readonly Dictionary<string, int> _categoryLimits = new Dictionary<string, int>
        {};
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
            LoadCategoryLimits();
            LoadData();
        }

        private void LoadCategoryLimits()
        {
            _categoryLimits.Clear();
            foreach (var pair in SQLiteHelper.GetItemCategoryLimits())
            {
                _categoryLimits[pair.Key] = pair.Value;
            }
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

            _btnLimitSettings.Text = "用品上限设置";
            _btnLimitSettings.SetBounds(670, 650, 110, 30);
            _btnLimitSettings.Click += (s, e) => OpenLimitSettingsDialog();
            Controls.Add(_btnLimitSettings);
        }

        private void BuildDynamicItemPanel(TabPage tab, string category)
        {
            var btnAdd = new Button { Text = "新增一行", Left = 20, Top = 12, Width = 90, Height = 26 };
            var btnRemove = new Button { Text = "删除选中", Left = 120, Top = 12, Width = 90, Height = 26 };
            var btnPaste = new Button { Text = "批量粘贴", Left = 220, Top = 12, Width = 90, Height = 26 };
            var btnTemplate = new Button { Text = "下载模板", Left = 320, Top = 12, Width = 90, Height = 26 };
            var btnImport = new Button { Text = "导入模板", Left = 420, Top = 12, Width = 90, Height = 26 };

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

            btnAdd.Click += (s, e) =>
            {
                try
                {
                    AddItemRowWithLimit(grid, category, string.Empty, string.Empty, string.Empty, DateTime.Now.ToString("yyyy-MM-dd"));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            btnRemove.Click += (s, e) => RemoveSelectedRow(grid);
            btnPaste.Click += (s, e) => PasteItemRows(grid, category);
            btnTemplate.Click += (s, e) => ExportItemTemplate(category);
            btnImport.Click += (s, e) => ImportItemTemplate(grid, category);

            tab.Controls.Add(btnAdd);
            tab.Controls.Add(btnRemove);
            tab.Controls.Add(btnPaste);
            tab.Controls.Add(btnTemplate);
            tab.Controls.Add(btnImport);
            tab.Controls.Add(grid);
            _itemGrids[category] = grid;
        }

        private void AddItemRowWithLimit(DataGridView grid, string category, string size, string itemCode, string itemCondition, string issueDate)
        {
            int max = GetCategoryLimit(category);
            if (grid.Rows.Count >= max)
            {
                throw new InvalidOperationException(string.Format("{0} 最多允许 {1} 行。", category, max));
            }

            AddItemRow(grid, category, size, itemCode, itemCondition, issueDate);
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
                int max = GetCategoryLimit(category);
                if (grid.Rows.Count > max)
                {
                    throw new InvalidOperationException(string.Format("{0} 超出最大行数 {1}，请删除后重试。", category, max));
                }

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

        private int GetCategoryLimit(string category)
        {
            int value;
            return _categoryLimits.TryGetValue(category, out value) ? value : 10;
        }

        private void PasteItemRows(DataGridView grid, string category)
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("剪贴板为空。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var preview = BuildImportPreview(category, text, false);
                if (!ShowImportPreviewDialog(preview, category))
                {
                    return;
                }

                int added = AppendValidRows(grid, category, preview);
                MessageBox.Show(string.Format("已批量粘贴 {0} 行。", added), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "批量粘贴失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ExportItemTemplate(string category)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "CSV 文件|*.csv";
                dialog.FileName = "用品模板_" + category + ".csv";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                string header = BuildItemTemplateHeader(category);
                string sample = BuildItemTemplateSample(category);
                File.WriteAllText(dialog.FileName, header + Environment.NewLine + sample, Encoding.UTF8);
                MessageBox.Show("模板下载成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportItemTemplate(DataGridView grid, string category)
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
                    string text = File.ReadAllText(dialog.FileName, Encoding.UTF8);
                    var preview = BuildImportPreview(category, text, true);
                    if (!ShowImportPreviewDialog(preview, category))
                    {
                        return;
                    }

                    int added = AppendValidRows(grid, category, preview);
                    MessageBox.Show(string.Format("导入成功，共新增 {0} 行。", added), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private List<ItemImportPreviewRow> BuildImportPreview(string category, string text, bool skipHeader)
        {
            string[] lines = text.Replace("\r", string.Empty).Split('\n');
            int start = skipHeader && lines.Length > 0 ? 1 : 0;
            var list = new List<ItemImportPreviewRow>();
            int slot = _itemGrids[category].Rows.Count + 1;
            int max = GetCategoryLimit(category);

            for (int i = start; i < lines.Length; i++)
            {
                string line = (lines[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var row = ParseImportLine(category, line, i + 1);
                if (string.IsNullOrWhiteSpace(row.Error) && slot > max)
                {
                    row.Error = string.Format("超过类别上限 {0}", max);
                }

                if (string.IsNullOrWhiteSpace(row.Error))
                {
                    slot++;
                }

                list.Add(row);
            }

            return list;
        }

        private ItemImportPreviewRow ParseImportLine(string category, string line, int sourceLine)
        {
            string[] cells = line.Split('\t');
            if (cells.Length == 1 && line.Contains(","))
            {
                cells = line.Split(',');
            }

            string size = GetCell(cells, 0);
            string itemCode = string.Empty;
            string itemCondition = string.Empty;
            string issueDate = string.Empty;

            if (NeedCodeColumn(category))
            {
                itemCode = GetCell(cells, 1);
                issueDate = GetCell(cells, 2);
            }
            else
            {
                itemCondition = GetCell(cells, 1);
                issueDate = GetCell(cells, 2);
            }

            string error = string.Empty;
            if (!string.IsNullOrWhiteSpace(issueDate))
            {
                DateTime date;
                if (!DateTime.TryParse(issueDate, out date))
                {
                    error = "领用日期格式错误(yyyy-MM-dd)";
                }
                else
                {
                    issueDate = date.ToString("yyyy-MM-dd");
                }
            }

            if (NeedConditionColumn(category) && !string.IsNullOrWhiteSpace(itemCondition) && itemCondition != "新" && itemCondition != "旧")
            {
                error = "新旧仅支持：新/旧";
            }

            return new ItemImportPreviewRow
            {
                SourceLine = sourceLine,
                Size = size,
                ItemCode = itemCode,
                ItemCondition = itemCondition,
                IssueDate = issueDate,
                Error = error
            };
        }

        private bool ShowImportPreviewDialog(List<ItemImportPreviewRow> preview, string category)
        {
            using (var dialog = new Form())
            {
                dialog.Text = "用品导入预检 - " + category;
                dialog.Width = 860;
                dialog.Height = 500;
                dialog.StartPosition = FormStartPosition.CenterParent;

                var grid = new DataGridView
                {
                    Left = 10,
                    Top = 10,
                    Width = 820,
                    Height = 390,
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                };

                grid.Columns.Add("line", "源行号");
                grid.Columns.Add("size", "尺码");
                grid.Columns.Add("code", "编码");
                grid.Columns.Add("condition", "新旧");
                grid.Columns.Add("date", "领用日期");
                grid.Columns.Add("error", "预检结果");

                int errorCount = 0;
                var errorLines = new StringBuilder();
                foreach (var row in preview)
                {
                    string status = string.IsNullOrWhiteSpace(row.Error) ? "通过" : row.Error;
                    grid.Rows.Add(row.SourceLine, row.Size, row.ItemCode, row.ItemCondition, row.IssueDate, status);
                    if (!string.IsNullOrWhiteSpace(row.Error))
                    {
                        errorCount++;
                        errorLines.AppendLine(string.Format("第 {0} 行：{1}", row.SourceLine, row.Error));
                    }
                }

                var lbl = new Label
                {
                    Left = 10,
                    Top = 408,
                    Width = 500,
                    Height = 24,
                    Text = string.Format("预检完成：共 {0} 行，错误 {1} 行。", preview.Count, errorCount)
                };

                var btnCopy = new Button { Left = 520, Top = 404, Width = 100, Height = 28, Text = "复制错误" };
                btnCopy.Click += (s, e) =>
                {
                    Clipboard.SetText(errorLines.ToString());
                    MessageBox.Show("错误信息已复制。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                var btnContinue = new Button { Left = 630, Top = 404, Width = 90, Height = 28, Text = "继续导入" };
                btnContinue.Click += (s, e) => dialog.DialogResult = DialogResult.OK;

                var btnCancel = new Button { Left = 730, Top = 404, Width = 90, Height = 28, Text = "取消" };
                btnCancel.Click += (s, e) => dialog.DialogResult = DialogResult.Cancel;

                dialog.Controls.Add(grid);
                dialog.Controls.Add(lbl);
                dialog.Controls.Add(btnCopy);
                dialog.Controls.Add(btnContinue);
                dialog.Controls.Add(btnCancel);

                if (preview.Count == 0)
                {
                    MessageBox.Show("未检测到可导入数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                return dialog.ShowDialog(this) == DialogResult.OK;
            }
        }

        private int AppendValidRows(DataGridView grid, string category, List<ItemImportPreviewRow> preview)
        {
            int added = 0;
            foreach (var row in preview)
            {
                if (!string.IsNullOrWhiteSpace(row.Error))
                {
                    continue;
                }

                AddItemRowWithLimit(grid, category, row.Size, row.ItemCode, row.ItemCondition, row.IssueDate);
                added++;
            }

            return added;
        }

        private static string GetCell(string[] cells, int index)
        {
            if (cells == null || index < 0 || index >= cells.Length)
            {
                return string.Empty;
            }

            return (cells[index] ?? string.Empty).Trim();
        }

        private static string BuildItemTemplateHeader(string category)
        {
            if (NeedCodeColumn(category))
            {
                return "尺码,编码,领用日期";
            }

            return "尺码,新旧,领用日期";
        }

        private static string BuildItemTemplateSample(string category)
        {
            if (NeedCodeColumn(category))
            {
                return "L,ABC001," + DateTime.Now.ToString("yyyy-MM-dd");
            }

            return "42,新," + DateTime.Now.ToString("yyyy-MM-dd");
        }

        private void OpenLimitSettingsDialog()
        {
            using (var dialog = new Form())
            {
                dialog.Text = "用品上限设置";
                dialog.Width = 360;
                dialog.Height = 290;
                dialog.StartPosition = FormStartPosition.CenterParent;

                var labels = new[] { "无尘服", "安全鞋", "帆布鞋", "洁净帽" };
                var inputs = new Dictionary<string, NumericUpDown>();
                int top = 20;
                foreach (var name in labels)
                {
                    var lbl = new Label { Left = 20, Top = top + 5, Width = 90, Text = name };
                    var num = new NumericUpDown
                    {
                        Left = 120,
                        Top = top,
                        Width = 120,
                        Minimum = 1,
                        Maximum = 999,
                        Value = GetCategoryLimit(name)
                    };

                    dialog.Controls.Add(lbl);
                    dialog.Controls.Add(num);
                    inputs[name] = num;
                    top += 40;
                }

                var btnSave = new Button { Left = 150, Top = top + 10, Width = 90, Height = 30, Text = "保存" };
                var btnCancel = new Button { Left = 250, Top = top + 10, Width = 70, Height = 30, Text = "取消" };
                btnCancel.Click += (s, e) => dialog.DialogResult = DialogResult.Cancel;
                btnSave.Click += (s, e) =>
                {
                    var limits = new Dictionary<string, int>(StringComparer.Ordinal)
                    {
                        { "无尘服", Convert.ToInt32(inputs["无尘服"].Value) },
                        { "安全鞋", Convert.ToInt32(inputs["安全鞋"].Value) },
                        { "帆布鞋", Convert.ToInt32(inputs["帆布鞋"].Value) },
                        { "洁净帽", Convert.ToInt32(inputs["洁净帽"].Value) }
                    };

                    try
                    {
                        SQLiteHelper.SaveItemCategoryLimits(limits);
                        LoadCategoryLimits();
                        MessageBox.Show("用品上限已保存。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dialog.DialogResult = DialogResult.OK;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };

                dialog.Controls.Add(btnSave);
                dialog.Controls.Add(btnCancel);
                dialog.ShowDialog(this);
            }
        }

        private class ItemImportPreviewRow
        {
            public int SourceLine { get; set; }
            public string Size { get; set; }
            public string ItemCode { get; set; }
            public string ItemCondition { get; set; }
            public string IssueDate { get; set; }
            public string Error { get; set; }
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
