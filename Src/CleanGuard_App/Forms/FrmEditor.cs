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

        private readonly List<ItemSlotControl> _itemSlots = new List<ItemSlotControl>();
        private readonly string _editingEmpNo;

        public FrmEditor(string editingEmpNo = null)
        {
            _editingEmpNo = editingEmpNo;
            Text = string.IsNullOrWhiteSpace(_editingEmpNo) ? "员工录入" : "员工编辑";
            Width = 900;
            Height = 700;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadData();
        }

        private void InitializeLayout()
        {
            var grpBase = new GroupBox { Text = "基本信息", Left = 20, Top = 20, Width = 840, Height = 130 };

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

            var grp1F = new GroupBox { Text = "一楼柜位配置 (1F)", Left = 20, Top = 170, Width = 840, Height = 100 };
            grp1F.Controls.Add(new Label { Text = "衣柜", Left = 20, Top = 42, Width = 50 });
            _cmb1FClothes.SetBounds(70, 38, 180, 28);
            grp1F.Controls.Add(_cmb1FClothes);

            grp1F.Controls.Add(new Label { Text = "鞋柜", Left = 290, Top = 42, Width = 50 });
            _cmb1FShoe.SetBounds(340, 38, 180, 28);
            grp1F.Controls.Add(_cmb1FShoe);
            Controls.Add(grp1F);

            var grp2F = new GroupBox { Text = "二楼柜位配置 (2F)", Left = 20, Top = 290, Width = 840, Height = 100 };
            grp2F.Controls.Add(new Label { Text = "衣柜", Left = 20, Top = 42, Width = 50 });
            _cmb2FClothes.SetBounds(70, 38, 180, 28);
            grp2F.Controls.Add(_cmb2FClothes);

            grp2F.Controls.Add(new Label { Text = "鞋柜", Left = 290, Top = 42, Width = 50 });
            _cmb2FShoe.SetBounds(340, 38, 180, 28);
            grp2F.Controls.Add(_cmb2FShoe);
            Controls.Add(grp2F);

            var tabItems = new TabControl { Left = 20, Top = 410, Width = 840, Height = 200 };
            var tabDustSuit = new TabPage("无尘服");
            var tabShoes = new TabPage("鞋类");
            var tabHat = new TabPage("帽子");
            tabItems.TabPages.Add(tabDustSuit);
            tabItems.TabPages.Add(tabShoes);
            tabItems.TabPages.Add(tabHat);
            Controls.Add(tabItems);

            BuildItemSlots(tabDustSuit, "无尘服", 3);
            BuildItemSlots(tabShoes, "安全鞋", 2, 10);
            BuildItemSlots(tabShoes, "帆布鞋", 2, 100);
            BuildItemSlots(tabHat, "洁净帽", 3);

            _btnSave.Text = "保存";
            _btnSave.SetBounds(760, 620, 100, 30);
            _btnSave.Click += (s, e) => SaveEmployee();
            Controls.Add(_btnSave);
        }

        private void BuildItemSlots(TabPage tab, string category, int count, int topOffset = 0)
        {
            int startTop = 15 + topOffset;
            for (int i = 1; i <= count; i++)
            {
                var chk = new CheckBox { Text = string.Format("{0}{1}", category, i), Left = 20, Top = startTop + (i - 1) * 30, Width = 90 };
                var txtSize = new TextBox { Left = 130, Top = startTop + (i - 1) * 30 - 2, Width = 120, Enabled = false };
                var dtIssue = new DateTimePicker
                {
                    Left = 280,
                    Top = startTop + (i - 1) * 30 - 2,
                    Width = 140,
                    Format = DateTimePickerFormat.Custom,
                    CustomFormat = "yyyy-MM-dd",
                    Enabled = false
                };

                chk.CheckedChanged += (s, e) =>
                {
                    txtSize.Enabled = chk.Checked;
                    dtIssue.Enabled = chk.Checked;
                };

                tab.Controls.Add(chk);
                tab.Controls.Add(new Label { Text = "尺码", Left = 110, Top = startTop + (i - 1) * 30 + 3, Width = 30 });
                tab.Controls.Add(txtSize);
                tab.Controls.Add(new Label { Text = "领用日期", Left = 250, Top = startTop + (i - 1) * 30 + 3, Width = 50 });
                tab.Controls.Add(dtIssue);

                _itemSlots.Add(new ItemSlotControl
                {
                    Category = category,
                    SlotIndex = i,
                    Check = chk,
                    Size = txtSize,
                    IssueDate = dtIssue
                });
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
            foreach (var slot in _itemSlots)
            {
                var item = items.Find(x => x.Category == slot.Category && x.SlotIndex == slot.SlotIndex);
                if (item == null)
                {
                    continue;
                }

                slot.Check.Checked = true;
                slot.Size.Text = item.Size;

                DateTime parsed;
                if (DateTime.TryParse(item.IssueDate, out parsed))
                {
                    slot.IssueDate.Value = parsed;
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
            foreach (var slot in _itemSlots)
            {
                if (!slot.Check.Checked)
                {
                    continue;
                }

                list.Add(new EmployeeItemInput
                {
                    Category = slot.Category,
                    SlotIndex = slot.SlotIndex,
                    Size = string.IsNullOrWhiteSpace(slot.Size.Text) ? null : slot.Size.Text.Trim(),
                    IssueDate = slot.IssueDate.Value.ToString("yyyy-MM-dd")
                });
            }

            return list;
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

    public class ItemSlotControl
    {
        public string Category { get; set; }
        public int SlotIndex { get; set; }
        public CheckBox Check { get; set; }
        public TextBox Size { get; set; }
        public DateTimePicker IssueDate { get; set; }
    }
}
