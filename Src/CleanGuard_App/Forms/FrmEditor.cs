using System;
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

        public FrmEditor()
        {
            Text = "员工录入与编辑";
            Width = 900;
            Height = 650;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadLockerOptions();
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

            var tabItems = new TabControl { Left = 20, Top = 410, Width = 840, Height = 150 };
            tabItems.TabPages.Add("无尘服");
            tabItems.TabPages.Add("鞋类");
            tabItems.TabPages.Add("帽子");
            Controls.Add(tabItems);

            _btnSave.Text = "保存";
            _btnSave.SetBounds(760, 570, 100, 30);
            _btnSave.Click += (s, e) => SaveEmployee();
            Controls.Add(_btnSave);
        }

        private void LoadLockerOptions()
        {
            BindLockerCombo(_cmb1FClothes, "1F", "衣柜");
            BindLockerCombo(_cmb1FShoe, "1F", "鞋柜");
            BindLockerCombo(_cmb2FClothes, "2F", "衣柜");
            BindLockerCombo(_cmb2FShoe, "2F", "鞋柜");
        }

        private void SaveEmployee()
        {
            try
            {
                SQLiteHelper.AddEmployee(
                    _txtEmpNo.Text,
                    _txtName.Text,
                    _cmbProcess.Text,
                    _cmb1FClothes.Text,
                    _cmb1FShoe.Text,
                    _cmb2FClothes.Text,
                    _cmb2FShoe.Text);

                MessageBox.Show("保存成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void BindLockerCombo(ComboBox comboBox, string location, string type)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add(string.Empty);
            comboBox.Items.AddRange(SQLiteHelper.GetAvailableLockers(location, type));
            comboBox.SelectedIndex = 0;
        }
    }
}
