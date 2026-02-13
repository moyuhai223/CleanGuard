using System;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmEditor : Form
    {
        private readonly ComboBox _cmbProcess = new ComboBox();
        private readonly ComboBox _cmb1FClothes = new ComboBox();
        private readonly ComboBox _cmb1FShoe = new ComboBox();
        private readonly ComboBox _cmb2FClothes = new ComboBox();
        private readonly ComboBox _cmb2FShoe = new ComboBox();

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
            var grpBase = new GroupBox { Text = "基本信息", Left = 20, Top = 20, Width = 840, Height = 100 };
            _cmbProcess.Left = 20;
            _cmbProcess.Top = 40;
            _cmbProcess.Width = 200;
            _cmbProcess.DropDownStyle = ComboBoxStyle.DropDown;
            _cmbProcess.Items.AddRange(SQLiteHelper.GetProcesses());
            grpBase.Controls.Add(_cmbProcess);
            Controls.Add(grpBase);

            var grp1F = new GroupBox { Text = "一楼柜位配置 (1F)", Left = 20, Top = 140, Width = 840, Height = 100 };
            _cmb1FClothes.SetBounds(20, 40, 180, 28);
            _cmb1FShoe.SetBounds(230, 40, 180, 28);
            grp1F.Controls.Add(_cmb1FClothes);
            grp1F.Controls.Add(_cmb1FShoe);
            Controls.Add(grp1F);

            var grp2F = new GroupBox { Text = "二楼柜位配置 (2F)", Left = 20, Top = 260, Width = 840, Height = 100 };
            _cmb2FClothes.SetBounds(20, 40, 180, 28);
            _cmb2FShoe.SetBounds(230, 40, 180, 28);
            grp2F.Controls.Add(_cmb2FClothes);
            grp2F.Controls.Add(_cmb2FShoe);
            Controls.Add(grp2F);

            var tabItems = new TabControl { Left = 20, Top = 380, Width = 840, Height = 200 };
            tabItems.TabPages.Add("无尘服");
            tabItems.TabPages.Add("鞋类");
            tabItems.TabPages.Add("帽子");
            Controls.Add(tabItems);
        }

        private void LoadLockerOptions()
        {
            BindLockerCombo(_cmb1FClothes, "1F", "衣柜");
            BindLockerCombo(_cmb1FShoe, "1F", "鞋柜");
            BindLockerCombo(_cmb2FClothes, "2F", "衣柜");
            BindLockerCombo(_cmb2FShoe, "2F", "鞋柜");
        }

        private static void BindLockerCombo(ComboBox comboBox, string location, string type)
        {
            comboBox.Items.Clear();
            comboBox.Items.Add(string.Empty);
            comboBox.Items.AddRange(SQLiteHelper.GetAvailableLockers(location, type));
        }
    }
}
