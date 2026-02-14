using System;
using System.Data;
using System.Windows.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmLockerManage : Form
    {
        private readonly ComboBox _cmbLocation = new ComboBox();
        private readonly ComboBox _cmbType = new ComboBox();
        private readonly CheckBox _chkAbnormal = new CheckBox();
        private readonly TextBox _txtRemark = new TextBox();
        private readonly Button _btnQuery = new Button();
        private readonly Button _btnSaveRemark = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmLockerManage()
        {
            Text = "柜位异常状态维护";
            Width = 980;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadLockers();
        }

        private void InitializeLayout()
        {
            Controls.Add(new Label { Text = "楼层", Left = 20, Top = 24, Width = 40 });
            _cmbLocation.SetBounds(60, 20, 90, 28);
            _cmbLocation.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbLocation.Items.AddRange(new object[] { "全部", "1F", "2F" });
            _cmbLocation.SelectedIndex = 0;
            Controls.Add(_cmbLocation);

            Controls.Add(new Label { Text = "类型", Left = 165, Top = 24, Width = 40 });
            _cmbType.SetBounds(205, 20, 100, 28);
            _cmbType.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbType.Items.AddRange(new object[] { "全部", "衣柜", "鞋柜" });
            _cmbType.SelectedIndex = 0;
            Controls.Add(_cmbType);

            _chkAbnormal.Text = "仅看异常";
            _chkAbnormal.SetBounds(320, 22, 90, 24);
            Controls.Add(_chkAbnormal);

            _btnQuery.Text = "查询";
            _btnQuery.SetBounds(420, 20, 70, 28);
            _btnQuery.Click += (s, e) => LoadLockers();
            Controls.Add(_btnQuery);

            Controls.Add(new Label { Text = "异常备注", Left = 510, Top = 24, Width = 60 });
            _txtRemark.SetBounds(575, 20, 240, 28);
            Controls.Add(_txtRemark);

            _btnSaveRemark.Text = "保存备注";
            _btnSaveRemark.SetBounds(825, 20, 90, 28);
            _btnSaveRemark.Click += (s, e) => SaveRemark();
            Controls.Add(_btnSaveRemark);

            _grid.SetBounds(20, 65, 920, 500);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.SelectionChanged += (s, e) => FillRemarkFromSelection();
            Controls.Add(_grid);
        }

        private void LoadLockers()
        {
            string location = Convert.ToString(_cmbLocation.SelectedItem);
            if (location == "全部") location = string.Empty;
            string type = Convert.ToString(_cmbType.SelectedItem);
            if (type == "全部") type = string.Empty;

            DataTable table = SQLiteHelper.QueryLockers(location, type, _chkAbnormal.Checked);
            _grid.DataSource = table;
            FillRemarkFromSelection();
        }

        private void FillRemarkFromSelection()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                _txtRemark.Text = string.Empty;
                return;
            }

            _txtRemark.Text = Convert.ToString(_grid.SelectedRows[0].Cells["异常备注"].Value);
        }

        private void SaveRemark()
        {
            if (_grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选择一个柜位。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string lockerId = Convert.ToString(_grid.SelectedRows[0].Cells["柜号"].Value);
            try
            {
                SQLiteHelper.UpdateLockerRemark(lockerId, _txtRemark.Text);
                LoadLockers();
                MessageBox.Show("柜位备注已保存。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
