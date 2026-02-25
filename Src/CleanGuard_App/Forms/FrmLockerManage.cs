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
        private readonly Button _btnDownloadTemplate = new Button();
        private readonly Button _btnImportLockers = new Button();
        private readonly DataGridView _grid = new DataGridView();

        public FrmLockerManage()
        {
            Text = "柜位异常状态维护";
            Width = 980;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            UiTheme.ApplyFormStyle(this);

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
            UiTheme.StylePrimaryButton(_btnQuery);

            Controls.Add(new Label { Text = "异常备注", Left = 510, Top = 24, Width = 60 });
            _txtRemark.SetBounds(575, 20, 240, 28);
            Controls.Add(_txtRemark);

            _btnSaveRemark.Text = "保存备注";
            _btnSaveRemark.SetBounds(825, 20, 90, 28);
            _btnSaveRemark.Click += (s, e) => SaveRemark();
            Controls.Add(_btnSaveRemark);
            UiTheme.StylePrimaryButton(_btnSaveRemark);

            _btnDownloadTemplate.Text = "下载柜位模板";
            _btnDownloadTemplate.SetBounds(575, 560, 130, 30);
            _btnDownloadTemplate.Click += (s, e) => DownloadLockerTemplate();
            Controls.Add(_btnDownloadTemplate);
            UiTheme.StylePrimaryButton(_btnDownloadTemplate);

            _btnImportLockers.Text = "导入柜位数据";
            _btnImportLockers.SetBounds(715, 560, 130, 30);
            _btnImportLockers.Click += (s, e) => ImportLockers();
            Controls.Add(_btnImportLockers);
            UiTheme.StylePrimaryButton(_btnImportLockers);

            _grid.SetBounds(20, 65, 920, 485);
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.MultiSelect = false;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid.SelectionChanged += (s, e) => FillRemarkFromSelection();
            Controls.Add(_grid);
            UiTheme.StyleDataGrid(_grid);
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

        private void DownloadLockerTemplate()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel 文件|*.xlsx|CSV 文件|*.csv|所有文件|*.*";
                dialog.FileName = "CleanGuard_柜位导入模板.xlsx";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                string warning;
                string actualPath = ImportHelper.ExportLockerTemplateWithFallback(dialog.FileName, out warning);
                string message = string.IsNullOrWhiteSpace(warning)
                    ? "柜位模板下载成功。"
                    : warning + "\n文件路径：" + actualPath;
                MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ImportLockers()
        {
            var tip = "导入将按文件重建柜位列表。支持选择 xlsx/csv 模板；首行为：1F衣柜、1F鞋柜、2F衣柜、2F鞋柜。是否继续？";
            if (MessageBox.Show(tip, "确认导入", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel 文件|*.xlsx|CSV 文件|*.csv|所有文件|*.*";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                var result = ImportHelper.ImportLockersFromFile(dialog.FileName);
                LoadLockers();
                var msg = string.Format("柜位导入完成：成功 {0}，失败 {1}", result.SuccessCount, result.FailedCount);
                if (result.Errors.Count > 0)
                {
                    msg += Environment.NewLine + string.Join(Environment.NewLine, result.Errors.ToArray());
                }

                MessageBox.Show(msg, "导入结果", MessageBoxButtons.OK,
                    result.FailedCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
        }
    }
}
