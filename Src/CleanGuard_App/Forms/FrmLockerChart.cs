using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmLockerChart : Form
    {
        private readonly Chart _chart = new Chart();
        private readonly Label _lblSummary = new Label();
        private readonly ComboBox _cmbFloor = new ComboBox();
        private readonly Button _btnRefresh = new Button();
        private readonly Button _btnExportImage = new Button();
        private readonly Button _btnTrend = new Button();

        public FrmLockerChart()
        {
            Text = "柜位分布图";
            Width = 900;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadData();
        }

        private void InitializeLayout()
        {
            _cmbFloor.SetBounds(20, 15, 100, 28);
            _cmbFloor.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbFloor.Items.AddRange(new object[] { "全部", "1F", "2F" });
            _cmbFloor.SelectedIndex = 0;
            Controls.Add(_cmbFloor);

            _btnRefresh.Text = "刷新";
            _btnRefresh.SetBounds(130, 15, 80, 28);
            _btnRefresh.Click += (s, e) => LoadData();
            Controls.Add(_btnRefresh);

            _btnExportImage.Text = "导出图片";
            _btnExportImage.SetBounds(220, 15, 100, 28);
            _btnExportImage.Click += (s, e) => ExportChartImage();
            Controls.Add(_btnExportImage);

            _btnTrend.Text = "占用趋势";
            _btnTrend.SetBounds(330, 15, 100, 28);
            _btnTrend.Click += (s, e) => OpenTrend();
            Controls.Add(_btnTrend);

            _chart.SetBounds(20, 55, 840, 420);
            _chart.ChartAreas.Add(new ChartArea("MainArea"));
            _chart.Legends.Add(new Legend("Legend"));
            Controls.Add(_chart);

            _lblSummary.SetBounds(20, 485, 840, 90);
            _lblSummary.Font = new Font("Microsoft YaHei", 10, FontStyle.Regular);
            _lblSummary.BorderStyle = BorderStyle.FixedSingle;
            Controls.Add(_lblSummary);
        }

        private void LoadData()
        {
            var summary = SQLiteHelper.GetLockerSummary();
            string floor = _cmbFloor.SelectedItem != null ? _cmbFloor.SelectedItem.ToString() : "全部";

            _chart.Series.Clear();

            if (floor == "全部" || floor == "1F")
            {
                AddPieSeries("1F衣柜", summary.OneFClothesOccupied, summary.OneFClothesTotal);
                AddPieSeries("1F鞋柜", summary.OneFShoeOccupied, summary.OneFShoeTotal);
            }

            if (floor == "全部" || floor == "2F")
            {
                AddPieSeries("2F衣柜", summary.TwoFClothesOccupied, summary.TwoFClothesTotal);
                AddPieSeries("2F鞋柜", summary.TwoFShoeOccupied, summary.TwoFShoeTotal);
            }

            _lblSummary.Text = BuildSummaryText(summary, floor);
        }

        private void AddPieSeries(string title, int occupied, int total)
        {
            int free = Math.Max(0, total - occupied);
            var series = new Series(title)
            {
                ChartType = SeriesChartType.Pie,
                IsValueShownAsLabel = true,
                ChartArea = "MainArea",
                Legend = "Legend"
            };

            series.Points.AddXY("占用", occupied);
            series.Points.AddXY("空闲", free);
            series.Points[0].Color = Color.IndianRed;
            series.Points[1].Color = Color.SeaGreen;
            series["PieLabelStyle"] = "Outside";
            _chart.Series.Add(series);
        }

        private static string BuildSummaryText(LockerSummary summary, string floor)
        {
            if (floor == "1F")
            {
                return "占用概览（1F）\n" +
                       BuildLine("1F衣柜", summary.OneFClothesOccupied, summary.OneFClothesTotal) +
                       BuildLine("1F鞋柜", summary.OneFShoeOccupied, summary.OneFShoeTotal);
            }

            if (floor == "2F")
            {
                return "占用概览（2F）\n" +
                       BuildLine("2F衣柜", summary.TwoFClothesOccupied, summary.TwoFClothesTotal) +
                       BuildLine("2F鞋柜", summary.TwoFShoeOccupied, summary.TwoFShoeTotal);
            }

            return "占用概览（全部）\n" +
                   BuildLine("1F衣柜", summary.OneFClothesOccupied, summary.OneFClothesTotal) +
                   BuildLine("1F鞋柜", summary.OneFShoeOccupied, summary.OneFShoeTotal) +
                   BuildLine("2F衣柜", summary.TwoFClothesOccupied, summary.TwoFClothesTotal) +
                   BuildLine("2F鞋柜", summary.TwoFShoeOccupied, summary.TwoFShoeTotal);
        }

        private static string BuildLine(string name, int occupied, int total)
        {
            decimal rate = total == 0 ? 0 : Math.Round(occupied * 100m / total, 1);
            return string.Format("{0}: {1}/{2}（占用率 {3}%）\n", name, occupied, total, rate);
        }

        private void OpenTrend()
        {
            using (var form = new FrmLockerTrend())
            {
                form.ShowDialog(this);
            }
        }

        private void ExportChartImage()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "PNG 图片|*.png";
                dialog.FileName = "LockerChart.png";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _chart.SaveImage(dialog.FileName, ChartImageFormat.Png);
                MessageBox.Show("图表已导出：" + Path.GetFileName(dialog.FileName), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
