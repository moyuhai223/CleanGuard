using System;
using System.Collections.Generic;
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
        private readonly FlowLayoutPanel _pnlHeatmap = new FlowLayoutPanel();
        private readonly Label _lblHeatLegend = new Label();

        public FrmLockerChart()
        {
            Text = "柜位分布图";
            Width = 900;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            UiTheme.ApplyFormStyle(this);

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
            UiTheme.StylePrimaryButton(_btnRefresh);

            _btnExportImage.Text = "导出图片";
            _btnExportImage.SetBounds(220, 15, 100, 28);
            _btnExportImage.Click += (s, e) => ExportChartImage();
            Controls.Add(_btnExportImage);
            UiTheme.StylePrimaryButton(_btnExportImage);

            _btnTrend.Text = "占用趋势";
            _btnTrend.SetBounds(330, 15, 100, 28);
            _btnTrend.Click += (s, e) => OpenTrend();
            Controls.Add(_btnTrend);
            UiTheme.StylePrimaryButton(_btnTrend);

            _chart.SetBounds(20, 55, 840, 420);
            _chart.ChartAreas.Add(new ChartArea("MainArea"));
            _chart.Legends.Add(new Legend("Legend"));
            Controls.Add(_chart);

            _lblHeatLegend.SetBounds(20, 485, 840, 20);
            _lblHeatLegend.Text = "区域热力块：绿色=低占用，橙色=中占用，红色=高占用；橙色边框表示该区域存在异常备注柜位";
            Controls.Add(_lblHeatLegend);

            _pnlHeatmap.SetBounds(20, 510, 840, 150);
            _pnlHeatmap.AutoScroll = true;
            _pnlHeatmap.WrapContents = true;
            _pnlHeatmap.BorderStyle = BorderStyle.FixedSingle;
            Controls.Add(_pnlHeatmap);

            _lblSummary.SetBounds(20, 665, 840, 90);
            _lblSummary.Font = new Font("Microsoft YaHei", 10, FontStyle.Regular);
            _lblSummary.BorderStyle = BorderStyle.FixedSingle;
            Controls.Add(_lblSummary);

            Height = 820;
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

            LoadHeatmapBlocks(floor);
            _lblSummary.Text = BuildSummaryText(summary, floor);
        }

        private void LoadHeatmapBlocks(string floor)
        {
            _pnlHeatmap.SuspendLayout();
            _pnlHeatmap.Controls.Clear();

            string filter = floor == "全部" ? string.Empty : floor;
            List<LockerHeatBlock> blocks = SQLiteHelper.GetLockerHeatBlocks(filter);
            if (blocks.Count == 0)
            {
                _pnlHeatmap.Controls.Add(new Label { Text = "暂无可展示的柜位区域数据。", AutoSize = true, Margin = new Padding(10) });
                _pnlHeatmap.ResumeLayout();
                return;
            }

            foreach (var block in blocks)
            {
                _pnlHeatmap.Controls.Add(CreateHeatBlockCard(block));
            }

            _pnlHeatmap.ResumeLayout();
        }

        private static Control CreateHeatBlockCard(LockerHeatBlock block)
        {
            var panel = new Panel
            {
                Width = 190,
                Height = 85,
                Margin = new Padding(6),
                BackColor = GetHeatColor(block.OccupancyRate),
                BorderStyle = block.AbnormalCount > 0 ? BorderStyle.Fixed3D : BorderStyle.FixedSingle
            };

            decimal rate = Math.Round(block.OccupancyRate * 100m, 1);
            var title = new Label
            {
                Left = 8,
                Top = 8,
                Width = 172,
                Height = 18,
                Text = string.Format("{0}-{1}-{2}", block.Location, block.Type, block.RegionName),
                Font = new Font("Microsoft YaHei", 9, FontStyle.Bold)
            };
            var line1 = new Label
            {
                Left = 8,
                Top = 33,
                Width = 172,
                Height = 18,
                Text = string.Format("占用：{0}/{1}（{2}%）", block.Occupied, block.Total, rate)
            };
            var line2 = new Label
            {
                Left = 8,
                Top = 55,
                Width = 172,
                Height = 18,
                Text = block.AbnormalCount > 0
                    ? string.Format("异常柜位：{0}", block.AbnormalCount)
                    : "异常柜位：0"
            };

            panel.Controls.Add(title);
            panel.Controls.Add(line1);
            panel.Controls.Add(line2);
            return panel;
        }

        private static Color GetHeatColor(decimal occupancyRate)
        {
            if (occupancyRate >= 0.9m)
            {
                return Color.FromArgb(224, 102, 102);
            }

            if (occupancyRate >= 0.7m)
            {
                return Color.FromArgb(246, 178, 107);
            }

            if (occupancyRate >= 0.5m)
            {
                return Color.FromArgb(255, 229, 153);
            }

            return Color.FromArgb(147, 196, 125);
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
