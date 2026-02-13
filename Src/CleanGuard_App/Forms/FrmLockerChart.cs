using System;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmLockerChart : Form
    {
        private readonly Chart _chart = new Chart();
        private readonly Label _lblSummary = new Label();

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
            _chart.Dock = DockStyle.Top;
            _chart.Height = 470;
            var area = new ChartArea("MainArea");
            _chart.ChartAreas.Add(area);
            _chart.Legends.Add(new Legend("Legend"));
            Controls.Add(_chart);

            _lblSummary.Dock = DockStyle.Fill;
            _lblSummary.Font = new Font("Microsoft YaHei", 10, FontStyle.Regular);
            _lblSummary.Padding = new Padding(10);
            Controls.Add(_lblSummary);
        }

        private void LoadData()
        {
            var summary = SQLiteHelper.GetLockerSummary();

            AddPieSeries("1F衣柜", summary.OneFClothesOccupied, summary.OneFClothesTotal);
            AddPieSeries("1F鞋柜", summary.OneFShoeOccupied, summary.OneFShoeTotal);
            AddPieSeries("2F衣柜", summary.TwoFClothesOccupied, summary.TwoFClothesTotal);
            AddPieSeries("2F鞋柜", summary.TwoFShoeOccupied, summary.TwoFShoeTotal);

            _lblSummary.Text =
                "占用概览\n" +
                BuildLine("1F衣柜", summary.OneFClothesOccupied, summary.OneFClothesTotal) +
                BuildLine("1F鞋柜", summary.OneFShoeOccupied, summary.OneFShoeTotal) +
                BuildLine("2F衣柜", summary.TwoFClothesOccupied, summary.TwoFClothesTotal) +
                BuildLine("2F鞋柜", summary.TwoFShoeOccupied, summary.TwoFShoeTotal);
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

        private static string BuildLine(string name, int occupied, int total)
        {
            decimal rate = total == 0 ? 0 : Math.Round(occupied * 100m / total, 1);
            return string.Format("{0}: {1}/{2}（占用率 {3}%）\n", name, occupied, total, rate);
        }
    }
}
