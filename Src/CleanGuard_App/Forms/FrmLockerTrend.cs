using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using CleanGuard_App.Utils;

namespace CleanGuard_App.Forms
{
    public class FrmLockerTrend : Form
    {
        private readonly NumericUpDown _numLimit = new NumericUpDown();
        private readonly Button _btnLoad = new Button();
        private readonly Chart _chart = new Chart();

        public FrmLockerTrend()
        {
            Text = "柜位占用趋势";
            Width = 960;
            Height = 620;
            StartPosition = FormStartPosition.CenterParent;

            InitializeLayout();
            LoadTrend();
        }

        private void InitializeLayout()
        {
            _numLimit.SetBounds(20, 15, 100, 28);
            _numLimit.Minimum = 10;
            _numLimit.Maximum = 500;
            _numLimit.Value = 100;
            Controls.Add(_numLimit);

            _btnLoad.Text = "加载趋势";
            _btnLoad.SetBounds(130, 15, 100, 28);
            _btnLoad.Click += (s, e) => LoadTrend();
            Controls.Add(_btnLoad);

            _chart.SetBounds(20, 55, 900, 520);
            var area = new ChartArea("MainArea");
            area.AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
            area.AxisX.LabelStyle.Angle = -45;
            _chart.ChartAreas.Add(area);
            _chart.Legends.Add(new Legend("Legend"));
            Controls.Add(_chart);
        }

        private void LoadTrend()
        {
            DataTable dt = SQLiteHelper.QueryLockerSnapshots((int)_numLimit.Value);
            _chart.Series.Clear();

            AddLineSeries(dt, "1F衣柜", "OneFClothesOccupied", Color.IndianRed);
            AddLineSeries(dt, "1F鞋柜", "OneFShoeOccupied", Color.OrangeRed);
            AddLineSeries(dt, "2F衣柜", "TwoFClothesOccupied", Color.SteelBlue);
            AddLineSeries(dt, "2F鞋柜", "TwoFShoeOccupied", Color.SeaGreen);
        }

        private void AddLineSeries(DataTable dt, string name, string field, Color color)
        {
            var series = new Series(name)
            {
                ChartType = SeriesChartType.Line,
                BorderWidth = 2,
                Color = color,
                XValueType = ChartValueType.String,
                ChartArea = "MainArea",
                Legend = "Legend"
            };

            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                DataRow row = dt.Rows[i];
                string x = Convert.ToString(row["SnapshotTime"]);
                int y = Convert.ToInt32(row[field]);
                series.Points.AddXY(x, y);
            }

            _chart.Series.Add(series);
        }
    }
}
