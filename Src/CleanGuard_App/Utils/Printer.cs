using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CleanGuard_App.Utils
{
    public static class Printer
    {
        public static string BuildQrPayload(string empNo, string name, string locker2F)
        {
            return string.Format("{0}|{1}|{2}", empNo ?? string.Empty, name ?? string.Empty, locker2F ?? string.Empty);
        }

        public static void ShowLabelPreview(string empNo, string name, string process, string locker2F)
        {
            string payload = BuildQrPayload(empNo, name, locker2F);
            string title = string.Format("{0} ({1}) - {2}", name, empNo, process);

            var printDoc = new PrintDocument();
            printDoc.DocumentName = "CleanGuard_员工标签";
            printDoc.PrintPage += (sender, args) =>
            {
                Graphics g = args.Graphics;
                Rectangle bounds = args.MarginBounds;

                using (var titleFont = new Font("Microsoft YaHei", 12, FontStyle.Bold))
                using (var bodyFont = new Font("Microsoft YaHei", 10, FontStyle.Regular))
                using (var pen = new Pen(Color.Black, 1))
                {
                    g.DrawRectangle(pen, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
                    g.DrawString("CleanGuard 员工标签", titleFont, Brushes.Black, bounds.Left + 10, bounds.Top + 10);
                    g.DrawString(title, bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 45);
                    g.DrawString("二维码内容：", bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 80);
                    g.DrawString(payload, bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 105);
                    g.DrawString("（当前为文本占位，后续接入 QRCoder 图像）", bodyFont, Brushes.Gray, bounds.Left + 10, bounds.Top + 140);
                }

                args.HasMorePages = false;
            };

            var preview = new PrintPreviewDialog();
            preview.Document = printDoc;
            preview.Width = 900;
            preview.Height = 700;
            preview.ShowDialog();
        }
    }
}
