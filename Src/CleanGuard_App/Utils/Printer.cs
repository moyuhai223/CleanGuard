using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace CleanGuard_App.Utils
{
    public static class Printer
    {
        public class LabelPrintLayoutSettings
        {
            public int BorderPadding { get; set; }
            public int TitleTop { get; set; }
            public int HeaderTop { get; set; }
            public int QrTop { get; set; }
            public int QrLeft { get; set; }
            public int QrSize { get; set; }
            public int PayloadLeft { get; set; }
            public int PayloadTop { get; set; }
            public int LockerTagWidth { get; set; }
            public int LockerTagHeight { get; set; }
            public int LockerTagGapX { get; set; }
            public int LockerTagGapY { get; set; }
            public int LockerTagTop { get; set; }
        }

        public static LabelPrintLayoutSettings GetDefaultLayoutSettings()
        {
            return new LabelPrintLayoutSettings
            {
                BorderPadding = 10,
                TitleTop = 10,
                HeaderTop = 45,
                QrTop = 80,
                QrLeft = 10,
                QrSize = 140,
                PayloadLeft = 170,
                PayloadTop = 100,
                LockerTagWidth = 180,
                LockerTagHeight = 95,
                LockerTagGapX = 20,
                LockerTagGapY = 16,
                LockerTagTop = 250
            };
        }

        private const string ConfigPrinterName = "Print.DefaultPrinter";
        private const string ConfigPaperName = "Print.DefaultPaper";
        private const string ConfigMarginLeft = "Print.MarginLeft";
        private const string ConfigMarginRight = "Print.MarginRight";
        private const string ConfigMarginTop = "Print.MarginTop";
        private const string ConfigMarginBottom = "Print.MarginBottom";
        private const string ConfigLandscape = "Print.Landscape";
        private const string ConfigLockerTagWidth = "Print.LockerTagWidth";
        private const string ConfigLockerTagHeight = "Print.LockerTagHeight";

        public static bool ConfigurePrintPreset(IWin32Window owner)
        {
            var printDoc = BuildPrintDocument("", "", "", "", "", "", "");
            try
            {
                ApplySavedPreset(printDoc);
                using (var pageSetup = new PageSetupDialog())
                {
                    pageSetup.Document = printDoc;
                    pageSetup.AllowMargins = true;
                    pageSetup.AllowOrientation = true;
                    pageSetup.AllowPaper = true;
                    pageSetup.AllowPrinter = true;
                    pageSetup.ShowNetwork = true;
                    if (pageSetup.ShowDialog(owner) != DialogResult.OK)
                    {
                        return false;
                    }

                    SavePrintPreset(printDoc.PrinterSettings, printDoc.DefaultPageSettings.PaperSize, printDoc.DefaultPageSettings.Margins, printDoc.DefaultPageSettings.Landscape);
                    int width = GetEffectiveLayoutSettings().LockerTagWidth;
                    int height = GetEffectiveLayoutSettings().LockerTagHeight;
                    PromptLockerTagSize(owner, ref width, ref height);
                    SQLiteHelper.SaveSystemConfigValue(ConfigLockerTagWidth, width.ToString());
                    SQLiteHelper.SaveSystemConfigValue(ConfigLockerTagHeight, height.ToString());
                    return true;
                }
            }
            finally
            {
                printDoc.Dispose();
            }
        }

        public static string GetPrintPresetSummary()
        {
            string printerName = SQLiteHelper.GetSystemConfigValue(ConfigPrinterName);
            string paperName = SQLiteHelper.GetSystemConfigValue(ConfigPaperName);
            bool landscape;
            string orientation = TryReadConfigBool(ConfigLandscape, out landscape)
                ? (landscape ? "横向" : "纵向")
                : "未设置";

            var layout = GetEffectiveLayoutSettings();
            return string.Format("打印机：{0}；纸张：{1}；方向：{2}；标签尺寸：{3}x{4}",
                string.IsNullOrWhiteSpace(printerName) ? "未设置" : printerName,
                string.IsNullOrWhiteSpace(paperName) ? "未设置" : paperName,
                orientation,
                layout.LockerTagWidth,
                layout.LockerTagHeight);
        }


        public static string BuildQrPayload(string empNo, string name, string locker1FClothes, string locker1FShoe, string locker2FClothes, string locker2FShoe)
        {
            return string.Format("{0}|{1}|1F衣柜:{2}|1F鞋柜:{3}|2F衣柜:{4}|2F鞋柜:{5}",
                empNo ?? string.Empty,
                name ?? string.Empty,
                locker1FClothes ?? string.Empty,
                locker1FShoe ?? string.Empty,
                locker2FClothes ?? string.Empty,
                locker2FShoe ?? string.Empty);
        }

        public static void ShowLabelPreview(string empNo, string name, string process, string locker1FClothes, string locker1FShoe, string locker2FClothes, string locker2FShoe)
        {
            var printDoc = BuildPrintDocument(empNo, name, process, locker1FClothes, locker1FShoe, locker2FClothes, locker2FShoe);
            ApplySavedPreset(printDoc);
            using (var preview = new PrintPreviewDialog())
            {
                preview.Document = printDoc;
                preview.Width = 900;
                preview.Height = 700;
                preview.ShowDialog();
            }

            printDoc.Dispose();
        }

        public static bool PrintLabelDirect(IWin32Window owner, string empNo, string name, string process, string locker1FClothes, string locker1FShoe, string locker2FClothes, string locker2FShoe)
        {
            var printDoc = BuildPrintDocument(empNo, name, process, locker1FClothes, locker1FShoe, locker2FClothes, locker2FShoe);
            ApplySavedPreset(printDoc);
            using (var dialog = new PrintDialog())
            {
                dialog.AllowSelection = false;
                dialog.AllowSomePages = false;
                dialog.UseEXDialog = true;
                dialog.Document = printDoc;
                if (dialog.ShowDialog(owner) != DialogResult.OK)
                {
                    printDoc.Dispose();
                    return false;
                }

                printDoc.PrinterSettings = dialog.PrinterSettings;
                SavePrintPreset(printDoc.PrinterSettings, printDoc.DefaultPageSettings.PaperSize, printDoc.DefaultPageSettings.Margins, printDoc.DefaultPageSettings.Landscape);
            }

            printDoc.Print();
            printDoc.Dispose();
            return true;
        }

        private static PrintDocument BuildPrintDocument(string empNo, string name, string process, string locker1FClothes, string locker1FShoe, string locker2FClothes, string locker2FShoe)
        {
            string payload = BuildQrPayload(empNo, name, locker1FClothes, locker1FShoe, locker2FClothes, locker2FShoe);
            string title = string.Format("{0} ({1}) - {2}", name, empNo, process);
            LabelPrintLayoutSettings settings = GetEffectiveLayoutSettings();

            var printDoc = new PrintDocument();
            printDoc.DocumentName = "CleanGuard_员工标签";
            printDoc.PrintPage += (sender, args) =>
            {
                Graphics g = args.Graphics;
                Rectangle bounds = args.MarginBounds;

                using (var titleFont = new Font("Microsoft YaHei", 12, FontStyle.Bold))
                using (var bodyFont = new Font("Microsoft YaHei", 10, FontStyle.Regular))
                using (var pen = new Pen(Color.Black, 1))
                using (var qrBitmap = TryGenerateQrBitmap(payload))
                {
                    g.DrawRectangle(pen, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
                    g.DrawString("CleanGuard 员工标签", titleFont, Brushes.Black, bounds.Left + settings.BorderPadding, bounds.Top + settings.TitleTop);
                    g.DrawString(title, bodyFont, Brushes.Black, bounds.Left + settings.BorderPadding, bounds.Top + settings.HeaderTop);

                    int qrX = bounds.Right - settings.BorderPadding - settings.QrSize;
                    int qrY = bounds.Top + settings.TitleTop;
                    if (qrBitmap != null)
                    {
                        g.DrawImage(qrBitmap, qrX, qrY, settings.QrSize, settings.QrSize);
                    }
                    else
                    {
                        g.DrawString("二维码内容：", bodyFont, Brushes.Black, qrX, qrY);
                        g.DrawString(payload, bodyFont, Brushes.Black, qrX, qrY + 20);
                        g.DrawString("（未检测到 QRCoder，当前为文本占位）", bodyFont, Brushes.Gray, qrX, qrY + 55);
                    }

                    int tagTop = bounds.Top + settings.LockerTagTop;
                    int tagLeft = bounds.Left + settings.BorderPadding;
                    var tag1 = new Rectangle(tagLeft, tagTop, settings.LockerTagWidth, settings.LockerTagHeight);
                    var tag2 = new Rectangle(tagLeft + settings.LockerTagWidth + settings.LockerTagGapX, tagTop, settings.LockerTagWidth, settings.LockerTagHeight);
                    var tag3 = new Rectangle(tagLeft, tagTop + settings.LockerTagHeight + settings.LockerTagGapY, settings.LockerTagWidth, settings.LockerTagHeight);
                    var tag4 = new Rectangle(tagLeft + settings.LockerTagWidth + settings.LockerTagGapX, tagTop + settings.LockerTagHeight + settings.LockerTagGapY, settings.LockerTagWidth, settings.LockerTagHeight);

                    DrawLockerTag(g, tag1, "1F衣柜", name, process, locker1FClothes, titleFont, bodyFont);
                    DrawLockerTag(g, tag2, "1F鞋柜", name, process, locker1FShoe, titleFont, bodyFont);
                    DrawLockerTag(g, tag3, "2F衣柜", name, process, locker2FClothes, titleFont, bodyFont);
                    DrawLockerTag(g, tag4, "2F鞋柜", name, process, locker2FShoe, titleFont, bodyFont);
                }

                args.HasMorePages = false;
            };

            return printDoc;
        }

        private static void ApplySavedPreset(PrintDocument printDoc)
        {
            if (printDoc == null)
            {
                return;
            }

            string printerName = SQLiteHelper.GetSystemConfigValue(ConfigPrinterName);
            string paperName = SQLiteHelper.GetSystemConfigValue(ConfigPaperName);

            if (!string.IsNullOrWhiteSpace(printerName))
            {
                foreach (string installed in PrinterSettings.InstalledPrinters)
                {
                    if (string.Equals(installed, printerName, StringComparison.OrdinalIgnoreCase))
                    {
                        printDoc.PrinterSettings.PrinterName = installed;
                        break;
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(paperName))
            {
                foreach (PaperSize paper in printDoc.PrinterSettings.PaperSizes)
                {
                    if (string.Equals(paper.PaperName, paperName, StringComparison.OrdinalIgnoreCase))
                    {
                        printDoc.DefaultPageSettings.PaperSize = paper;
                        break;
                    }
                }
            }

            var margins = printDoc.DefaultPageSettings.Margins;
            int left;
            int right;
            int top;
            int bottom;
            if (TryReadConfigInt(ConfigMarginLeft, out left))
            {
                margins.Left = left;
            }
            if (TryReadConfigInt(ConfigMarginRight, out right))
            {
                margins.Right = right;
            }
            if (TryReadConfigInt(ConfigMarginTop, out top))
            {
                margins.Top = top;
            }
            if (TryReadConfigInt(ConfigMarginBottom, out bottom))
            {
                margins.Bottom = bottom;
            }

            bool landscape;
            if (TryReadConfigBool(ConfigLandscape, out landscape))
            {
                printDoc.DefaultPageSettings.Landscape = landscape;
            }
        }

        private static void SavePrintPreset(PrinterSettings printerSettings, PaperSize paperSize, Margins margins, bool landscape)
        {
            if (printerSettings != null && !string.IsNullOrWhiteSpace(printerSettings.PrinterName))
            {
                SQLiteHelper.SaveSystemConfigValue(ConfigPrinterName, printerSettings.PrinterName);
            }

            if (paperSize != null && !string.IsNullOrWhiteSpace(paperSize.PaperName))
            {
                SQLiteHelper.SaveSystemConfigValue(ConfigPaperName, paperSize.PaperName);
            }

            if (margins != null)
            {
                SQLiteHelper.SaveSystemConfigValue(ConfigMarginLeft, margins.Left.ToString());
                SQLiteHelper.SaveSystemConfigValue(ConfigMarginRight, margins.Right.ToString());
                SQLiteHelper.SaveSystemConfigValue(ConfigMarginTop, margins.Top.ToString());
                SQLiteHelper.SaveSystemConfigValue(ConfigMarginBottom, margins.Bottom.ToString());
            }

            SQLiteHelper.SaveSystemConfigValue(ConfigLandscape, landscape ? "1" : "0");

        }

        private static LabelPrintLayoutSettings GetEffectiveLayoutSettings()
        {
            var settings = GetDefaultLayoutSettings();
            int width;
            int height;
            if (TryReadConfigInt(ConfigLockerTagWidth, out width) && width > 60)
            {
                settings.LockerTagWidth = width;
            }
            if (TryReadConfigInt(ConfigLockerTagHeight, out height) && height > 40)
            {
                settings.LockerTagHeight = height;
            }

            return settings;
        }

        private static void PromptLockerTagSize(IWin32Window owner, ref int width, ref int height)
        {
            using (var form = new Form())
            {
                form.Text = "标签尺寸设置";
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.Width = 320;
                form.Height = 190;

                var lblW = new Label { Text = "标签宽度", Left = 20, Top = 25, Width = 80 };
                var numW = new NumericUpDown { Left = 110, Top = 20, Width = 160, Minimum = 100, Maximum = 400, Value = Math.Max(100, Math.Min(400, width)) };
                var lblH = new Label { Text = "标签高度", Left = 20, Top = 65, Width = 80 };
                var numH = new NumericUpDown { Left = 110, Top = 60, Width = 160, Minimum = 60, Maximum = 220, Value = Math.Max(60, Math.Min(220, height)) };

                var btnOk = new Button { Text = "确定", Left = 110, Top = 105, Width = 75, DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "取消", Left = 195, Top = 105, Width = 75, DialogResult = DialogResult.Cancel };

                form.Controls.Add(lblW);
                form.Controls.Add(numW);
                form.Controls.Add(lblH);
                form.Controls.Add(numH);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog(owner) == DialogResult.OK)
                {
                    width = Convert.ToInt32(numW.Value);
                    height = Convert.ToInt32(numH.Value);
                }
            }
        }

        private static void DrawLockerTag(Graphics g, Rectangle rect, string title, string name, string process, string lockerNo, Font titleFont, Font bodyFont)
        {
            g.DrawRectangle(Pens.Black, rect);
            g.DrawString(string.Format("{0} / {1}", name, process), titleFont, Brushes.Black, rect.Left + 6, rect.Top + 6);
            g.DrawString(title + "：" + (string.IsNullOrWhiteSpace(lockerNo) ? "(空)" : lockerNo), bodyFont, Brushes.Black, rect.Left + 6, rect.Top + rect.Height - 28);
        }

        private static bool TryReadConfigInt(string key, out int value)
        {
            value = 0;
            string raw = SQLiteHelper.GetSystemConfigValue(key);
            return int.TryParse(raw, out value) && value >= 0;
        }

        private static bool TryReadConfigBool(string key, out bool value)
        {
            value = false;
            string raw = SQLiteHelper.GetSystemConfigValue(key);
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            if (raw == "1")
            {
                value = true;
                return true;
            }

            if (raw == "0")
            {
                value = false;
                return true;
            }

            return bool.TryParse(raw, out value);
        }

        private static Bitmap TryGenerateQrBitmap(string payload)
        {
            try
            {
                Type qrGeneratorType = Type.GetType("QRCoder.QRCodeGenerator, QRCoder", false);
                if (qrGeneratorType == null)
                {
                    return null;
                }

                object generator = Activator.CreateInstance(qrGeneratorType);
                Type eccLevelType = qrGeneratorType.GetNestedType("ECCLevel");
                object ecc = Enum.Parse(eccLevelType, "Q");

                object qrData = qrGeneratorType.GetMethod("CreateQrCode", new[] { typeof(string), eccLevelType })
                    .Invoke(generator, new[] { payload, ecc });

                Type qrCodeType = Type.GetType("QRCoder.QRCode, QRCoder", false);
                if (qrCodeType == null)
                {
                    DisposeIfNeeded(qrData);
                    DisposeIfNeeded(generator);
                    return null;
                }

                object qrCode = Activator.CreateInstance(qrCodeType, qrData);
                object graphic = qrCodeType.GetMethod("GetGraphic", new[] { typeof(int) }).Invoke(qrCode, new object[] { 8 });

                DisposeIfNeeded(qrCode);
                DisposeIfNeeded(qrData);
                DisposeIfNeeded(generator);

                return graphic as Bitmap;
            }
            catch
            {
                return null;
            }
        }

        private static void DisposeIfNeeded(object instance)
        {
            var disposable = instance as IDisposable;
            if (disposable != null)
            {
                disposable.Dispose();
            }
        }
    }
}
