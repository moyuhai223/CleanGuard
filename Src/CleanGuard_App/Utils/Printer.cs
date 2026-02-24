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
                PayloadTop = 100
            };
        }

        private const string ConfigPrinterName = "Print.DefaultPrinter";
        private const string ConfigPaperName = "Print.DefaultPaper";
        private const string ConfigMarginLeft = "Print.MarginLeft";
        private const string ConfigMarginRight = "Print.MarginRight";
        private const string ConfigMarginTop = "Print.MarginTop";
        private const string ConfigMarginBottom = "Print.MarginBottom";
        private const string ConfigLandscape = "Print.Landscape";

        public static bool ConfigurePrintPreset(IWin32Window owner)
        {
            var printDoc = BuildPrintDocument("", "", "", "");
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

            return string.Format("打印机：{0}；纸张：{1}；方向：{2}",
                string.IsNullOrWhiteSpace(printerName) ? "未设置" : printerName,
                string.IsNullOrWhiteSpace(paperName) ? "未设置" : paperName,
                orientation);
        }


        public static string BuildQrPayload(string empNo, string name, string locker2F)
        {
            return string.Format("{0}|{1}|{2}", empNo ?? string.Empty, name ?? string.Empty, locker2F ?? string.Empty);
        }

        public static void ShowLabelPreview(string empNo, string name, string process, string locker2F)
        {
            var printDoc = BuildPrintDocument(empNo, name, process, locker2F);
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

        public static bool PrintLabelDirect(IWin32Window owner, string empNo, string name, string process, string locker2F)
        {
            var printDoc = BuildPrintDocument(empNo, name, process, locker2F);
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

        private static PrintDocument BuildPrintDocument(string empNo, string name, string process, string locker2F)
        {
            string payload = BuildQrPayload(empNo, name, locker2F);
            string title = string.Format("{0} ({1}) - {2}", name, empNo, process);
            LabelPrintLayoutSettings settings = GetDefaultLayoutSettings();

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

                    if (qrBitmap != null)
                    {
                        g.DrawImage(qrBitmap, bounds.Left + settings.QrLeft, bounds.Top + settings.QrTop, settings.QrSize, settings.QrSize);
                        g.DrawString("二维码内容：" + payload, bodyFont, Brushes.Black, bounds.Left + settings.PayloadLeft, bounds.Top + settings.PayloadTop);
                    }
                    else
                    {
                        g.DrawString("二维码内容：", bodyFont, Brushes.Black, bounds.Left + settings.BorderPadding, bounds.Top + settings.QrTop);
                        g.DrawString(payload, bodyFont, Brushes.Black, bounds.Left + settings.BorderPadding, bounds.Top + settings.QrTop + 25);
                        g.DrawString("（未检测到 QRCoder，当前为文本占位）", bodyFont, Brushes.Gray, bounds.Left + settings.BorderPadding, bounds.Top + settings.QrTop + 60);
                    }
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
