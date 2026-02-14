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
            var printDoc = BuildPrintDocument(empNo, name, process, locker2F);
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
            }

            printDoc.Print();
            printDoc.Dispose();
            return true;
        }

        private static PrintDocument BuildPrintDocument(string empNo, string name, string process, string locker2F)
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
                using (var qrBitmap = TryGenerateQrBitmap(payload))
                {
                    g.DrawRectangle(pen, bounds.Left, bounds.Top, bounds.Width, bounds.Height);
                    g.DrawString("CleanGuard 员工标签", titleFont, Brushes.Black, bounds.Left + 10, bounds.Top + 10);
                    g.DrawString(title, bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 45);

                    if (qrBitmap != null)
                    {
                        g.DrawImage(qrBitmap, bounds.Left + 10, bounds.Top + 80, 140, 140);
                        g.DrawString("二维码内容：" + payload, bodyFont, Brushes.Black, bounds.Left + 170, bounds.Top + 100);
                    }
                    else
                    {
                        g.DrawString("二维码内容：", bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 80);
                        g.DrawString(payload, bodyFont, Brushes.Black, bounds.Left + 10, bounds.Top + 105);
                        g.DrawString("（未检测到 QRCoder，当前为文本占位）", bodyFont, Brushes.Gray, bounds.Left + 10, bounds.Top + 140);
                    }
                }

                args.HasMorePages = false;
            };

            return printDoc;
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
