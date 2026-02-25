using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CleanGuard_App.Forms;
using CleanGuard_App.Utils;

namespace CleanGuard_App
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            RegisterGlobalExceptionHandlers();

            try
            {
                ValidateRuntimeDependencies();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                SQLiteHelper.InitializeDatabase();

                var mainForm = new FrmMain();
                mainForm.FormClosing += (sender, args) => BackupDatabase();

                Application.Run(mainForm);
            }
            catch (Exception ex)
            {
                ShowStartupError("程序启动失败", ex);
            }
        }

        private static void RegisterGlobalExceptionHandlers()
        {
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (sender, args) => ShowStartupError("程序发生未处理异常", args.Exception);
            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                var exception = args.ExceptionObject as Exception;
                if (exception == null)
                {
                    exception = new Exception(Convert.ToString(args.ExceptionObject));
                }

                ShowStartupError("程序发生未处理异常", exception);
            };
        }

        private static void ValidateRuntimeDependencies()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string[] requiredFiles =
            {
                "System.Data.SQLite.dll",
                "NPOI.dll",
                "NPOI.OOXML.dll",
                "NPOI.OpenXml4Net.dll",
                "NPOI.OpenXmlFormats.dll"
            };

            var missing = requiredFiles.Where(file => !File.Exists(Path.Combine(baseDir, file))).ToArray();
            if (missing.Length > 0)
            {
                throw new FileNotFoundException("程序依赖文件缺失：" + string.Join(", ", missing));
            }

            bool hasX86Interop = File.Exists(Path.Combine(baseDir, "x86", "SQLite.Interop.dll"));
            bool hasX64Interop = File.Exists(Path.Combine(baseDir, "x64", "SQLite.Interop.dll"));
            if (!hasX86Interop || !hasX64Interop)
            {
                throw new FileNotFoundException("SQLite 运行库缺失：请确认 x86/x64 目录中的 SQLite.Interop.dll 已随程序一同部署。");
            }
        }

        private static void ShowStartupError(string title, Exception ex)
        {
            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var logPath = Path.Combine(baseDir, "startup-error.log");

                var builder = new StringBuilder();
                builder.AppendLine("[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + title);
                builder.AppendLine(ex.ToString());
                builder.AppendLine(new string('-', 80));

                File.AppendAllText(logPath, builder.ToString(), Encoding.UTF8);

                MessageBox.Show(
                    title + "。\n\n" + ex.Message + "\n\n详细错误已记录：" + logPath,
                    "CleanGuard 启动错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch
            {
                // 避免异常处理再次抛错导致进程直接退出。
            }
        }

        private static void BackupDatabase()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var source = Path.Combine(baseDir, "CleanGuard.db");
            if (!File.Exists(source))
            {
                return;
            }

            var backupFolder = Path.Combine(baseDir, "Backup");
            if (!Directory.Exists(backupFolder))
            {
                Directory.CreateDirectory(backupFolder);
            }

            var fileName = $"Backup_{DateTime.Now:yyyyMMdd_HHmmss}.db";
            var destination = Path.Combine(backupFolder, fileName);
            File.Copy(source, destination, true);

            var expiredDate = DateTime.Now.AddDays(-7);
            var oldFiles = Directory.GetFiles(backupFolder, "Backup_*.db")
                .Select(path => new FileInfo(path))
                .Where(info => info.CreationTime < expiredDate);

            foreach (var file in oldFiles)
            {
                file.Delete();
            }

            SQLiteHelper.WriteSystemLog("Backup", $"自动备份成功: {fileName}");
        }
    }
}
