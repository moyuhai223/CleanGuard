using System;
using System;
using System.IO;
using System.Linq;
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.ThreadException += (sender, args) =>
            {
                MessageBox.Show(
                    "程序运行发生异常：\n" + args.Exception.Message + "\n\n" + args.Exception.StackTrace,
                    "CleanGuard 异常",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += (sender, args) =>
            {
                var ex = args.ExceptionObject as Exception;
                MessageBox.Show(
                    "程序发生严重异常：\n" + (ex != null ? ex.Message + "\n\n" + ex.StackTrace : args.ExceptionObject.ToString()),
                    "CleanGuard 异常",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            };

            try
            {
                SQLiteHelper.InitializeDatabase();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "数据库初始化失败，程序无法启动。\n\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "CleanGuard 启动失败",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            try
            {
                var mainForm = new FrmMain();
                mainForm.FormClosing += (sender, args) => BackupDatabase();
                Application.Run(mainForm);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "程序启动失败：\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "CleanGuard 启动失败",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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
