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

            SQLiteHelper.InitializeDatabase();

            var mainForm = new FrmMain();
            mainForm.FormClosing += (sender, args) => BackupDatabase();

            Application.Run(mainForm);
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
