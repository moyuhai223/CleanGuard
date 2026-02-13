using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;

namespace CleanGuard_App.Utils
{
    public static class SQLiteHelper
    {
        private static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CleanGuard.db");
        private static readonly string ConnectionString = $"Data Source={DbPath};Version=3;";

        public static void InitializeDatabase()
        {
            if (!File.Exists(DbPath))
            {
                SQLiteConnection.CreateFile(DbPath);
            }

            ExecuteNonQuery(@"
CREATE TABLE IF NOT EXISTS T_Lockers (
    LockerID TEXT PRIMARY KEY,
    Location TEXT NOT NULL,
    Type TEXT NOT NULL,
    IsOccupied INTEGER DEFAULT 0,
    Remark TEXT
);
CREATE INDEX IF NOT EXISTS idx_locker_filter ON T_Lockers(Location, Type, IsOccupied);

CREATE TABLE IF NOT EXISTS T_Process (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    Name TEXT NOT NULL UNIQUE
);
INSERT OR IGNORE INTO T_Process (Name) VALUES ('热切'), ('检包'), ('整备'), ('窑炉'), ('通道'), ('成型');

CREATE TABLE IF NOT EXISTS T_Employee (
    ID INTEGER PRIMARY KEY AUTOINCREMENT,
    EmpNo TEXT NOT NULL UNIQUE,
    Name TEXT NOT NULL,
    Pinyin TEXT,
    Process TEXT,
    Locker_1F_Clothes TEXT,
    Locker_1F_Shoe TEXT,
    Locker_2F_Clothes TEXT,
    Locker_2F_Shoe TEXT,
    Status INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS T_Emp_Items (
    ItemID INTEGER PRIMARY KEY AUTOINCREMENT,
    EmpID INTEGER NOT NULL,
    Category TEXT NOT NULL,
    SlotIndex INTEGER,
    Size TEXT,
    IssueDate TEXT,
    FOREIGN KEY(EmpID) REFERENCES T_Employee(ID) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS T_SystemLog (
    LogID INTEGER PRIMARY KEY AUTOINCREMENT,
    LogType TEXT,
    Message TEXT,
    LogTime DATETIME DEFAULT CURRENT_TIMESTAMP
);
");
        }

        public static DataTable QueryEmployees(string keyword)
        {
            var table = new DataTable();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            using (var adapter = new SQLiteDataAdapter(cmd))
            {
                conn.Open();
                cmd.CommandText = @"
SELECT EmpNo AS 工号, Name AS 姓名, Process AS 工序,
       Locker_1F_Clothes AS [1F衣柜], Locker_1F_Shoe AS [1F鞋柜],
       Locker_2F_Clothes AS [2F衣柜], Locker_2F_Shoe AS [2F鞋柜],
       Status AS 状态
FROM T_Employee
WHERE (@key = '')
   OR Name LIKE '%' || @key || '%'
   OR Pinyin LIKE '%' || @key || '%'
   OR Process LIKE '%' || @key || '%'
ORDER BY Name;";
                cmd.Parameters.AddWithValue("@key", keyword ?? string.Empty);
                adapter.Fill(table);
            }

            return table;
        }

        public static string[] GetProcesses()
        {
            var list = new List<string>();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "SELECT Name FROM T_Process ORDER BY ID";
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(reader.GetString(0));
                    }
                }
            }

            return list.ToArray();
        }

        public static string[] GetAvailableLockers(string location, string type)
        {
            var list = new List<string>();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"SELECT LockerID FROM T_Lockers
WHERE Location = @location AND Type = @type AND IsOccupied = 0
ORDER BY LockerID";
                cmd.Parameters.AddWithValue("@location", location);
                cmd.Parameters.AddWithValue("@type", type);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(reader.GetString(0));
                    }
                }
            }

            return list.ToArray();
        }

        public static void WriteSystemLog(string logType, string message)
        {
            ExecuteNonQuery("INSERT INTO T_SystemLog (LogType, Message) VALUES (@type, @message)",
                new SQLiteParameter("@type", logType),
                new SQLiteParameter("@message", message));
        }

        private static void ExecuteNonQuery(string sql, params SQLiteParameter[] parameters)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = sql;
                if (parameters != null && parameters.Length > 0)
                {
                    cmd.Parameters.AddRange(parameters);
                }

                cmd.ExecuteNonQuery();
            }
        }
    }
}
