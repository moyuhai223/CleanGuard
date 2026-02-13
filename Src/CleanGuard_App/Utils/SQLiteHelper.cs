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
       CASE Status WHEN 1 THEN '在职' ELSE '离职' END AS 状态
FROM T_Employee
WHERE (@key = '')
   OR Name LIKE '%' || @key || '%'
   OR Pinyin LIKE '%' || @key || '%'
   OR Process LIKE '%' || @key || '%'
ORDER BY Status DESC, Name;";
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

        public static void AddEmployee(
            string empNo,
            string name,
            string process,
            string locker1FClothes,
            string locker1FShoe,
            string locker2FClothes,
            string locker2FShoe)
        {
            if (string.IsNullOrWhiteSpace(empNo) || string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("工号和姓名不能为空。");
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    EnsureLockerValid(conn, tx, locker1FClothes, "1F", "衣柜", "1F衣柜");
                    EnsureLockerValid(conn, tx, locker1FShoe, "1F", "鞋柜", "1F鞋柜");
                    EnsureLockerValid(conn, tx, locker2FClothes, "2F", "衣柜", "2F衣柜");
                    EnsureLockerValid(conn, tx, locker2FShoe, "2F", "鞋柜", "2F鞋柜");

                    using (var insertCmd = conn.CreateCommand())
                    {
                        insertCmd.Transaction = tx;
                        insertCmd.CommandText = @"INSERT INTO T_Employee
(EmpNo, Name, Pinyin, Process, Locker_1F_Clothes, Locker_1F_Shoe, Locker_2F_Clothes, Locker_2F_Shoe, Status)
VALUES (@EmpNo, @Name, @Pinyin, @Process, @L1C, @L1S, @L2C, @L2S, 1);";
                        insertCmd.Parameters.AddWithValue("@EmpNo", empNo.Trim());
                        insertCmd.Parameters.AddWithValue("@Name", name.Trim());
                        insertCmd.Parameters.AddWithValue("@Pinyin", PinYin.GetFirstLetter(name));
                        insertCmd.Parameters.AddWithValue("@Process", NormalizeNull(process));
                        insertCmd.Parameters.AddWithValue("@L1C", NormalizeNull(locker1FClothes));
                        insertCmd.Parameters.AddWithValue("@L1S", NormalizeNull(locker1FShoe));
                        insertCmd.Parameters.AddWithValue("@L2C", NormalizeNull(locker2FClothes));
                        insertCmd.Parameters.AddWithValue("@L2S", NormalizeNull(locker2FShoe));
                        insertCmd.ExecuteNonQuery();
                    }

                    MarkLockerOccupied(conn, tx, locker1FClothes, 1);
                    MarkLockerOccupied(conn, tx, locker1FShoe, 1);
                    MarkLockerOccupied(conn, tx, locker2FClothes, 1);
                    MarkLockerOccupied(conn, tx, locker2FShoe, 1);

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"新增员工成功: {empNo}-{name}");
        }

        public static EmployeeLockerInfo GetEmployeeLockerInfo(string empNo)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"SELECT Name, Locker_1F_Clothes, Locker_1F_Shoe, Locker_2F_Clothes, Locker_2F_Shoe, Status
FROM T_Employee WHERE EmpNo = @EmpNo";
                cmd.Parameters.AddWithValue("@EmpNo", empNo);
                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    return new EmployeeLockerInfo
                    {
                        EmpNo = empNo,
                        Name = reader["Name"].ToString(),
                        Locker1FClothes = reader["Locker_1F_Clothes"].ToString(),
                        Locker1FShoe = reader["Locker_1F_Shoe"].ToString(),
                        Locker2FClothes = reader["Locker_2F_Clothes"].ToString(),
                        Locker2FShoe = reader["Locker_2F_Shoe"].ToString(),
                        Status = Convert.ToInt32(reader["Status"])
                    };
                }
            }
        }

        public static void MarkEmployeeResigned(string empNo)
        {
            var info = GetEmployeeLockerInfo(empNo);
            if (info == null)
            {
                throw new InvalidOperationException("未找到该员工。");
            }

            if (info.Status == 0)
            {
                return;
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = "UPDATE T_Employee SET Status = 0 WHERE EmpNo = @EmpNo";
                        cmd.Parameters.AddWithValue("@EmpNo", empNo);
                        cmd.ExecuteNonQuery();
                    }

                    MarkLockerOccupied(conn, tx, info.Locker1FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker1FShoe, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FShoe, 0);

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"员工离职并释放柜位: {info.EmpNo}-{info.Name}");
        }

        public static void WriteSystemLog(string logType, string message)
        {
            ExecuteNonQuery("INSERT INTO T_SystemLog (LogType, Message) VALUES (@type, @message)",
                new SQLiteParameter("@type", logType),
                new SQLiteParameter("@message", message));
        }

        private static void EnsureLockerValid(SQLiteConnection conn, SQLiteTransaction tx, string lockerId, string location, string type, string displayName)
        {
            if (string.IsNullOrWhiteSpace(lockerId))
            {
                return;
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "SELECT Location, Type, IsOccupied FROM T_Lockers WHERE LockerID = @LockerID";
                cmd.Parameters.AddWithValue("@LockerID", lockerId);
                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        throw new InvalidOperationException($"{displayName} 柜号不存在: {lockerId}");
                    }

                    var actualLocation = reader["Location"].ToString();
                    var actualType = reader["Type"].ToString();
                    var occupied = Convert.ToInt32(reader["IsOccupied"]);

                    if (!string.Equals(actualLocation, location, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"{lockerId} 不属于 {location}。");
                    }

                    if (!string.Equals(actualType, type, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"{lockerId} 是{actualType}，不能分配给{displayName}。");
                    }

                    if (occupied == 1)
                    {
                        throw new InvalidOperationException($"{lockerId} 已被占用，不能重复分配。");
                    }
                }
            }
        }

        private static void MarkLockerOccupied(SQLiteConnection conn, SQLiteTransaction tx, string lockerId, int occupied)
        {
            if (string.IsNullOrWhiteSpace(lockerId))
            {
                return;
            }

            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "UPDATE T_Lockers SET IsOccupied = @Occupied WHERE LockerID = @LockerID";
                cmd.Parameters.AddWithValue("@Occupied", occupied);
                cmd.Parameters.AddWithValue("@LockerID", lockerId);
                cmd.ExecuteNonQuery();
            }
        }

        private static object NormalizeNull(string input)
        {
            return string.IsNullOrWhiteSpace(input) ? (object)DBNull.Value : input.Trim();
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

    public class EmployeeLockerInfo
    {
        public string EmpNo { get; set; }
        public string Name { get; set; }
        public string Locker1FClothes { get; set; }
        public string Locker1FShoe { get; set; }
        public string Locker2FClothes { get; set; }
        public string Locker2FShoe { get; set; }
        public int Status { get; set; }
    }
}
