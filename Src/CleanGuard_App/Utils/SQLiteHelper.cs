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

CREATE TABLE IF NOT EXISTS T_LockerSnapshot (
    SnapshotID INTEGER PRIMARY KEY AUTOINCREMENT,
    SnapshotTime DATETIME DEFAULT CURRENT_TIMESTAMP,
    OneFClothesOccupied INTEGER,
    OneFClothesTotal INTEGER,
    OneFShoeOccupied INTEGER,
    OneFShoeTotal INTEGER,
    TwoFClothesOccupied INTEGER,
    TwoFClothesTotal INTEGER,
    TwoFShoeOccupied INTEGER,
    TwoFShoeTotal INTEGER,
    Source TEXT
);
");

            SeedDefaultLockers();
            CaptureLockerSnapshot("Startup");
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

        public static DataTable QueryProcessesTable()
        {
            var table = new DataTable();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            using (var adapter = new SQLiteDataAdapter(cmd))
            {
                conn.Open();
                cmd.CommandText = @"SELECT ID AS 编号, Name AS 工序名称 FROM T_Process ORDER BY ID";
                adapter.Fill(table);
            }

            return table;
        }

        public static void AddProcess(string processName)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                throw new ArgumentException("工序名称不能为空。");
            }

            string name = processName.Trim();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "INSERT INTO T_Process(Name) VALUES(@Name)";
                cmd.Parameters.AddWithValue("@Name", name);
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (SQLiteException ex)
                {
                    if (ex.ResultCode == SQLiteErrorCode.Constraint)
                    {
                        throw new InvalidOperationException("工序已存在，不能重复新增。");
                    }

                    throw;
                }
            }

            WriteSystemLog("Employee", "新增工序字典: " + name);
        }

        public static void DeleteProcess(string processName)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                throw new ArgumentException("工序名称不能为空。");
            }

            string name = processName.Trim();
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();

                using (var check = conn.CreateCommand())
                {
                    check.CommandText = "SELECT COUNT(1) FROM T_Employee WHERE Process = @Name";
                    check.Parameters.AddWithValue("@Name", name);
                    int used = Convert.ToInt32(check.ExecuteScalar());
                    if (used > 0)
                    {
                        throw new InvalidOperationException("该工序正在被员工使用，无法删除。");
                    }
                }

                using (var del = conn.CreateCommand())
                {
                    del.CommandText = "DELETE FROM T_Process WHERE Name = @Name";
                    del.Parameters.AddWithValue("@Name", name);
                    int affected = del.ExecuteNonQuery();
                    if (affected <= 0)
                    {
                        throw new InvalidOperationException("未找到该工序，无法删除。");
                    }
                }
            }

            WriteSystemLog("Employee", "删除工序字典: " + name);
        }

        public static void RenameProcess(string oldName, string newName)
        {
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName))
            {
                throw new ArgumentException("原工序和新工序名称均不能为空。");
            }

            string from = oldName.Trim();
            string to = newName.Trim();
            if (string.Equals(from, to, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("新工序名称与原名称相同，无需修改。");
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    using (var checkOld = conn.CreateCommand())
                    {
                        checkOld.Transaction = tx;
                        checkOld.CommandText = "SELECT COUNT(1) FROM T_Process WHERE Name = @Name";
                        checkOld.Parameters.AddWithValue("@Name", from);
                        int exists = Convert.ToInt32(checkOld.ExecuteScalar());
                        if (exists <= 0)
                        {
                            throw new InvalidOperationException("未找到原工序，无法重命名。");
                        }
                    }

                    using (var checkNew = conn.CreateCommand())
                    {
                        checkNew.Transaction = tx;
                        checkNew.CommandText = "SELECT COUNT(1) FROM T_Process WHERE Name = @Name";
                        checkNew.Parameters.AddWithValue("@Name", to);
                        int exists = Convert.ToInt32(checkNew.ExecuteScalar());
                        if (exists > 0)
                        {
                            throw new InvalidOperationException("新工序名称已存在，请更换后重试。");
                        }
                    }

                    using (var updProcess = conn.CreateCommand())
                    {
                        updProcess.Transaction = tx;
                        updProcess.CommandText = "UPDATE T_Process SET Name = @NewName WHERE Name = @OldName";
                        updProcess.Parameters.AddWithValue("@NewName", to);
                        updProcess.Parameters.AddWithValue("@OldName", from);
                        updProcess.ExecuteNonQuery();
                    }

                    using (var updEmployees = conn.CreateCommand())
                    {
                        updEmployees.Transaction = tx;
                        updEmployees.CommandText = "UPDATE T_Employee SET Process = @NewName WHERE Process = @OldName";
                        updEmployees.Parameters.AddWithValue("@NewName", to);
                        updEmployees.Parameters.AddWithValue("@OldName", from);
                        updEmployees.ExecuteNonQuery();
                    }

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", "重命名工序字典: " + from + " -> " + to);
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

        public static string[] GetAvailableLockersIncluding(string location, string type, string selectedLockerId)
        {
            var list = new List<string>(GetAvailableLockers(location, type));
            if (!string.IsNullOrWhiteSpace(selectedLockerId) && !list.Contains(selectedLockerId))
            {
                list.Add(selectedLockerId);
                list.Sort(StringComparer.OrdinalIgnoreCase);
            }

            return list.ToArray();
        }


        public static LockerSummary GetLockerSummary()
        {
            var summary = new LockerSummary();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"
SELECT Location, Type,
       SUM(CASE WHEN IsOccupied = 1 THEN 1 ELSE 0 END) AS OccupiedCount,
       COUNT(1) AS TotalCount
FROM T_Lockers
GROUP BY Location, Type;";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string location = reader["Location"].ToString();
                        string type = reader["Type"].ToString();
                        int occupied = Convert.ToInt32(reader["OccupiedCount"]);
                        int total = Convert.ToInt32(reader["TotalCount"]);

                        summary.Set(location, type, occupied, total);
                    }
                }
            }

            return summary;
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

            try
            {
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
            }
            catch (SQLiteException ex)
            {
                if (ex.ResultCode == SQLiteErrorCode.Constraint)
                {
                    throw new InvalidOperationException("工号已存在，请检查后重试。");
                }

                throw;
            }

            WriteSystemLog("Employee", $"新增员工成功: {empNo}-{name}");
            CaptureLockerSnapshot("AddEmployee");
        }

        public static EmployeeEditModel GetEmployeeEditModel(string empNo)
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"SELECT EmpNo, Name, Process, Locker_1F_Clothes, Locker_1F_Shoe, Locker_2F_Clothes, Locker_2F_Shoe, Status
FROM T_Employee WHERE EmpNo = @EmpNo";
                cmd.Parameters.AddWithValue("@EmpNo", empNo);
                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    return new EmployeeEditModel
                    {
                        EmpNo = reader["EmpNo"].ToString(),
                        Name = reader["Name"].ToString(),
                        Process = reader["Process"].ToString(),
                        Locker1FClothes = reader["Locker_1F_Clothes"].ToString(),
                        Locker1FShoe = reader["Locker_1F_Shoe"].ToString(),
                        Locker2FClothes = reader["Locker_2F_Clothes"].ToString(),
                        Locker2FShoe = reader["Locker_2F_Shoe"].ToString(),
                        Status = Convert.ToInt32(reader["Status"])
                    };
                }
            }
        }

        public static void UpdateEmployee(
            string empNo,
            string name,
            string process,
            string locker1FClothes,
            string locker1FShoe,
            string locker2FClothes,
            string locker2FShoe)
        {
            var oldData = GetEmployeeEditModel(empNo);
            if (oldData == null)
            {
                throw new InvalidOperationException("未找到该员工，无法更新。");
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    MarkLockerOccupied(conn, tx, oldData.Locker1FClothes, 0);
                    MarkLockerOccupied(conn, tx, oldData.Locker1FShoe, 0);
                    MarkLockerOccupied(conn, tx, oldData.Locker2FClothes, 0);
                    MarkLockerOccupied(conn, tx, oldData.Locker2FShoe, 0);

                    EnsureLockerValid(conn, tx, locker1FClothes, "1F", "衣柜", "1F衣柜");
                    EnsureLockerValid(conn, tx, locker1FShoe, "1F", "鞋柜", "1F鞋柜");
                    EnsureLockerValid(conn, tx, locker2FClothes, "2F", "衣柜", "2F衣柜");
                    EnsureLockerValid(conn, tx, locker2FShoe, "2F", "鞋柜", "2F鞋柜");

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"UPDATE T_Employee
SET Name = @Name,
    Pinyin = @Pinyin,
    Process = @Process,
    Locker_1F_Clothes = @L1C,
    Locker_1F_Shoe = @L1S,
    Locker_2F_Clothes = @L2C,
    Locker_2F_Shoe = @L2S
WHERE EmpNo = @EmpNo";
                        cmd.Parameters.AddWithValue("@EmpNo", empNo.Trim());
                        cmd.Parameters.AddWithValue("@Name", name.Trim());
                        cmd.Parameters.AddWithValue("@Pinyin", PinYin.GetFirstLetter(name));
                        cmd.Parameters.AddWithValue("@Process", NormalizeNull(process));
                        cmd.Parameters.AddWithValue("@L1C", NormalizeNull(locker1FClothes));
                        cmd.Parameters.AddWithValue("@L1S", NormalizeNull(locker1FShoe));
                        cmd.Parameters.AddWithValue("@L2C", NormalizeNull(locker2FClothes));
                        cmd.Parameters.AddWithValue("@L2S", NormalizeNull(locker2FShoe));
                        cmd.ExecuteNonQuery();
                    }

                    MarkLockerOccupied(conn, tx, locker1FClothes, 1);
                    MarkLockerOccupied(conn, tx, locker1FShoe, 1);
                    MarkLockerOccupied(conn, tx, locker2FClothes, 1);
                    MarkLockerOccupied(conn, tx, locker2FShoe, 1);

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"更新员工成功: {empNo}-{name}");
            CaptureLockerSnapshot("UpdateEmployee");
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
                    MarkLockerOccupied(conn, tx, info.Locker1FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker1FShoe, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FShoe, 0);

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = @"UPDATE T_Employee
SET Status = 0,
    Locker_1F_Clothes = NULL,
    Locker_1F_Shoe = NULL,
    Locker_2F_Clothes = NULL,
    Locker_2F_Shoe = NULL
WHERE EmpNo = @EmpNo";
                        cmd.Parameters.AddWithValue("@EmpNo", empNo);
                        cmd.ExecuteNonQuery();
                    }

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"员工离职并释放柜位: {info.EmpNo}-{info.Name}");
            CaptureLockerSnapshot("Resign");
        }

        public static void RestoreEmployee(string empNo)
        {
            var info = GetEmployeeLockerInfo(empNo);
            if (info == null)
            {
                throw new InvalidOperationException("未找到该员工。");
            }

            if (info.Status == 1)
            {
                return;
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = "UPDATE T_Employee SET Status = 1 WHERE EmpNo = @EmpNo";
                cmd.Parameters.AddWithValue("@EmpNo", empNo);
                cmd.ExecuteNonQuery();
            }

            WriteSystemLog("Employee", $"员工复职: {info.EmpNo}-{info.Name}");
            CaptureLockerSnapshot("Restore");
        }

        public static void DeleteEmployee(string empNo)
        {
            var info = GetEmployeeLockerInfo(empNo);
            if (info == null)
            {
                throw new InvalidOperationException("未找到该员工。");
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    MarkLockerOccupied(conn, tx, info.Locker1FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker1FShoe, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FClothes, 0);
                    MarkLockerOccupied(conn, tx, info.Locker2FShoe, 0);

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        cmd.CommandText = "DELETE FROM T_Employee WHERE EmpNo = @EmpNo";
                        cmd.Parameters.AddWithValue("@EmpNo", empNo);
                        int affected = cmd.ExecuteNonQuery();
                        if (affected <= 0)
                        {
                            throw new InvalidOperationException("删除失败：员工记录不存在或已被删除。");
                        }
                    }

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"删除员工成功: {info.EmpNo}-{info.Name}");
            CaptureLockerSnapshot("DeleteEmployee");
        }

        public static void CaptureLockerSnapshot(string source)
        {
            var summary = GetLockerSummary();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"INSERT INTO T_LockerSnapshot
(SnapshotTime, OneFClothesOccupied, OneFClothesTotal, OneFShoeOccupied, OneFShoeTotal,
 TwoFClothesOccupied, TwoFClothesTotal, TwoFShoeOccupied, TwoFShoeTotal, Source)
VALUES (CURRENT_TIMESTAMP, @OneFCO, @OneFCT, @OneFSO, @OneFST, @TwoFCO, @TwoFCT, @TwoFSO, @TwoFST, @Source)";
                cmd.Parameters.AddWithValue("@OneFCO", summary.OneFClothesOccupied);
                cmd.Parameters.AddWithValue("@OneFCT", summary.OneFClothesTotal);
                cmd.Parameters.AddWithValue("@OneFSO", summary.OneFShoeOccupied);
                cmd.Parameters.AddWithValue("@OneFST", summary.OneFShoeTotal);
                cmd.Parameters.AddWithValue("@TwoFCO", summary.TwoFClothesOccupied);
                cmd.Parameters.AddWithValue("@TwoFCT", summary.TwoFClothesTotal);
                cmd.Parameters.AddWithValue("@TwoFSO", summary.TwoFShoeOccupied);
                cmd.Parameters.AddWithValue("@TwoFST", summary.TwoFShoeTotal);
                cmd.Parameters.AddWithValue("@Source", source ?? string.Empty);
                cmd.ExecuteNonQuery();
            }
        }

        public static DataTable QueryLockerSnapshots(int limit)
        {
            var table = new DataTable();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            using (var adapter = new SQLiteDataAdapter(cmd))
            {
                conn.Open();
                cmd.CommandText = @"SELECT SnapshotTime, OneFClothesOccupied, OneFShoeOccupied,
       TwoFClothesOccupied, TwoFShoeOccupied, Source
FROM T_LockerSnapshot
ORDER BY SnapshotTime DESC
LIMIT @limit";
                cmd.Parameters.AddWithValue("@limit", limit <= 0 ? 100 : limit);
                adapter.Fill(table);
            }

            return table;
        }

        public static DataTable QuerySystemLogs(string logType, int limit)
        {
            var table = new DataTable();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            using (var adapter = new SQLiteDataAdapter(cmd))
            {
                conn.Open();
                cmd.CommandText = @"SELECT LogType AS 类型, Message AS 内容, LogTime AS 时间
FROM T_SystemLog
WHERE (@type = '' OR LogType = @type)
ORDER BY LogTime DESC
LIMIT @limit;";
                cmd.Parameters.AddWithValue("@type", logType ?? string.Empty);
                cmd.Parameters.AddWithValue("@limit", limit <= 0 ? 50 : limit);
                adapter.Fill(table);
            }

            return table;
        }

        public static List<EmployeeItemInput> GetEmployeeItems(string empNo)
        {
            var list = new List<EmployeeItemInput>();
            using (var conn = new SQLiteConnection(ConnectionString))
            using (var cmd = conn.CreateCommand())
            {
                conn.Open();
                cmd.CommandText = @"SELECT I.Category, I.SlotIndex, I.Size, I.IssueDate
FROM T_Emp_Items I
INNER JOIN T_Employee E ON E.ID = I.EmpID
WHERE E.EmpNo = @EmpNo
ORDER BY I.Category, I.SlotIndex";
                cmd.Parameters.AddWithValue("@EmpNo", empNo);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new EmployeeItemInput
                        {
                            Category = reader["Category"].ToString(),
                            SlotIndex = Convert.ToInt32(reader["SlotIndex"]),
                            Size = reader["Size"].ToString(),
                            IssueDate = reader["IssueDate"].ToString()
                        });
                    }
                }
            }

            return list;
        }

        public static void ReplaceEmployeeItems(string empNo, List<EmployeeItemInput> items)
        {
            if (string.IsNullOrWhiteSpace(empNo))
            {
                throw new ArgumentException("工号不能为空。", "empNo");
            }

            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    int empId;
                    using (var findCmd = conn.CreateCommand())
                    {
                        findCmd.Transaction = tx;
                        findCmd.CommandText = "SELECT ID FROM T_Employee WHERE EmpNo = @EmpNo";
                        findCmd.Parameters.AddWithValue("@EmpNo", empNo);
                        object id = findCmd.ExecuteScalar();
                        if (id == null || id == DBNull.Value)
                        {
                            throw new InvalidOperationException("未找到员工，无法保存劳保用品。");
                        }

                        empId = Convert.ToInt32(id);
                    }

                    using (var delCmd = conn.CreateCommand())
                    {
                        delCmd.Transaction = tx;
                        delCmd.CommandText = "DELETE FROM T_Emp_Items WHERE EmpID = @EmpID";
                        delCmd.Parameters.AddWithValue("@EmpID", empId);
                        delCmd.ExecuteNonQuery();
                    }

                    if (items != null)
                    {
                        foreach (var item in items)
                        {
                            using (var insertCmd = conn.CreateCommand())
                            {
                                insertCmd.Transaction = tx;
                                insertCmd.CommandText = @"INSERT INTO T_Emp_Items(EmpID, Category, SlotIndex, Size, IssueDate)
VALUES (@EmpID, @Category, @SlotIndex, @Size, @IssueDate)";
                                insertCmd.Parameters.AddWithValue("@EmpID", empId);
                                insertCmd.Parameters.AddWithValue("@Category", item.Category ?? string.Empty);
                                insertCmd.Parameters.AddWithValue("@SlotIndex", item.SlotIndex);
                                insertCmd.Parameters.AddWithValue("@Size", NormalizeNull(item.Size));
                                insertCmd.Parameters.AddWithValue("@IssueDate", NormalizeNull(item.IssueDate));
                                insertCmd.ExecuteNonQuery();
                            }
                        }
                    }

                    tx.Commit();
                }
            }

            WriteSystemLog("Employee", $"更新劳保用品信息: {empNo}, 数量={(items == null ? 0 : items.Count)}");
        }

        public static void WriteSystemLog(string logType, string message)
        {
            ExecuteNonQuery("INSERT INTO T_SystemLog (LogType, Message) VALUES (@type, @message)",
                new SQLiteParameter("@type", logType),
                new SQLiteParameter("@message", message));
        }

        private static void SeedDefaultLockers()
        {
            using (var conn = new SQLiteConnection(ConnectionString))
            {
                conn.Open();
                using (var tx = conn.BeginTransaction())
                {
                    SeedByType(conn, tx, "1F", "衣柜", "1F-C", 60);
                    SeedByType(conn, tx, "1F", "鞋柜", "1F-S", 60);
                    SeedByType(conn, tx, "2F", "衣柜", "2F-C", 60);
                    SeedByType(conn, tx, "2F", "鞋柜", "2F-S", 60);
                    tx.Commit();
                }
            }
        }

        private static void SeedByType(SQLiteConnection conn, SQLiteTransaction tx, string location, string type, string prefix, int count)
        {
            for (int i = 1; i <= count; i++)
            {
                string lockerId = string.Format("{0}-{1:00}", prefix, i);
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"INSERT OR IGNORE INTO T_Lockers (LockerID, Location, Type, IsOccupied)
VALUES (@LockerID, @Location, @Type, 0)";
                    cmd.Parameters.AddWithValue("@LockerID", lockerId);
                    cmd.Parameters.AddWithValue("@Location", location);
                    cmd.Parameters.AddWithValue("@Type", type);
                    cmd.ExecuteNonQuery();
                }
            }
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

    public class EmployeeEditModel
    {
        public string EmpNo { get; set; }
        public string Name { get; set; }
        public string Process { get; set; }
        public string Locker1FClothes { get; set; }
        public string Locker1FShoe { get; set; }
        public string Locker2FClothes { get; set; }
        public string Locker2FShoe { get; set; }
        public int Status { get; set; }
    }

    public class EmployeeItemInput
    {
        public string Category { get; set; }
        public int SlotIndex { get; set; }
        public string Size { get; set; }
        public string IssueDate { get; set; }
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

    public class LockerSummary
    {
        public int OneFClothesOccupied { get; set; }
        public int OneFClothesTotal { get; set; }
        public int OneFShoeOccupied { get; set; }
        public int OneFShoeTotal { get; set; }
        public int TwoFClothesOccupied { get; set; }
        public int TwoFClothesTotal { get; set; }
        public int TwoFShoeOccupied { get; set; }
        public int TwoFShoeTotal { get; set; }

        public void Set(string location, string type, int occupied, int total)
        {
            if (location == "1F" && type == "衣柜")
            {
                OneFClothesOccupied = occupied;
                OneFClothesTotal = total;
            }
            else if (location == "1F" && type == "鞋柜")
            {
                OneFShoeOccupied = occupied;
                OneFShoeTotal = total;
            }
            else if (location == "2F" && type == "衣柜")
            {
                TwoFClothesOccupied = occupied;
                TwoFClothesTotal = total;
            }
            else if (location == "2F" && type == "鞋柜")
            {
                TwoFShoeOccupied = occupied;
                TwoFShoeTotal = total;
            }
        }
    }
}
