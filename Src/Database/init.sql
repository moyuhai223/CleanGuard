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

-- 默认柜位初始化（每层 60 衣柜 + 60 鞋柜）
WITH RECURSIVE seq(x) AS (
    SELECT 1
    UNION ALL
    SELECT x + 1 FROM seq WHERE x < 60
)
INSERT OR IGNORE INTO T_Lockers (LockerID, Location, Type, IsOccupied)
SELECT printf('1F-C-%02d', x), '1F', '衣柜', 0 FROM seq;

WITH RECURSIVE seq(x) AS (
    SELECT 1
    UNION ALL
    SELECT x + 1 FROM seq WHERE x < 60
)
INSERT OR IGNORE INTO T_Lockers (LockerID, Location, Type, IsOccupied)
SELECT printf('1F-S-%02d', x), '1F', '鞋柜', 0 FROM seq;

WITH RECURSIVE seq(x) AS (
    SELECT 1
    UNION ALL
    SELECT x + 1 FROM seq WHERE x < 60
)
INSERT OR IGNORE INTO T_Lockers (LockerID, Location, Type, IsOccupied)
SELECT printf('2F-C-%02d', x), '2F', '衣柜', 0 FROM seq;

WITH RECURSIVE seq(x) AS (
    SELECT 1
    UNION ALL
    SELECT x + 1 FROM seq WHERE x < 60
)
INSERT OR IGNORE INTO T_Lockers (LockerID, Location, Type, IsOccupied)
SELECT printf('2F-S-%02d', x), '2F', '鞋柜', 0 FROM seq;
