CleanGuard
基于 .NET Framework 4.6.2 + WinForms + SQLite 构建的生产部门劳动保护和储物柜管理系统。专为工厂更衣室场景设计，提供员工、储物柜和劳动保护用品的端到端管理。

特征
员工管理（FrmMain）
行动	描述
添加员工	输入员工编号、姓名、工号；分配1楼/2楼的服装鞋柜
编辑员工	修改信息和储物柜；旧储物柜自动释放，新储物柜自动占用
辞职	释放所有储物柜并标记为已辞职（高风险，需要确认）
恢复	恢复离职员工的在职状态，重新分配储物柜
删除	永久删除并释放储物柜（高风险，不可逆）
搜索	按名称/拼音首字母/流程进行模糊搜索
劳动保护项目（FrmEditor）
按类别（洁净服、安全鞋、帆布鞋、洁净帽）动态添加/删除行
请输入尺码、代码（西装/帽子）、新旧状态（鞋子）、发行日期
批量粘贴、下载和导入 CSV 模板，并带有预检查验证功能
类别限制可在线配置（持久化在 T_SystemConfig 中）
标签打印（打印机）
打印预览和直接打印模式，批量多人打印
标签内容：姓名/ID/流程 + 四个储物柜模块 + 二维码（二维码可选，文本作为备用）
打印预设持久化（默认打印机、纸张、边距、方向、标签尺寸）
批量打印，具备重试/跳过/中止错误处理和缺失字段警告功能
数据导入（FrmImport + ImportHelper）
CSV/XLSX批量导入员工
首选 XLSX 模板；如果 SharpZipLib 缺失，则自动回退到 CSV 格式。
导入结果：成功/失败计数、失败详情预览（前 20 条）、一键复制错误信息
导出回填模板：失败的行将按原样导出，以便进行更正和重新导入。
错误代码系统（CG-IMP-001 ~ CG-IMP-012）：精确到行和列，并提供修复建议
储物柜管理
模块	描述
储物柜维护（FrmLockerManage）	按楼层/类型/异常情况筛选，保留异常备注，批量导入储物柜
储物柜图表（FrmLockerChart）	入住率饼图 + 面积热力图（每组 10 个区块）+ 汇总，PNG 导出
入住率趋势（FrmLockerTrend）	基于 T_LockerSnapshot 的四种储物柜类型的折线图
进程字典（FrmProcessManage）
添加、删除、重命名（自动同步已分配员工）
CSV/TXT批量导入
审计视图（FrmProcessAudit）：按操作类型/日期/关键字筛选、行颜色编码、导出为 CSV 文件
系统日志（FrmSystemLog）
按类型（导入/备份/打印/员工）筛选并导出为 CSV 文件
所有关键操作均自动记录
运行和容错
自动备份：退出时将 CleanGuard.db 备份到 Backup/ 目录，自动清理超过 7 天的备份
全局异常处理：ThreadException + UnhandledException + 键路径 try-catch
数据库自动迁移：CREATE TABLE IF NOT EXISTS + EnsureColumnExists
技术栈
物品	描述
框架	.NET Framework 4.6.2，Windows Forms
数据库	SQLite（System.Data.SQLite 1.0.119.0）
Excel	NPOI 2.5.6
压缩	SharpZipLib 1.4.2（依赖 NPOI，缺失时回退到 CSV）
二维码	QR编码器（可选，反射加载）
图表	System.Windows.Forms.DataVisualization
拼音	内置 GB2312 区号首字母 (PinYin.cs)
构建目标	x86（兼容 ARM64 Windows）
数据库
桌子	目的
T_员工	员工主数据（ID、姓名、拼音、工号、四个储物柜、状态）
T_Emp_Items	人工项目子表（类别、槽位、尺寸、代码、新/旧、发放日期）
T_Lockers	储物柜主数据（储物柜 ID、楼层、类型、占用情况、备注）
T_Process	流程字典
T_SystemLog	操作日志
T_SystemConfig	键值配置（项目限制、打印预设等）
T_LockerSnapshot	储物柜占用情况快照（趋势图数据源）
默认初始化：每层（1F/2F）60 个服装柜 + 60 个鞋柜，6 个预设流程。

项目结构
Src/CleanGuard_App/
  Program.cs                  # Entry point
  Forms/
    FrmMain.cs                # Main form: employee list, search, action buttons
    FrmEditor.cs              # Employee editor: basic info + lockers + items tabs
    FrmImport.cs              # Import wizard: template download, import, failure preview
    FrmLockerChart.cs         # Locker chart: pie + heatmap + summary
    FrmLockerTrend.cs         # Locker trend: line chart
    FrmLockerManage.cs        # Locker maintenance: filter, remark, bulk import
    FrmProcessManage.cs       # Process dictionary: CRUD + bulk import
    FrmProcessAudit.cs        # Process audit: filter, color coding, export
    FrmSystemLog.cs           # System log
  Utils/
    SQLiteHelper.cs           # Data access layer
    ImportHelper.cs           # CSV/XLSX import/export engine
    Printer.cs                # Label printing engine
    PinYin.cs                 # Chinese pinyin first-letter
    UiTheme.cs                # UI theme
Src/Database/
  init.sql                    # Reference schema script
Output/                       # Build output
构建和运行
要求
Windows 7+（.NET Framework 4.6.2）
Visual Studio 2017+ / MSBuild 15.0+
建造
nuget restore Src/CleanGuard_App/packages.config -PackagesDirectory packages
msbuild Src/CleanGuard_App/CleanGuard_App.csproj /t:Build /p:Configuration=Debug
构建输出将保存到Output/该目录。

跑步
直接运行Output/CleanGuard_App.exe。首次启动时，它会自动创建CleanGuard.db并初始化所有表，使用默认数据。

启动流程
Program.Main()
  |-- Register global exception handlers (ThreadException + UnhandledException)
  |-- SQLiteHelper.InitializeDatabase()
  |     |-- CREATE TABLE IF NOT EXISTS (7 tables)
  |     |-- EnsureColumnExists (forward compat)
  |     |-- SeedDefaultConfig
  |     +-- CaptureLockerSnapshot("Startup")
  |-- new FrmMain() -> InitializeLayout() -> LoadEmployeeData()
  |-- Application.Run(mainForm)
  +-- FormClosing -> BackupDatabase()
