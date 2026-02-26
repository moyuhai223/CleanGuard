以下为该 README 的中文版翻译（根据原文完整翻译整理）：

---

# CleanGuard

基于 **.NET Framework 4.6.2 + WinForms + SQLite** 构建的生产部门劳保及更衣柜管理系统。
面向工厂更衣室场景设计，提供员工、更衣柜与劳保用品的全流程管理。

---

## 功能特性

### 员工管理（FrmMain）

| 操作   | 说明                        |
| ---- | ------------------------- |
| 新增员工 | 输入工号、姓名、工序；分配 1F/2F 衣柜与鞋柜 |
| 编辑员工 | 修改信息与柜号；旧柜自动释放，新柜自动占用     |
| 离职   | 释放全部柜位并标记离职（高风险操作，需确认）    |
| 恢复   | 将离职员工恢复为在职状态，并重新分配柜位      |
| 删除   | 永久删除员工并释放柜位（高风险，不可逆）      |
| 搜索   | 支持按姓名 / 拼音首字母 / 工序模糊查询    |

---

### 劳保用品管理（FrmEditor）

* 支持按类别动态增减行（无尘服、安全鞋、帆布鞋、洁净帽）
* 可录入尺码、编码（服装/帽子）、新旧状态（鞋类）、发放日期
* 支持批量粘贴
* 提供 CSV 模板下载与导入（带预校验）
* 各类别数量限制支持在线配置（持久化存储于 T_SystemConfig）

---

### 标签打印（Printer）

* 支持打印预览与直接打印
* 支持多人批量打印
* 标签内容包含：

  * 姓名 / 工号 / 工序
  * 四个柜位区块
  * 二维码（可选 QRCoder，若缺失则文本回退）
* 打印预设可持久化保存：

  * 默认打印机
  * 纸张
  * 页边距
  * 方向
  * 标签尺寸
* 批量打印支持重试 / 跳过 / 终止
* 缺失字段自动警告

---

### 数据导入（FrmImport + ImportHelper）

* 支持 CSV / XLSX 批量导入员工
* 优先使用 XLSX 模板；若缺少 SharpZipLib 自动回退为 CSV
* 导入结果统计：

  * 成功 / 失败数量
  * 失败详情预览（前 20 条）
  * 一键复制错误
* 支持导出回填模板：

  * 将失败行原样导出以便修正后重新导入
* 错误代码体系（CG-IMP-001 ~ CG-IMP-012）：

  * 精确到行列
  * 提供修复建议

---

### 更衣柜管理

| 模块                    | 说明                                    |
| --------------------- | ------------------------------------- |
| 柜位维护（FrmLockerManage） | 按楼层/类型/异常筛选，维护异常备注，批量导入柜位             |
| 柜位图表（FrmLockerChart）  | 占用率饼图 + 分区热力块（每10个一组）+ 汇总信息，支持 PNG 导出 |
| 占用趋势（FrmLockerTrend）  | 基于 T_LockerSnapshot 的四类柜位折线趋势图        |

---

### 工序字典管理（FrmProcessManage）

* 新增 / 删除 / 重命名（自动同步已分配员工）
* 支持 CSV/TXT 批量导入
* 审计视图（FrmProcessAudit）：

  * 按操作类型/日期/关键字筛选
  * 行颜色区分
  * 支持 CSV 导出

---

### 系统日志（FrmSystemLog）

* 按类型筛选（导入 / 备份 / 打印 / 员工）
* 支持 CSV 导出
* 所有关键操作自动记录日志

---

## 运维与容错机制

* **自动备份**

  * 程序退出时自动备份 CleanGuard.db 至 Backup/ 目录
  * 自动清理 7 天前的备份文件

* **全局异常处理**

  * ThreadException
  * UnhandledException
  * 关键路径 try-catch

* **数据库自动迁移**

  * CREATE TABLE IF NOT EXISTS
  * EnsureColumnExists（字段自动补齐）

---

## 技术栈

| 项目    | 说明                                     |
| ----- | -------------------------------------- |
| 框架    | .NET Framework 4.6.2, Windows Forms    |
| 数据库   | SQLite (System.Data.SQLite 1.0.119.0)  |
| Excel | NPOI 2.5.6                             |
| 压缩    | SharpZipLib 1.4.2（NPOI 依赖，缺失时 CSV 回退）  |
| 二维码   | QRCoder（可选，反射加载）                       |
| 图表    | System.Windows.Forms.DataVisualization |
| 拼音    | 内置 GB2312 区位码首字母算法（PinYin.cs）          |
| 构建目标  | x86（兼容 ARM64 Windows）                  |

---

## 数据库结构

| 表名               | 用途                          |
| ---------------- | --------------------------- |
| T_Employee       | 员工主表（工号、姓名、拼音、工序、四个柜位、状态）   |
| T_Emp_Items      | 劳保用品子表（类别、槽位、尺码、编码、新旧、发放日期） |
| T_Lockers        | 柜位主数据（柜号、楼层、类型、占用情况、备注）     |
| T_Process        | 工序字典                        |
| T_SystemLog      | 操作日志                        |
| T_SystemConfig   | 键值配置（数量限制、打印预设等）            |
| T_LockerSnapshot | 柜位占用快照（趋势图数据源）              |

默认初始化：

* 每层（1F/2F）60 个衣柜 + 60 个鞋柜
* 6 个预设工序

---

## 项目结构

```
Src/CleanGuard_App/
  Program.cs                  # 程序入口
  Forms/
    FrmMain.cs                # 主界面：员工列表、搜索、操作按钮
    FrmEditor.cs              # 员工编辑：基本信息 + 柜位 + 劳保用品
    FrmImport.cs              # 导入向导：模板下载、导入、失败预览
    FrmLockerChart.cs         # 柜位图表：饼图 + 热力图 + 汇总
    FrmLockerTrend.cs         # 柜位趋势图
    FrmLockerManage.cs        # 柜位维护：筛选、备注、批量导入
    FrmProcessManage.cs       # 工序字典：增删改 + 批量导入
    FrmProcessAudit.cs        # 工序审计：筛选、颜色区分、导出
    FrmSystemLog.cs           # 系统日志
  Utils/
    SQLiteHelper.cs           # 数据访问层
    ImportHelper.cs           # CSV/XLSX 导入导出引擎
    Printer.cs                # 标签打印引擎
    PinYin.cs                 # 拼音首字母处理
    UiTheme.cs                # UI 主题
Src/Database/
  init.sql                    # 数据库结构参考脚本
Output/                       # 构建输出目录
```

---

## 构建与运行

### 环境要求

* Windows 7 及以上（需 .NET Framework 4.6.2）
* Visual Studio 2017+ / MSBuild 15.0+

---

### 构建命令

```bash
nuget restore Src/CleanGuard_App/packages.config -PackagesDirectory packages
msbuild Src/CleanGuard_App/CleanGuard_App.csproj /t:Build /p:Configuration=Debug
```

构建产物输出至 `Output/` 目录。

---

### 运行

直接运行：

```
Output/CleanGuard_App.exe
```

首次启动将自动创建 `CleanGuard.db` 并初始化所有表及默认数据。

---

## 启动流程

```
Program.Main()
  |-- 注册全局异常处理
  |-- SQLiteHelper.InitializeDatabase()
  |     |-- 创建表（7张）
  |     |-- 自动补齐字段
  |     |-- 初始化默认配置
  |     +-- CaptureLockerSnapshot("Startup")
  |-- new FrmMain()
  |-- Application.Run(mainForm)
  +-- 关闭时自动备份数据库
```

---

## 开源协议

MIT License 
