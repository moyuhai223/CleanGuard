# CleanGuard

生产部劳保与更衣柜管理系统（CG-200）初始开发骨架，基于 .NET Framework 4.6.2 + SQLite。

## 当前已实现

- WinForms 项目骨架（主界面、编辑界面）
- SQLite 自动初始化（按 V4.0 文档建表）
- 员工列表查询（姓名/拼音/工序模糊检索）
- 四柜位下拉筛选（1F/2F + 衣柜/鞋柜）
- 程序退出自动备份数据库并清理 7 天前备份

## 目录

- `Src/CleanGuard_App`：主程序代码
- `Src/Database/init.sql`：数据库初始化脚本
- `Doc`：文档（预留）

## 下一步建议

1. 接入 `System.Data.SQLite` NuGet 引用（目前仅代码层已准备）。
2. 完成 `FrmEditor` 的保存/校验/离职释放逻辑。
3. 新增 Excel 导入、二维码打印与柜位分布图模块。
