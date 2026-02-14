using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CleanGuard_App.Utils
{
    public static class ImportHelper
    {
        public static readonly string[] TemplateHeaders =
        {
            "工号", "姓名", "工序", "1F衣柜", "1F鞋柜", "2F衣柜", "2F鞋柜", "无尘服1尺码", "鞋码"
        };
        public static readonly string[] LockerTemplateHeaders =
        {
            "柜号", "楼层", "类型", "异常备注"
        };

        public static void ExportTemplate(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                ExportTemplateXlsx(filePath);
                return;
            }

            ExportTemplateCsv(filePath);
        }

        public static ImportResult ImportFromFile(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                return ImportFromXlsx(filePath);
            }

            return ImportFromCsv(filePath);
        }

        public static void ExportLockerTemplate(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                ExportLockerTemplateXlsx(filePath);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", LockerTemplateHeaders));
            sb.AppendLine("1F-C-01,1F,衣柜,");
            sb.AppendLine("1F-S-01,1F,鞋柜,");
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static ProcessImportResult ImportLockersFromFile(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                return ImportLockersFromXlsx(filePath);
            }

            return ImportLockersFromCsv(filePath);
        }

        public static ProcessImportResult ImportLockersFromCsv(string filePath)
        {
            var result = new ProcessImportResult();
            if (!File.Exists(filePath))
            {
                result.Errors.Add("导入文件不存在。");
                return result;
            }

            var rows = new List<LockerImportRow>();
            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            for (int i = 1; i < lines.Length; i++)
            {
                string line = (lines[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                TryParseLockerRow(SplitCsvLine(line), i + 1, result, rows);
            }

            FinalizeLockerImport(filePath, result, rows);
            return result;
        }

        public static ProcessImportResult ImportLockersFromXlsx(string filePath)
        {
            var result = new ProcessImportResult();
            if (!File.Exists(filePath))
            {
                result.Errors.Add("导入文件不存在。");
                return result;
            }

            var rows = new List<LockerImportRow>();
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.NumberOfSheets > 0 ? workbook.GetSheetAt(0) : null;
                if (sheet == null)
                {
                    result.Errors.Add("xlsx 文件无可用工作表。");
                    return result;
                }

                for (int r = 1; r <= sheet.LastRowNum; r++)
                {
                    IRow row = sheet.GetRow(r);
                    if (row == null)
                    {
                        continue;
                    }

                    string[] cells = ReadRowCells(row);
                    if (IsEmptyRow(cells))
                    {
                        continue;
                    }

                    TryParseLockerRow(cells, r + 1, result, rows);
                }
            }

            FinalizeLockerImport(filePath, result, rows);
            return result;
        }

        public static void ExportLockerTemplateXlsx(string filePath)
        {
            var workbook = new XSSFWorkbook();
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                ISheet sheet = workbook.CreateSheet("柜位导入模板");
                WriteCellsToRow(sheet.CreateRow(0), LockerTemplateHeaders);
                WriteCellsToRow(sheet.CreateRow(1), new[] { "1F-C-01", "1F", "衣柜", string.Empty });
                WriteCellsToRow(sheet.CreateRow(2), new[] { "1F-S-01", "1F", "鞋柜", string.Empty });
                workbook.Write(fs);
            }
        }

        public static void ExportTemplateCsv(string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", TemplateHeaders));
            sb.AppendLine("P001,张三,热切,1F-C-01,1F-S-01,2F-C-01,2F-S-01,L,42");
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static ImportResult ImportFromCsv(string filePath)
        {
            var result = CreateResult(filePath);
            if (!File.Exists(filePath))
            {
                result.Errors.Add("导入文件不存在。");
                return result;
            }

            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            if (lines.Length <= 1)
            {
                result.Errors.Add("文件内容为空或仅包含表头。");
                return result;
            }

            ValidateHeader(SplitCsvLine(lines[0]), result);
            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                ImportRow(SplitCsvLine(line), i + 1, result);
            }

            SQLiteHelper.WriteSystemLog("Import", string.Format("CSV导入完成：文件={0}，成功 {1}，失败 {2}", Path.GetFileName(filePath), result.SuccessCount, result.FailedCount));
            return result;
        }

        public static ImportResult ImportFromXlsx(string filePath)
        {
            var result = CreateResult(filePath);
            if (!File.Exists(filePath))
            {
                result.Errors.Add("导入文件不存在。");
                return result;
            }

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.NumberOfSheets > 0 ? workbook.GetSheetAt(0) : null;
                if (sheet == null)
                {
                    result.Errors.Add("xlsx 文件无可用工作表。");
                    return result;
                }

                IRow headerRow = sheet.GetRow(0);
                ValidateHeader(ReadRowCells(headerRow), result);

                for (int r = 1; r <= sheet.LastRowNum; r++)
                {
                    IRow row = sheet.GetRow(r);
                    if (row == null)
                    {
                        continue;
                    }

                    string[] cells = ReadRowCells(row);
                    if (IsEmptyRow(cells))
                    {
                        continue;
                    }

                    ImportRow(cells, r + 1, result);
                }
            }

            SQLiteHelper.WriteSystemLog("Import", string.Format("XLSX导入完成：文件={0}，成功 {1}，失败 {2}", Path.GetFileName(filePath), result.SuccessCount, result.FailedCount));
            return result;
        }

        public static void ExportTemplateXlsx(string filePath)
        {
            var workbook = new XSSFWorkbook();
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                ISheet sheet = workbook.CreateSheet("导入模板");
                WriteCellsToRow(sheet.CreateRow(0), TemplateHeaders);
                WriteCellsToRow(sheet.CreateRow(1), new[] { "P001", "张三", "热切", "1F-C-01", "1F-S-01", "2F-C-01", "2F-S-01", "L", "42" });

                workbook.Write(fs);
            }
        }

        private static ImportResult CreateResult(string filePath)
        {
            return new ImportResult
            {
                SourceFile = filePath,
                ImportTime = DateTime.Now
            };
        }

        private static void WriteCellsToRow(IRow row, string[] values)
        {
            if (row == null || values == null)
            {
                return;
            }

            for (int i = 0; i < values.Length; i++)
            {
                row.CreateCell(i).SetCellValue(values[i] ?? string.Empty);
            }
        }

        private static string[] ReadRowCells(IRow row)
        {
            if (row == null)
            {
                return new string[0];
            }

            int lastCellNum = row.LastCellNum;
            if (lastCellNum <= 0)
            {
                return new string[0];
            }

            var cells = new string[lastCellNum];
            for (int i = 0; i < lastCellNum; i++)
            {
                ICell cell = row.GetCell(i);
                cells[i] = cell == null ? string.Empty : cell.ToString();
            }

            return cells;
        }

        private static bool IsEmptyRow(string[] cells)
        {
            if (cells == null || cells.Length == 0)
            {
                return true;
            }

            for (int i = 0; i < cells.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(cells[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static void ValidateHeader(string[] headerCells, ImportResult result)
        {
            if (headerCells == null)
            {
                result.Errors.Add("未检测到表头，系统将继续尝试按模板列顺序导入。");
                return;
            }

            for (int i = 0; i < TemplateHeaders.Length; i++)
            {
                string actual = i < headerCells.Length ? (headerCells[i] ?? string.Empty).Trim() : string.Empty;
                if (!string.Equals(actual, TemplateHeaders[i], StringComparison.OrdinalIgnoreCase))
                {
                    result.Errors.Add(string.Format("表头第 {0} 列建议为“{1}”，当前为“{2}”。系统将继续按模板顺序读取。", i + 1, TemplateHeaders[i], actual));
                }
            }
        }

        private static void ImportRow(string[] cells, int rowNumber, ImportResult result)
        {
            if (cells == null || cells.Length < 7)
            {
                result.FailedCount++;
                result.Errors.Add(string.Format("第 {0} 行列数不足，至少需要 7 列。", rowNumber));
                return;
            }

            try
            {
                SQLiteHelper.AddEmployee(
                    SafeCell(cells, 0),
                    SafeCell(cells, 1),
                    SafeCell(cells, 2),
                    SafeCell(cells, 3),
                    SafeCell(cells, 4),
                    SafeCell(cells, 5),
                    SafeCell(cells, 6));
                result.SuccessCount++;
            }
            catch (Exception ex)
            {
                result.FailedCount++;
                result.Errors.Add(string.Format("第 {0} 行导入失败（工号: {1}）：{2}", rowNumber, SafeCell(cells, 0), ex.Message));
            }
        }

        private static void TryParseLockerRow(string[] cells, int rowNumber, ProcessImportResult result, List<LockerImportRow> rows)
        {
            int offset = 0;
            string secondCell = SafeCell(cells, 1);
            if (secondCell == "1F" || secondCell == "2F")
            {
                offset = 0;
            }
            else if (SafeCell(cells, 2) == "1F" || SafeCell(cells, 2) == "2F")
            {
                // 兼容旧模板：第1列为原始编码
                offset = 1;
            }

            string lockerId = SafeCell(cells, offset + 0);
            string location = SafeCell(cells, offset + 1);
            string type = SafeCell(cells, offset + 2);
            string remark = SafeCell(cells, offset + 3);

            if (string.IsNullOrWhiteSpace(lockerId) || string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(type))
            {
                result.FailedCount++;
                result.Errors.Add(string.Format("第 {0} 行导入失败：柜号/楼层/类型不能为空。", rowNumber));
                return;
            }

            if (location != "1F" && location != "2F")
            {
                result.FailedCount++;
                result.Errors.Add(string.Format("第 {0} 行导入失败：楼层仅支持 1F 或 2F。", rowNumber));
                return;
            }

            if (type != "衣柜" && type != "鞋柜")
            {
                result.FailedCount++;
                result.Errors.Add(string.Format("第 {0} 行导入失败：类型仅支持 衣柜 或 鞋柜。", rowNumber));
                return;
            }

            rows.Add(new LockerImportRow
            {
                LockerID = lockerId,
                Location = location,
                Type = type,
                Remark = remark
            });
            result.SuccessCount++;
        }

        private static void FinalizeLockerImport(string filePath, ProcessImportResult result, List<LockerImportRow> rows)
        {
            if (rows.Count == 0)
            {
                return;
            }

            try
            {
                SQLiteHelper.ImportLockers(rows);
                SQLiteHelper.WriteSystemLog("Import", string.Format("柜位导入完成：文件={0}，成功 {1}，失败 {2}", Path.GetFileName(filePath), result.SuccessCount, result.FailedCount));
            }
            catch (Exception ex)
            {
                result.Errors.Add("保存柜位数据失败：" + ex.Message);
                result.FailedCount += rows.Count;
                result.SuccessCount = 0;
            }
        }

        private static string SafeCell(string[] cells, int index)
        {
            if (cells == null || index < 0 || index >= cells.Length)
            {
                return string.Empty;
            }

            return (cells[index] ?? string.Empty).Trim();
        }

        private static string[] SplitCsvLine(string line)
        {
            var values = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    values.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            values.Add(sb.ToString().Trim());
            return values.ToArray();
        }
    }

    public class ImportResult
    {
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public string SourceFile { get; set; }
        public DateTime ImportTime { get; set; }
        public List<string> Errors { get; } = new List<string>();

        public void ExportErrors(string filePath)
        {
            string ext = (Path.GetExtension(filePath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xlsx")
            {
                ExportErrorsXlsx(filePath);
                return;
            }

            ExportErrorsCsv(filePath);
        }

        public void ExportErrorsCsv(string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("序号,错误信息");
            for (int i = 0; i < Errors.Count; i++)
            {
                sb.Append(i + 1);
                sb.Append(",\"");
                sb.Append(Escape(Errors[i]));
                sb.AppendLine("\"");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public void ExportErrorsXlsx(string filePath)
        {
            var workbook = new XSSFWorkbook();
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                ISheet sheet = workbook.CreateSheet("失败明细");
                var header = sheet.CreateRow(0);
                header.CreateCell(0).SetCellValue("序号");
                header.CreateCell(1).SetCellValue("错误信息");

                for (int i = 0; i < Errors.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(i + 1);
                    row.CreateCell(1).SetCellValue(Errors[i] ?? string.Empty);
                }

                workbook.Write(fs);
            }
        }

        private static string Escape(string text)
        {
            return (text ?? string.Empty).Replace("\"", "\"\"");
        }
    }
}
