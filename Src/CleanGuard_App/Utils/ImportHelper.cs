using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CleanGuard_App.Utils
{
    public static class ImportHelper
    {
        public static readonly string[] TemplateHeaders =
        {
            "工号", "姓名", "工序", "1F衣柜", "1F鞋柜", "2F衣柜", "2F鞋柜", "无尘服1尺码", "鞋码"
        };

        public static void ExportTemplateCsv(string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", TemplateHeaders));
            sb.AppendLine("P001,张三,热切,1F-C-01,1F-S-01,2F-C-01,2F-S-01,L,42");
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        public static ImportResult ImportFromCsv(string filePath)
        {
            var result = new ImportResult();
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

            for (int i = 1; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                string[] cells = SplitCsvLine(line);
                if (cells.Length < 7)
                {
                    result.FailedCount++;
                    result.Errors.Add($"第 {i + 1} 行列数不足，至少需要 7 列。内容: {line}");
                    continue;
                }

                try
                {
                    SQLiteHelper.AddEmployee(
                        cells[0],
                        cells[1],
                        cells[2],
                        cells[3],
                        cells[4],
                        cells[5],
                        cells[6]);
                    result.SuccessCount++;
                }
                catch (Exception ex)
                {
                    result.FailedCount++;
                    string empNo = cells.Length > 0 ? cells[0] : "(空)";
                    result.Errors.Add($"第 {i + 1} 行导入失败（工号: {empNo}）：{ex.Message}");
                }
            }

            SQLiteHelper.WriteSystemLog("Import", $"CSV导入完成：成功 {result.SuccessCount}，失败 {result.FailedCount}");
            return result;
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
        public List<string> Errors { get; } = new List<string>();
    }
}
