using System;
using System.Text;

namespace CleanGuard_App.Utils
{
    public static class PinYin
    {
        private static readonly int[] AreaCode =
        {
            45217, 45253, 45761, 46318, 46826, 47010, 47297, 47614,
            48119, 49062, 49324, 49896, 50371, 50614, 50622, 50906,
            51387, 51446, 52218, 52698, 52980, 53689, 54481
        };

        private static readonly char[] Letters =
        {
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H',
            'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
            'R', 'S', 'T', 'W', 'X', 'Y', 'Z'
        };

        public static string GetFirstLetter(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            foreach (char ch in input.Trim())
            {
                char initial = GetCharInitial(ch);
                if (initial != '\0')
                {
                    sb.Append(initial);
                }
            }

            return sb.ToString();
        }

        private static char GetCharInitial(char ch)
        {
            if (char.IsWhiteSpace(ch))
            {
                return '\0';
            }

            if (char.IsLetterOrDigit(ch))
            {
                return char.ToUpperInvariant(ch);
            }

            var bytes = Encoding.GetEncoding("gb2312").GetBytes(ch.ToString());
            if (bytes.Length < 2)
            {
                return '#';
            }

            int code = bytes[0] * 256 + bytes[1];
            for (int i = 0; i < AreaCode.Length; i++)
            {
                int max = i == AreaCode.Length - 1 ? int.MaxValue : AreaCode[i + 1];
                if (code >= AreaCode[i] && code < max)
                {
                    return Letters[i];
                }
            }

            return '#';
        }
    }
}
