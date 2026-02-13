using System.Linq;
using System.Text;

namespace CleanGuard_App.Utils
{
    public static class PinYin
    {
        public static string GetFirstLetter(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            foreach (char ch in input.Trim())
            {
                if (char.IsLetterOrDigit(ch))
                {
                    sb.Append(char.ToUpperInvariant(ch));
                }
                else if (ch >= 0x4e00 && ch <= 0x9fa5)
                {
                    sb.Append('#');
                }
            }

            return new string(sb.ToString().Where(c => c != ' ').ToArray());
        }
    }
}
