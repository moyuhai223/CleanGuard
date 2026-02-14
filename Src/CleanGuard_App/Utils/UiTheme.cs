using System.Drawing;
using System.Windows.Forms;

namespace CleanGuard_App.Utils
{
    public static class UiTheme
    {
        public static readonly Color PrimaryBlue = ColorTranslator.FromHtml("#004a80");
        public static readonly Color PrimaryBlueHover = ColorTranslator.FromHtml("#005fa3");
        public static readonly Color PrimaryBlueDown = ColorTranslator.FromHtml("#003355");
        public static readonly Color WarningOrange = ColorTranslator.FromHtml("#e67e22");
        public static readonly Color FormBackground = Color.WhiteSmoke;
        public static readonly Color PanelBackground = Color.White;
        public static readonly Color TextDark = ColorTranslator.FromHtml("#333333");
        public static readonly Font DefaultFont = new Font("Segoe UI", 9.75f, FontStyle.Regular);

        public static void ApplyFormStyle(Form form)
        {
            form.Font = DefaultFont;
            form.BackColor = FormBackground;
            form.ForeColor = TextDark;
        }

        public static void StylePrimaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = PrimaryBlueHover;
            button.FlatAppearance.MouseDownBackColor = PrimaryBlueDown;
            button.BackColor = PrimaryBlue;
            button.ForeColor = Color.White;
            button.Cursor = Cursors.Hand;
        }

        public static void StyleWarningButton(Button button)
        {
            StylePrimaryButton(button);
            button.BackColor = WarningOrange;
        }

        public static void StyleDataGrid(DataGridView grid)
        {
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.RowHeadersVisible = false;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersHeight = 35;

            grid.ColumnHeadersDefaultCellStyle.BackColor = PrimaryBlue;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);

            grid.RowsDefaultCellStyle.BackColor = Color.White;
            grid.RowsDefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#d6eaf8");
            grid.RowsDefaultCellStyle.SelectionForeColor = Color.Black;
            grid.RowsDefaultCellStyle.Padding = new Padding(5, 2, 0, 2);
            grid.AlternatingRowsDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#f2f6f8");
        }
    }
}
