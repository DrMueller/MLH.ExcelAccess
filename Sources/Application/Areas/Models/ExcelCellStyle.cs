using System.Drawing;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelCellStyle
    {
        public Color BackgroundColor { get; }

        public ExcelCellStyle(Color backgroundColor)
        {
            BackgroundColor = backgroundColor;
        }
    }
}