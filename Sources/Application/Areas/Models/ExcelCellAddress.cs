using Mmu.Mlh.LanguageExtensions.Areas.Invariance;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelCellAddress
    {
        public int Column { get; }
        public int Row { get; }

        public ExcelCellAddress(int row, int column)
        {
            Guard.That(() => row >= 1, "Row must start with Index 1.");
            Guard.That(() => column >= 1, "Column must start with Index 1.");
            Row = row;
            Column = column;
        }
    }
}