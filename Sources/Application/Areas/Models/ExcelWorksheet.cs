using System.Collections.Generic;
using System.Linq;
using Mmu.Mlh.LanguageExtensions.Areas.Invariance;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelWorksheet
    {
        public IReadOnlyCollection<ExcelCell> Cells { get; }
        public string Name { get; }

        public int MaxRow => Cells.Max(f => f.Address.Row);

        public ExcelCell this[int row, int col]
        {
            get
            {
                return Cells.SingleOrDefault(f => f.Address.Row == row && f.Address.Column == col);
            }
        }

        public ExcelWorksheet(string name, IReadOnlyCollection<ExcelCell> cells)
        {
            Guard.ObjectNotNull(() => cells);
            Name = name;
            Cells = cells;
        }
    }
}