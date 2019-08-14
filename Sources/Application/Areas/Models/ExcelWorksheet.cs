using System.Collections.Generic;
using Mmu.Mlh.LanguageExtensions.Areas.Invariance;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelWorksheet
    {
        public IReadOnlyCollection<ExcelCell> Cells { get; }
        public string Name { get; }

        public ExcelWorksheet(string name, IReadOnlyCollection<ExcelCell> cells)
        {
            Guard.ObjectNotNull(() => cells);
            Name = name;
            Cells = cells;
        }
    }
}