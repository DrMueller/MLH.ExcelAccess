using System.Collections.Generic;
using Mmu.Mlh.LanguageExtensions.Areas.Invariance;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelWorkbook
    {
        public IReadOnlyCollection<ExcelWorksheet> Worksheets { get; }

        public ExcelWorkbook(IReadOnlyCollection<ExcelWorksheet> worksheets)
        {
            Guard.ObjectNotNull(() => worksheets);
            Worksheets = worksheets;
        }
    }
}