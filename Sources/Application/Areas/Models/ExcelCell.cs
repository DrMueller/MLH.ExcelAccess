using Mmu.Mlh.LanguageExtensions.Areas.Invariance;

namespace Mmu.Mlh.ExcelAccess.Areas.Models
{
    public class ExcelCell
    {
        public ExcelCellAddress Address { get; }
        public ExcelCellStyle Style { get; }
        public object Value { get; }

        public ExcelCell(object value, ExcelCellAddress address, ExcelCellStyle style)
        {
            Guard.ObjectNotNull(() => address);
            Value = value;
            Address = address;
            Style = style;
        }
    }
}