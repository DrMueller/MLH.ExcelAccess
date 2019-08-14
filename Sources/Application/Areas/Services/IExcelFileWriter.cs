using System.IO;
using Mmu.Mlh.ExcelAccess.Areas.Models;

namespace Mmu.Mlh.ExcelAccess.Areas.Services
{
    public interface IExcelFileWriter
    {
        void WriteToStream(ExcelWorkbook workbook);
    }
}