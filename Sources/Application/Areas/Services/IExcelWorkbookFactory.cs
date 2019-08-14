using System.IO;
using Mmu.Mlh.ExcelAccess.Areas.Models;

namespace Mmu.Mlh.ExcelAccess.Areas.Services
{
    public interface IExcelWorkbookFactory
    {
        ExcelWorkbook Create(Stream stream);

        ExcelWorkbook Create(string filePath);
    }
}