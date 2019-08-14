using System.IO.Abstractions;
using Mmu.Mlh.ExcelAccess.Areas.Services;
using Mmu.Mlh.ExcelAccess.Areas.Services.Implementation;
using StructureMap;

namespace Mmu.Mlh.ExcelAccess.Infrastructure.DependencyInjection
{
    public class ExcelAccessRegistry : Registry
    {
        public ExcelAccessRegistry()
        {
            For<IFileSystem>().Use<FileSystem>().Singleton();

            For<IExcelFileWriter>().Use<ExcelFileWriter>().Singleton();
            For<IExcelWorkbookFactory>().Use<ExcelWorkbookFactory>().Singleton();
        }
    }
}