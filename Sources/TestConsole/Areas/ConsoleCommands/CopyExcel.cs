using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Mmu.Mlh.ConsoleExtensions.Areas.Commands.Models;
using Mmu.Mlh.ExcelAccess.Areas.Services;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ConsoleCommands
{
    public class CopyExcel : IConsoleCommand
    {
        private readonly IExcelWorkbookFactory _excelWorkbookFactory;
        private readonly IExcelFileWriter _excelFileWriter;

        public string Description => "Copy Excel";

        public ConsoleKey Key => ConsoleKey.F1;

        public CopyExcel(
            IExcelWorkbookFactory excelWorkbookFactory,
            IExcelFileWriter excelFileWriter)
        {
            _excelWorkbookFactory = excelWorkbookFactory;
            _excelFileWriter = excelFileWriter;
        }

        public Task ExecuteAsync()
        {
            var fp = @"C:\Users\mlm\Desktop\TestFile.xlsx";
            var workbook = _excelWorkbookFactory.Create(fp);

            _excelFileWriter.WriteToStream(workbook);
            return Task.CompletedTask;
        }
    }
}
