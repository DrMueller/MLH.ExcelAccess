using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mmu.Mlh.ConsoleExtensions.Areas.Commands.Models;
using Mmu.Mlh.ExcelAccess.Areas.Services;
using Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Services;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ConsoleCommands
{
    public class CheckExcel : IConsoleCommand
    {
        private readonly ICheckExcelService _service;

        public CheckExcel(ICheckExcelService service)
        {
            _service = service;
        }

        public Task ExecuteAsync()
        {
            _service.Check();

            return Task.CompletedTask;
        }

        public string Description { get; } = "Check excel";
        public ConsoleKey Key { get; } = ConsoleKey.F2;
    }
}
