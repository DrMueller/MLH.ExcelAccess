using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Models;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Services.Servants
{
    public interface IAktivListeExcelMemberFactory
    {
        Task<IReadOnlyCollection<ExcelMember>> CreateAllAsync();
    }
}
