using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mmu.Mlh.ExcelAccess.Areas.Models;
using Mmu.Mlh.ExcelAccess.Areas.Services;
using Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Models;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Services.Servants.Implementation
{
    public class ExcellMemberFactory : IExcellMemberFactory
    {
        private readonly IExcelWorkbookFactory _wbFactory;

        public ExcellMemberFactory(IExcelWorkbookFactory wbFactory)
        {
            _wbFactory = wbFactory;
        }

        public IReadOnlyCollection<ExcelMember> CreateClubMembers()
        {
            var ws = _wbFactory.Create(@"C:\Users\mlm\Desktop\Club.xlsx").Worksheets.Single();

            var result = new List<ExcelMember>();

            for (var iRow = 2; iRow <= ws.MaxRow; iRow++)
            {
                var firstName = AsString(ws[iRow, 4]);
                var lastName = AsString(ws[iRow, 3]);
                var street = AsString(ws[iRow, 8]);
                var zip = AsString(ws[iRow, 10]);
                var city = AsString(ws[iRow, 11]);
                var email = AsString(ws[iRow, 18]);

                result.Add(new ExcelMember(firstName, lastName, email, street, zip, city));
            }

            return result;
        }

        private string AsString(ExcelCell cell)
        {
            return cell?.Value.ToString().Trim() ?? string.Empty;
        }

        public IReadOnlyCollection<ExcelMember> CreateAktivMembers()
        {
            var ws = _wbFactory.Create(@"C:\Users\mlm\Desktop\AktivListe.xlsx").Worksheets.Single();

            var result = new List<ExcelMember>();

            for (var iRow = 2; iRow <= ws.MaxRow; iRow++)
            {
                var firstName = AsString(ws[iRow, 3]);
                var lastName = AsString(ws[iRow, 4]);
                var street = AsString(ws[iRow, 6]);
                var zip = AsString(ws[iRow, 7]);
                var city = AsString(ws[iRow, 8]);
                var email = AsString(ws[iRow, 5]);

                result.Add(new ExcelMember(firstName, lastName, email, street, zip, city));
            }

            return result;
        }
    }
}
