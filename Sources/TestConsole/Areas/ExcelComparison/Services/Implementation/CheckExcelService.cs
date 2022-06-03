using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mmu.Mlh.ConsoleExtensions.Areas.ConsoleOutput.Services;
using Mmu.Mlh.ExcelAccess.Areas.Services;
using Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Models;
using Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Services.Servants;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Services.Implementation
{
    public class CheckExcelService : ICheckExcelService
    {
        private readonly IExcellMemberFactory _memberFactory;
        private readonly IConsoleWriter _consoleWriter;

        public CheckExcelService(
            IExcellMemberFactory memberFactory,
            IConsoleWriter consoleWriter)
        {
            _memberFactory = memberFactory;
            _consoleWriter = consoleWriter;
        }

        public void Check()
        {
            var clubMembers = _memberFactory.CreateClubMembers();
            var aktivMembers = _memberFactory.CreateAktivMembers();

            CheckNames(clubMembers, aktivMembers, "Club Members");
            CheckNames(aktivMembers, clubMembers, "Aktiv Members");
            CheckAddresses(clubMembers, aktivMembers);
        }

        private void CheckAddresses(
            IReadOnlyCollection<ExcelMember> clubMembers,
            IReadOnlyCollection<ExcelMember> aktivMembers)
        {
            _consoleWriter.WriteLine("Not matching addresses:");
            foreach (var aktivMember in aktivMembers)
            {
                var clubMember = clubMembers.SingleOrDefault(f => f.CompareByName(aktivMember));
                if (clubMember == null)
                {
                    continue;
                }

                if (!aktivMember.CompareByAddress(clubMember))
                {
                    _consoleWriter.WriteLine($"{aktivMember.FirstName} {aktivMember.LastName}", foregroundColor: System.ConsoleColor.Red);
                }
            }
        }

        private void CheckNames(
            IReadOnlyCollection<ExcelMember> listToCheck,
            IReadOnlyCollection<ExcelMember> checkList,
            string name)
        {
            _consoleWriter.WriteLine($"Missing names in {name}:");

            foreach (var aktivMember in checkList)
            {
                if (!listToCheck.Any(f => f.CompareByName(aktivMember)))
                {
                    _consoleWriter.WriteLine($"{aktivMember.FirstName} {aktivMember.LastName}", foregroundColor: System.ConsoleColor.Red);
                }
            }

       
        }
    }
}