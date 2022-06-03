using System;
using System.Collections.Generic;
using System.Text;

namespace Mmu.Mlh.ExcelAccess.TestConsole.Areas.ExcelComparison.Models
{
    public class ExcelMember
    {
        public string FirstName { get; }
        public string LastName { get; }
        public string Email { get; }
        public string Street { get; }
        public string Zip { get; }
        public string City { get; }

        public bool CompareByAddress(ExcelMember other)
        {
            return other.Email == Email
                && other.Street == Street
                && other.Zip == Zip
                && other.City == City;
        }
        
        public bool CompareByName(ExcelMember other)
        {
            return other.FirstName == FirstName && other.LastName == LastName;
        }

        public ExcelMember(
            string firstName,
            string lastName,
            string email,
            string street,
            string zip,
            string city)
        {
            FirstName = firstName;
            LastName = lastName;
            Email = email;
            Street = street;
            Zip = zip;
            City = city;
        }

    }
}
