using DynamicsCrmApp.Domain.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsCrmApp.Domain.Derived
{
    public class Person : MyEntity
    {
        public Person(string firstName, string lastName, string email, string jobTitle, string telephoneNumber, string companyName) : base(firstName, lastName, email, telephoneNumber)
        {
            FirstName = firstName;
            LastName = lastName;
            Email = email;
            JobTitle = jobTitle;
            TelephoneNumber = telephoneNumber;
            CompanyName = companyName;
        }

        public string JobTitle { get; set; }
        public string CompanyName { get; set; }
    }
}
