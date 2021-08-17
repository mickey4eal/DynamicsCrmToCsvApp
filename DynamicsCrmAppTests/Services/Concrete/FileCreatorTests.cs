using Microsoft.VisualStudio.TestTools.UnitTesting;
using DynamicsCrmApp.Services.Concrete;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using DynamicsCrmApp.Domain.Derived;

namespace DynamicsCrmApp.Services.Concrete.Tests
{

    [TestClass()]
    public class FileCreatorTests
    {
        [Fact]
        public List<Person> GetTestListOfPersons()
        {
            return new List<Person>
            {
                new Person("John", "Mayer", "johnmayer@email.com", "admin",07957246351, "DummyCompany"),
                new Person("Emily", "Blunt", "emilyblunt@email.com", "admin",07957246351, "DummyCompany"),
                new Person("Austin", "Rivers", "austinrivers@email.com", "admin",07957246351, "DummyCompany"),
                new Person("Jason", "Kidd", "jasonkidd@email.com", "admin",07957246351, "DummyCompany"),
                new Person("Juliet", "June", "julietjune@email.com", "admin",07957246351, "DummyCompany"),
                new Person("Omari", "Dudd", "omaridudd@email.com", "admin",07957246351, "DummyCompany"),
            };
        }

        public string GetFileName()
        {
            return $"CrmTestExport" + DateTime.Now.ToFileTime();
        }

        [TestMethod()]
        public void ExportToCsvFileTest()
        {
            var testPersons = GetTestListOfPersons();
            var testFileName = GetFileName();

            var result = new FileCreator().ExportToCsvFile(testPersons, testFileName);

            Assert.AreEqual(0, result);
        }
    }
}