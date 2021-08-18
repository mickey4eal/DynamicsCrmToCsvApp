using DynamicsCrmApp.Domain.Derived;
using DynamicsCrmApp.Services.Abstract;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsCrmApp.Services.Concrete
{
    public class FileCreator : IFileCreator
    {
        public int ExportToCsvFile(List<Person> people, string filename)
        {
            var successCode = 0;//0 is Ok
            Directory.CreateDirectory("C:\\CrmExport\\");
            var path = "C:\\CrmExport\\" + filename + ".csv";

            var headerLine = "First Name,Last Name,Email,Job Title,Telephone Number,Company Name";

            try
            {
                var csv = new StringBuilder();
                csv.AppendLine(headerLine);

                foreach (var person in people)
                {
                    var newLine = string.Format("{0},{1},{2},{3},{4},{5}",
                        person.FirstName, person.LastName, person.Email, person.JobTitle, person.TelephoneNumber, person.CompanyName);
                    csv.AppendLine(newLine);
                }

                if (path != "")
                {
                    File.WriteAllText(path, csv.ToString());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                successCode = 1;
            }

            return successCode;
        }
    }
}
