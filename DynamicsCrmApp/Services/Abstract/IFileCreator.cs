using DynamicsCrmApp.Domain.Derived;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsCrmApp.Services.Abstract
{
    interface IFileCreator
    {
        int ExportToCsvFile(List<Person> people, string filename);
    }
}
