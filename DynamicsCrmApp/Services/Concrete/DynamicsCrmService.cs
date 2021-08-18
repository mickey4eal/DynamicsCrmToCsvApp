using DynamicsCrmApp.Domain.Derived;
using DynamicsCrmApp.Services.Concrete;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace DynamicsCrmApp.Services.Abstract
{
    public class DynamicsCrmService : IDynamicsCrmService
    {
        private readonly IOrganizationService _service;

        public DynamicsCrmService(IOrganizationService service)
        {
            _service = service;
        }

        private readonly string uri = "https://org4be57bdb.crm11.dynamics.com/main.aspx";
        private readonly string username = "testuser1@mobkoiltd.onmicrosoft.com";
        private readonly string password = "Mobkoi@199";
        private readonly string accountName = "contact";

        public DynamicsCrmService()
        {
            _service = SetOrganisationService(uri, username, password);
        }

        public void ConnectToDynamicsCrmAndGetEntityCollection()
        {
            try
            {
                if (_service != null)
                {
                    Guid userid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Successful!...");

                        var account = new Entity(accountName);//"account"
                        account["name"] = "contact";

                        Console.WriteLine("Attempting to Retrieve from Dynamics CRM....\n");

                        var retrieve = GetEntityCollection(account);

                        Console.WriteLine("Exporting to CSV File...\n");

                        new FileCreator().ExportToCsvFile(retrieve, "CsvExport"+DateTime.Now.ToFileTime());

                        Console.WriteLine("Excel spreadsheet / CSV file Generated - See Folder c:\\CrmExport");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM - " + ex.Message);
            }
        }

        public IOrganizationService SetOrganisationService(string uri, string username, string password)
        {
            IOrganizationService organizationService = null;

            try
            {
                var clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = username;
                clientCredentials.UserName.Password = password;

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var svc = new CrmServiceClient($"Url={new Uri(uri)}; Username={username}; Password={password}; AuthType=Office365");

                organizationService = (IOrganizationService)svc.OrganizationWebProxyClient != null ?
                    (IOrganizationService)svc.OrganizationWebProxyClient :
                    (IOrganizationService)svc.OrganizationServiceProxy;
            }
            catch (Exception e)
            {
                Console.WriteLine(" - " + e.Message);
            }
            return organizationService;
        }

        private List<Person> GetEntityCollection(Entity entity)
        {
            var people = new List<Person>();

            ConditionExpression condition1 = new ConditionExpression();

            FilterExpression filter1 = new FilterExpression();

            QueryExpression query = new QueryExpression(entity.LogicalName);
            query.ColumnSet.AddColumns("firstname", "lastname", "emailaddress1", "jobtitle", "telephone1", "parentcustomerid");

            EntityCollection result1 = _service.RetrieveMultiple(query);
            var count = 0;

            foreach (var a in result1.Entities)
            {
                var entityReference = (EntityReference) a.Attributes["parentcustomerid"];
                var compName = entityReference.Name;
                people.Add(new Person(firstName: $"{a.Attributes["firstname"]}", lastName: $"{a.Attributes["lastname"]}", email: $"{a.Attributes["emailaddress1"]}", jobTitle: $"{a.Attributes["jobtitle"]}", telephoneNumber: $"{a.Attributes["telephone1"]}", companyName: $"{compName}"));
                count++;
            }

            Console.WriteLine($"Obtained {count} Accounts from Dynamics CRM...\n");

            return people;
        }
    }
}
