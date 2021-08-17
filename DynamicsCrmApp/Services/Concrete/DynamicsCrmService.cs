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
        private readonly string accountName = "account";

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
                        account["name"] = "account";

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

                var account = new Entity(accountName);//"account"
                account["name"] = "account";

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
            condition1.AttributeName = "name";//firstname
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add("Jimmy Jones");

            ConditionExpression condition2 = new ConditionExpression();
            condition2.AttributeName = "telephone1";
            condition2.Operator = ConditionOperator.BeginsWith;
            condition2.Values.Add("44");

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);
            filter1.Conditions.Add(condition2);

            QueryExpression query = new QueryExpression(entity.LogicalName);
            query.ColumnSet.AddColumns("name", "telephone1");
            query.Criteria.AddFilter(filter1);

            EntityCollection result1 = _service.RetrieveMultiple(query);
            var count = 0;

            foreach (var a in result1.Entities)
            {
                people.Add(new Person(firstName: $"{a.Attributes["name"]}", lastName: "", email: "", jobTitle: "", telephoneNumber: long.Parse($"{a.Attributes["telephone1"]}"), companyName: ""));
                count++;
            }

            Console.WriteLine($"Obtained {count} Accounts from Dynamics CRM...\n");

            return people;
        }
    }
}
