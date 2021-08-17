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
        {}

        /// <summary>
        /// CreateCrmAccount()
        /// Method for Test Class
        /// </summary>
        /// <param name="accountName"></param>
        /// <returns></returns>
        public Guid CreateCrmAccount(string accountName)
        {
            var account = new Entity("account");
            account["name"] = accountName;
            var newId = _service.Create(account);
            return newId;
        }

        /// <summary>
        /// CreateCrmAccountAndRetrieveEntity()
        /// Method for Test Class
        /// </summary>
        /// <param name="accountName"></param>
        public void CreateCrmAccountAndRetrieveEntity(string accountName)
        {
            //create the account
            var account = new Entity("account");
            account["name"] = accountName;
            var newId = _service.Create(account);

            //get the account number
            account = _service.Retrieve("account", newId, new Microsoft.Xrm.Sdk.Query.ColumnSet(new string[] { "name", "accountid", "accountnumber" }));
            var accountNumber = account["accountnumber"].ToString();

            //create the task
            var task = new Entity("task");
            task["subject"] = "Finish account set up for " + accountName + " - " + accountNumber;
            task["regardingobjectid"] = new EntityReference("account", newId);
            _service.Create(task);
        }

        /// <summary>
        /// CreateCrmAccount()
        /// Connects to Dynamics CRM using credentials
        /// Attempts to Create an Account
        /// </summary>
        public void CreateCrmAccount()
        {
            Console.WriteLine("Connect to Dynamics CRM Process Started....\n");
            try
            {
                var clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = username;
                clientCredentials.UserName.Password = password;
                // For Dynamics 365 Customer Engagement V9.X, Set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                Console.WriteLine("Attempting to Connect to Dynamics CRM....\n");
                var svc = new CrmServiceClient($"Url={new Uri(uri)}; Username={username}; Password={password}; AuthType=Office365");

                // Create an account record  
                Dictionary<string, CrmDataTypeWrapper> inData = new Dictionary<string, CrmDataTypeWrapper>();
                inData.Add("name", new CrmDataTypeWrapper("Jimmy Jones", CrmFieldType.String));
                inData.Add("telephone1", new CrmDataTypeWrapper("447963954299", CrmFieldType.String));
                inData.Add("address1_city", new CrmDataTypeWrapper("4 Demo Street", CrmFieldType.String));
                inData.Add("address1_postalcode", new CrmDataTypeWrapper("NW1 3DL", CrmFieldType.String));
                inData.Add("emailaddress1", new CrmDataTypeWrapper("jimmyjones@email.com", CrmFieldType.String));
                inData.Add("websiteurl", new CrmDataTypeWrapper("https://www.mobkoi.com/", CrmFieldType.String));

                // Verify that you are connected  
                if (svc != null && svc.IsReady)
                {
                    svc.CreateNewRecord(accountName, inData, applyToSolution: "", enabledDuplicateDetection: false);
                    Console.WriteLine("Attempting to Create a New Account....\n");
                }

                IOrganizationService organizationService = (IOrganizationService)svc.OrganizationWebProxyClient != null ?
                    (IOrganizationService)svc.OrganizationWebProxyClient :
                    (IOrganizationService)svc.OrganizationServiceProxy;

                if (organizationService != null)
                {
                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Successful!...\n");
                        Console.WriteLine("New CRM Account Was Created Successfully...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM - " + ex.Message);
            }
        }
    }
}
