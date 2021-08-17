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
        private readonly string password = "Mobkoi@199";//Mobkoi@100!
        private readonly string accountName = "account";

        public DynamicsCrmService()
        {}

        //public void CreateCrmAccount()
        //{
        //    try
        //    {
        //        var clientCredentials = new ClientCredentials();
        //        clientCredentials.UserName.UserName = username;
        //        clientCredentials.UserName.Password = password;
        //        // For Dynamics 365 Customer Engagement, Set Security Protocol as TLS12
        //        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

        //        var svc = new CrmServiceClient($"Url={new Uri(uri)}; Username={username}; Password={password}; AuthType=Office365");

        //        var account = new Entity("account");
        //        account["name"] = "account";

        //        // Create an account record  
        //        Dictionary<string, CrmDataTypeWrapper> inData = new Dictionary<string, CrmDataTypeWrapper>();
        //        inData.Add("name", new CrmDataTypeWrapper("Jimmy Jones", CrmFieldType.String));
        //        inData.Add("telephone1", new CrmDataTypeWrapper("447963954299", CrmFieldType.String));
        //        inData.Add("address1_city", new CrmDataTypeWrapper("2 Demo Street", CrmFieldType.String));
        //        inData.Add("address1_postalcode", new CrmDataTypeWrapper("NW2 3DL", CrmFieldType.String));
        //        inData.Add("emailaddress1", new CrmDataTypeWrapper("jimmysjones@email.com", CrmFieldType.String));
        //        inData.Add("websiteurl", new CrmDataTypeWrapper("https://www.mobkoi.com/", CrmFieldType.String));

        //        //var newId = svc.Create(account);

        //        // Verify that you are connected  
        //        if (svc != null && svc.IsReady)
        //        {
        //            svc.CreateNewRecord("account", inData, applyToSolution: " ", enabledDuplicateDetection: false);
        //        }


        //        //var newId = svc.Create(account);
        //        //var _svc = svc.Retrieve("account", newId, new ColumnSet(true));
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(" - " + e.Message);
        //    }
        //}

        public Guid CreateCrmAccount(string accountName)
        {
            var account = new Entity("account");
            account["name"] = accountName;
            var newId = _service.Create(account);
            return newId;
        }

        public void CreateCrmAccount2(string accountName)
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

        //public void ConnectToDynamicsCrm()
        //{
        //    try
        //    {
        //        if (_service != null)
        //        {
        //            Guid userid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;

        //            if (userid != Guid.Empty)
        //            {
        //                Console.WriteLine("Connection Successful!...");
        //                //var accountId = CreateCrmAccount(accountName);//organizationService.Create(myEntity)
        //            }
        //        }
        //        else
        //        {
        //            Console.WriteLine("Failed to Established Connection!!!");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error while connecting to CRM - " + ex.Message);
        //    }
        //}

        public void CreateCrmAccount()
        {
            try
            {
                var clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = username;
                clientCredentials.UserName.Password = password;
                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                var svc = new CrmServiceClient($"Url={new Uri(uri)}; Username={username}; Password={password}; AuthType=Office365");

                //var account = new Entity(accountName);//"account"
                //account["name"] = "account";

                // Create an account record  
                Dictionary<string, CrmDataTypeWrapper> inData = new Dictionary<string, CrmDataTypeWrapper>();
                inData.Add("name", new CrmDataTypeWrapper("Jimmy Jones", CrmFieldType.String));
                inData.Add("telephone1", new CrmDataTypeWrapper("447963954299", CrmFieldType.String));
                inData.Add("address1_city", new CrmDataTypeWrapper("4 Demo Street", CrmFieldType.String));
                inData.Add("address1_postalcode", new CrmDataTypeWrapper("NW1 3DL", CrmFieldType.String));
                inData.Add("emailaddress1", new CrmDataTypeWrapper("jimmyjones@email.com", CrmFieldType.String));
                inData.Add("websiteurl", new CrmDataTypeWrapper("https://www.mobkoi.com/", CrmFieldType.String));

                //var newId = svc.Create(account);

                //Guid guid = new Guid();

                // Verify that you are connected  
                if (svc != null && svc.IsReady)
                {
                    svc.CreateNewRecord(accountName, inData, applyToSolution: "", enabledDuplicateDetection: false);
                    //guid = svc.CreateNewRecord("account", inData, applyToSolution: "", enabledDuplicateDetection: false);
                }


                //var newId = svc.Create(account);
                //var _svc = svc.Retrieve("account", newId, new ColumnSet(true));

                // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
                //organizationService = new OrganizationServiceProxy(new Uri(uri), null, clientCredentials, null);

                IOrganizationService organizationService = (IOrganizationService)svc.OrganizationWebProxyClient != null ?
                    (IOrganizationService)svc.OrganizationWebProxyClient :
                    (IOrganizationService)svc.OrganizationServiceProxy;

                if (organizationService != null)
                {
                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Successful!...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }

                //var retrieve = organizationService.Retrieve("account", guid, new ColumnSet(true));

                //foreach (var a in retrieve)
                //{
                //    Console.WriteLine(a);
                //}

                //var retrieve1 = GetEntityCollection(account, organizationService);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM - " + ex.Message);
            }
            //return organizationService;
        }

        /**
         * Account Name : "name"
            mainphone: "telephone1"
            city : "address1_city"
            zip / postal code : "address1_postalcode"
            email address : "emailaddress1"
            website : "websiteurl"
         * ***/

        //private static EntityCollection GetEntityCollection(Entity entity, IOrganizationService service)
        //{
        //    ConditionExpression condition1 = new ConditionExpression();
        //    condition1.AttributeName = "name";//firstname
        //    condition1.Operator = ConditionOperator.Equal;
        //    condition1.Values.Add("Jimmy Jones");

        //    ConditionExpression condition2 = new ConditionExpression();
        //    condition2.AttributeName = "telephone1";
        //    condition2.Operator = ConditionOperator.BeginsWith;
        //    condition2.Values.Add("44");

        //    //ConditionExpression condition3 = new ConditionExpression();
        //    //condition2.AttributeName = "address1_postalcode";
        //    //condition2.Operator = ConditionOperator.Contains;
        //    //condition2.Values.Add(" ");

        //    //ConditionExpression condition4 = new ConditionExpression();
        //    //condition2.AttributeName = "jobtitle";
        //    //condition2.Operator = ConditionOperator.NotNull;

        //    //ConditionExpression condition5 = new ConditionExpression();
        //    //condition2.AttributeName = "telephone1";
        //    ////condition2.Operator = ConditionOperator.Equal;

        //    //ConditionExpression condition6 = new ConditionExpression();
        //    //condition2.AttributeName = "parentcustomerid";
        //    ////condition2.Operator = ConditionOperator.Equal;

        //    FilterExpression filter1 = new FilterExpression();
        //    filter1.Conditions.Add(condition1);
        //    filter1.Conditions.Add(condition2);
        //    //filter1.Conditions.Add(condition3);
        //    //filter1.Conditions.Add(condition4);
        //    //filter1.Conditions.Add(condition5);
        //    //filter1.Conditions.Add(condition6);

        //    QueryExpression query = new QueryExpression(entity.LogicalName);
        //    query.ColumnSet.AddColumns("name", "telephone1");// "emailaddress1", "jobtitle", "telephone1", "parentcustomerid"
        //    query.Criteria.AddFilter(filter1);

        //    var _orgService = service;

        //    EntityCollection result1 = _orgService.RetrieveMultiple(query);

        //    foreach (var a in result1.Entities)
        //    {
        //        Console.WriteLine("Name: " + a.Attributes["name"] + " " + "Telephone: " + a.Attributes["telephone1"] + " " + "AccountId: " + a.Attributes["accountid"]);
        //        // + " " + a.Attributes["emailaddress1"] + " " + a.Attributes["jobtitle"]+ " " + a.Attributes["telephone1"] + " " + a.Attributes["parentcustomerid"]
        //    }

        //    return result1;
        //}
    }
}
