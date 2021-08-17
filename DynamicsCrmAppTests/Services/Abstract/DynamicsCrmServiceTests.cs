using Microsoft.VisualStudio.TestTools.UnitTesting;
using DynamicsCrmApp.Services.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Moq;
using Microsoft.Xrm.Sdk;

namespace DynamicsCrmApp.Services.Abstract.Tests
{
    [TestClass()]
    public class DynamicsCrmServiceTests
    {
        [TestMethod()]
        public void ConnectToDynamicsCrmTest()
        {
            //ARRANGE - set up everything our test needs  

            //first - set up a mock service to act like the CRM organization service  
            var serviceMock = new Mock<IOrganizationService>();

            //next - set a name for our fake account record to create  
            string accountName = "Demo Company";

            //next - create a guid that we want our mock service Create method to return when called  
            Guid idToReturn = Guid.NewGuid();

            //next - create an entity object that will allow us to capture the entity record that is passed to the Create method  
            Entity actualEntity = new Entity();

            //finally - tell our mock service what to do when the Create method is called  
            serviceMock.Setup(t =>
            t.Create(It.IsAny<Entity>())) //when Create is called with any entity as an invocation parameter  
            .Returns(idToReturn) //return the idToReturn guid  
            .Callback<Entity>(s => actualEntity = s); //store the Create method invocation parameter for inspection later  

            //ACT - do the thing(s) we want to test  

            //call the CreateCrmAccount method like usual, but supply the mock service as an invocation parameter  
            Guid actualGuid = new DynamicsCrmService(serviceMock.Object).CreateCrmAccount(accountName);

            //ASSERT - verify the results are correct  

            //verify the entity created inside the CreateCrmAccount method has the name we supplied   
            Assert.AreEqual(accountName, actualEntity["name"]);

            //verify the guid returned by the CreateCrmAccount is the same guid the Create method returns  
            Assert.AreEqual(idToReturn, actualGuid);
        }

        [TestMethod()]
        public void CreateCrmAccount2()
        {
            //ARRANGE - set up everything our test needs  

            //first - set up a mock service to act like the CRM organization service  
            var serviceMock = new Mock<IOrganizationService>();
            IOrganizationService service = serviceMock.Object;

            //next - set a name and account number for our fake account record to create  
            string accountName = "Demo Company";
            string accountNumber = "LPA1234";

            //next - create a guid that we want our mock service Create method to return when called  
            Guid idToReturn = Guid.NewGuid();

            //next - create an object that will allow us to capture the account object that is passed to the Create method  
            Entity createdAccount = new Entity();

            //next - create an entity object that will allow us to capture the task object that is passed to the Create method  
            Entity createdTask = new Entity();

            //next - create an mock account record to pass back to the Retrieve method  
            Entity mockReturnedAccount = new Entity("account");
            mockReturnedAccount["name"] = accountName;
            mockReturnedAccount["accountnumber"] = accountNumber;
            mockReturnedAccount["accountid"] = idToReturn;
            mockReturnedAccount.Id = idToReturn;

            //finally - tell our mock service what to do when the CRM service methods are called  

            //handle the account creation  
            serviceMock.Setup(t =>
            t.Create(It.Is<Entity>(e => e.LogicalName.ToUpper() == "account".ToUpper()))) //only match an entity with a logical name of "account"  
            .Returns(idToReturn) //return the idToReturn guid  
            .Callback<Entity>(s => createdAccount = s); //store the Create method invocation parameter for inspection later  

            //handle the task creation  
            serviceMock.Setup(t =>
            t.Create(It.Is<Entity>(e => e.LogicalName.ToUpper() == "task".ToUpper()))) //only match an entity with a logical name of "task"  
            .Returns(Guid.NewGuid()) //can return any guid here  
            .Callback<Entity>(s => createdTask = s); //store the Create method invocation parameter for inspection later  

            //handle the retrieve account operation  
            serviceMock.Setup(t =>
            t.Retrieve(
            It.Is<string>(e => e.ToUpper() == "account".ToUpper()),
            It.Is<Guid>(e => e == idToReturn),
            It.IsAny<Microsoft.Xrm.Sdk.Query.ColumnSet>())
            ) //here we match on logical name of account and the correct id  
            .Returns(mockReturnedAccount);

            //ACT - do the thing(s) we want to test  

            //call the CreateCrmAccount2 method like usual, but supply the mock service as an invocation parameter  
            new DynamicsCrmService(service).CreateCrmAccountAndRetrieveEntity(accountName);

            //ASSERT - verify the results are correct  

            //verify the entity created inside the CreateCrmAccount method has the name we supplied   
            Assert.AreEqual(accountName, createdAccount["name"]);

            //verify task regardingobjectid is the same as the id we returned upon account creation  
            Assert.AreEqual(idToReturn, ((Microsoft.Xrm.Sdk.EntityReference)createdTask["regardingobjectid"]).Id);
            Assert.AreEqual("Finish account set up for " + accountName + " - " + accountNumber, (string)createdTask["subject"]);
        }
    }
}