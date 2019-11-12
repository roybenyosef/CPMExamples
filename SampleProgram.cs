using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerApps.Samples
{
   public partial class SampleProgram
    {
        [STAThread] // Added to support UX
        static void Main(string[] args)
        {
            CrmServiceClient service = null;

            try
            {
                service = SampleHelpers.Connect("Connect");
                if (service.IsReady)
                {
                    #region Sample Code
                    ////////////////////////////////////
                    #region Set up
                    SetUpSample(service);
                    #endregion Set up
                    #region Demonstrate
                    // Obtain information about the logged on user from the web service.
                    Guid userid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
                    SystemUser systemUser = (SystemUser)service.Retrieve("systemuser", userid,
                        new ColumnSet(new string[] { "firstname", "lastname" }));
                    Console.WriteLine("Logged on user is {0} {1}.", systemUser.FirstName, systemUser.LastName);

                    // Retrieve the version of Microsoft Dynamics CRM.
                    RetrieveVersionRequest versionRequest = new RetrieveVersionRequest();
                    RetrieveVersionResponse versionResponse =
                        (RetrieveVersionResponse)service.Execute(versionRequest);
                    Console.WriteLine("Microsoft Dynamics CRM version {0}.", versionResponse.Version);

                    bool fetchEverything = false;

                    if (fetchEverything)
                    {
                        QueryExpression queGetAllIncidents = new QueryExpression("incident");
                        queGetAllIncidents.ColumnSet = new ColumnSet(true);
                        EntityCollection allIncidents = service.RetrieveMultiple(queGetAllIncidents);
                        WriteToJson(allIncidents.Entities[0], @"entity0.json");

                        QueryExpression queGetAllContacts = new QueryExpression("contact");
                        queGetAllContacts.ColumnSet = new ColumnSet(true);
                        EntityCollection allContacts = service.RetrieveMultiple(queGetAllContacts);
                        WriteToJson(allContacts, @"contacts.json");

                        QueryExpression queGetAllAccounts = new QueryExpression("account");
                        queGetAllAccounts.ColumnSet = new ColumnSet(true);
                        //queGetAllAccounts.ColumnSet = new ColumnSet(new String[]{ "AccountId", "Name" });
                        EntityCollection allAccounts = service.RetrieveMultiple(queGetAllAccounts);
                        WriteToJson(allAccounts, @"accounts.json");
                    }

                    
                    //TODO - fetch by name
                    var incident = new Incident
                    {
                        Title = "Testing From Roy",
                        //PrimaryContactId = new EntityReference("contact", new Guid("{6a9334ad-bd04-ea11-a811-000d3a4a1025}")),
                        CustomerId = new EntityReference("account", new Guid("{852179cc-e2a6-e911-a97f-000d3a2cba5f}"))
                    };
                    
                    var incidentId = service.Create(incident);
                    ColumnSet cols = new ColumnSet(
                        new String[] { "title", "new_tguvot_text" });

                    Incident retrievedIncident = (Incident)service.Retrieve("incident", incidentId, cols);
                    Console.Write("retrieved, ");

                    var incidentAttributes = retrievedIncident.Attributes;
                    incidentAttributes["new_tguvot_text"] = 5;

                    service.Update(retrievedIncident);
                    Console.WriteLine("and updated.");



                }
                #endregion Demonstrate
                #endregion Sample Code
                else
                {
                    const string UNABLE_TO_LOGIN_ERROR = "Unable to Login to Common Data Service";
                    if (service.LastCrmError.Equals(UNABLE_TO_LOGIN_ERROR))
                    {
                        Console.WriteLine("Check the connection string values in cds/App.config.");
                        throw new Exception(service.LastCrmError);
                    }
                    else
                    {
                        throw service.LastCrmException;
                    }
                }
            }
            catch (Exception ex)
            {
                SampleHelpers.HandleException(ex);
            }

            finally
            {
                if (service != null)
                    service.Dispose();

                Console.WriteLine("Press <Enter> to exit.");
                Console.ReadLine();
            }
        }

        private static void WriteToJson(Object entity, String filename)
        {
            using (var outputJson = new StreamWriter(filename))
            {
                outputJson.Write(JsonConvert.SerializeObject(entity, Formatting.Indented));
            }
        }
    }
}
