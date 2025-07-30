using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Text.Json;

namespace DeleteRecordValidation
{
    public class DeleteRecordValidation : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.MessageName.Equals("Delete", StringComparison.OrdinalIgnoreCase))
            {
                // The InputParameters collection contains all the data passed in the message request. 
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    // Obtain the target entity from the input parameters.

                    //   Entity entity = (Entity)context.InputParameters["Target"];
                    EntityReference entity = (EntityReference)context.InputParameters["Target"];

                    // Obtain the IOrganizationService instance which you will need for web service calls.
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    try
                    {
                        tracingService.Trace("I am Plugin");
                       
                        Guid RecordId = entity.Id;
                        /*
                        string EntityName = entity.LogicalName;

                        Guid userId = context.InitiatingUserId;

                        var FullEntity = service.Retrieve(EntityName, RecordId, new ColumnSet("ownerid"));


                        EntityReference OwnerRef = FullEntity.GetAttributeValue<EntityReference>("ownerid"); //lookup
                        Guid OwnerId = OwnerRef.Id;

                        tracingService.Trace($"User ID :{userId} and Owener ID : {OwnerId}");


                        tracingService.Trace("Delete event triggered.");
                        if (OwnerId != userId)
                        {
                            throw new InvalidPluginExecutionException($"You are unable to Delete Account.Beacause You are not Owner Thats Record.");
                        }
                        else
                        {
                            tracingService.Trace("Delate Record Succefully");

                        }

                        */
                        EntityCollection Student = GetAuditDelatedData(service, RecordId);

                        string Name = null;
                        string Subject = null; 
                        string Marks = null;
                        decimal Scholarship= 0 ;
                        decimal Parcentage =0;
                        string Teacher = null;
                        string Birthday = null; 

                        foreach (var student in Student.Entities)
                        {
                            var auditId = student.Id;
                            var createdOn = student.GetAttributeValue<DateTime>("createdon");
                            var operation = student.GetAttributeValue<OptionSetValue>("operation")?.Value;
                            var user = student.GetAttributeValue<EntityReference>("userid")?.Name;
                            var Data = student.GetAttributeValue<string>("changedata");

                            tracingService.Trace($"AuditID: {auditId}, Date: {createdOn}, Operation: {operation}, User: {user}");
                            tracingService.Trace(Data);

                            JsonDocument doc = JsonDocument.Parse(Data);

                            var changedAttributes = doc.RootElement.GetProperty("changedAttributes");

                            foreach (var x in changedAttributes.EnumerateArray())
                            {
                                string logicalNames = x.GetProperty("logicalName").GetString();
                               // tracingService.Trace("" + logicalNames);

                                switch (logicalNames)
                                {
                                    case "new_name":
                                        {
                                            Name = x.GetProperty("oldValue").GetString()?? null;
                                            break;
                                        }
                                    case "crf39_subjectname":
                                        {
                                            Subject = x.GetProperty("oldValue").GetString()?? null;
                                            break;
                                        }
                                    case "new_teacher":
                                        {
                                            Teacher = x.GetProperty("oldValue").GetString()?? null;
                                            if (Teacher != null)
                                            {
                                                Entity teacher = service.Retrieve(Teacher.Split(',')[0], new Guid(Teacher.Split(',')[1]), new ColumnSet("new_name"));
                                                Teacher = teacher.GetAttributeValue<string>("new_name");
                                            }
                                            break;
                                        }
                                    case "new_birthday":
                                        {
                                            Birthday = x.GetProperty("oldValue").GetString() ?? null;
                                            if (Birthday != null)
                                            {
                                                Birthday = Birthday.Split(' ')[0];
                                            } 
                                            break;
                                        }
                                    case "new_maxscholarshipamount":
                                        {
                                            string Scholarships = x.GetProperty("oldValue").GetString();
                                            if (Scholarships !=null)
                                            {
                                                Scholarship = Convert.ToDecimal(Scholarships);
                                                Scholarship = Math.Round(Scholarship, 2); 
                                            }
                                            break;
                                        }
                                    case "new_parcentage":
                                        {
                                            string Parcentages = x.GetProperty("oldValue").GetString();
                                            if (Parcentages != null)
                                            {
                                                Parcentage = Convert.ToDecimal(Parcentages);
                                               Parcentage = Math.Round(Parcentage, 2);   
                                            }

                                            break;
                                        }
                                    case "new_marks":
                                        {
                                            Marks = x.GetProperty("oldValue").GetString() ?? null;
                                            break;
                                        }
                                    default: {
                                            
                                            break; 
                                    }

                                }
                                 
                            }
                            tracingService.Trace($"{Name}{Subject}{Marks}{Scholarship}{Parcentage}{Teacher}{Birthday}");
 
                        }

                        tracingService.Trace("try CaseFinished");
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                        throw;
                    }

                }
            }
        }

        public static EntityCollection GetAuditDelatedData(IOrganizationService service, Guid recordId) //post-operations
        {
            string fetchXml = $@"<fetch top=""1"">
                <entity name=""audit"">
                <attribute name=""auditid"" />
                <attribute name=""changedata"" />
                <attribute name=""createdon"" />
                <attribute name=""objectid"" />
                <attribute name=""operation"" />
                <attribute name=""transactionid"" />
                <attribute name=""userid"" />
                <filter type=""and"">
                    <condition attribute=""transactionid"" operator=""ne"" value=""00000000-0000-0000-0000-000000000000"" />
                    <condition attribute=""operation"" operator=""eq"" value=""3"" />
                    <condition attribute='objectid' operator='eq' value='{recordId}' />
                </filter>
                <order attribute=""createdon"" descending=""true"" />
                </entity>
            </fetch>";
            
            return service.RetrieveMultiple(new FetchExpression(fetchXml));
        }
    }
}