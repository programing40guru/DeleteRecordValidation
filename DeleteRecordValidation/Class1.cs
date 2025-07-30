using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Xml.Linq;

 

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
                        decimal Marks = 0;
                        decimal Scholarship = 0;
                        decimal Parcentage = 0;

                        string Teacher = null;
                        Guid TeacherId = Guid.Empty;
                        string TeacherEntityName = null;
                        string TeacherName = null;

                        string Birthday = null;
                        string Owner = null;
                        Guid OwnerId = Guid.Empty;
                        string OwnerEntityName = null;
                        string OwnerName = null;

                        string UserName = null;
                        Guid UserId = Guid.Empty;
                        string UserEntityName = null;


                        foreach (var student in Student.Entities)
                        {
                            var auditId = student.Id;
                            var createdOn = student.GetAttributeValue<DateTime>("createdon");
                            var operation = student.GetAttributeValue<OptionSetValue>("operation")?.Value;
                            UserName = student.GetAttributeValue<EntityReference>("userid")?.Name;
                            UserId = student.GetAttributeValue<EntityReference>("userid")?.Id ?? Guid.Empty;
                            UserEntityName = student.GetAttributeValue<EntityReference>("userid")?.LogicalName;

                            var Data = student.GetAttributeValue<string>("changedata");

                            tracingService.Trace($"AuditID: {auditId}, Date: {createdOn}, Operation: {operation}, User: {UserName} :{UserId}:{UserEntityName}");
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
                                            Name = x.GetProperty("oldValue").GetString() ?? null;
                                            break;
                                        }
                                    case "crf39_subjectname":
                                        {
                                            Subject = x.GetProperty("oldValue").GetString() ?? null;
                                            break;
                                        }
                                    case "new_teacher":
                                        {
                                            Teacher = x.GetProperty("oldValue").GetString() ?? null;
                                            if (Teacher != null)
                                            {
                                                TeacherEntityName = Teacher.Split(',')[0];
                                                TeacherId = new Guid(Teacher.Split(',')[1]);
                                                
                                                Entity teacher = service.Retrieve(TeacherEntityName,TeacherId, new ColumnSet("new_name"));
                                                TeacherName = teacher.GetAttributeValue<string>("new_name");
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
                                            if (Scholarships != null)
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
                                            string Mark = x.GetProperty("oldValue").GetString() ?? null;
                                            if (Mark != null)
                                            {
                                                Marks = Convert.ToDecimal(Mark);
                                                Marks = Math.Round(Marks, 2);
                                            }
                                            break;
                                        }
                                    case "ownerid":
                                        {
                                            Owner = x.GetProperty("oldValue").GetString() ?? null;
                                            if (Owner != null)
                                            {
                                                OwnerId = new Guid(Owner.Split(',')[1]);
                                                OwnerEntityName = Owner.Split(',')[0]; 
                                                Entity Owners = service.Retrieve(OwnerEntityName,OwnerId, new ColumnSet("fullname"));
                                                OwnerName = Owners.GetAttributeValue<string>("fullname");
                                            }
                                            break;
                                        }
                                    default: {

                                            break;
                                        }

                                }

                            }
                            tracingService.Trace($"{Name},{Subject},{Marks},{Scholarship},{Parcentage},{Teacher},{Birthday}");
                            CreateEmailActivity(service, OwnerId, OwnerEntityName, OwnerName, Name, Subject, Marks, Scholarship, Parcentage, TeacherName,TeacherId, TeacherEntityName,UserName, UserId, UserEntityName, Birthday);
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

        public static void  CreateEmailActivity(IOrganizationService service, Guid OwnerId,string  OwnerEntityName, string OwnerName, string Name, string Subject,decimal Marks,decimal  Scholarship,decimal Parcentage, string TeacherName, Guid TeacherId , string TeacherEntityName,string  UserName, Guid UserId, string UserEntityName, string Birthday) {

            Entity Emails = new Entity("email");

            var FromPointer = new Entity("activityparty");
            FromPointer["partyid"] = new EntityReference(OwnerEntityName, OwnerId);

            var ToOwner = new Entity("activityparty");
            ToOwner["partyid"] = new EntityReference(UserEntityName, UserId);

 
            Emails["to"] = new EntityCollection(new List<Entity> { ToOwner });
            if (TeacherName != null)
            {
                
                Emails["regardingobjectid"] = new EntityReference(TeacherEntityName, TeacherId);
            } 
            Emails["from"] = new EntityCollection(new List<Entity> { FromPointer });

 
            Emails["ownerid"] = new EntityReference(UserEntityName, UserId);

            //Email Body Operations 
            Name = Name ?? "Name is not available";
            Birthday = "Birthday : " + Birthday;
            TeacherName = TeacherName ?? "Teacher Name is not available";
            Subject = Subject ?? "Subject is not available";

            string Body = $"Hello {TeacherName},\nI am writing this email to inform you about the student termination and record deletion.\nBelow are the student details:\n\nStudent Name: {Name}\n{Birthday}\nMarks: {Marks}\nPercentage: {Parcentage}\nSubject: {Subject}\nScholarship: {Scholarship}\nTeacher Name: {TeacherName}\n\n Thank you.\nRegards,\n{OwnerName}";
          
            Emails["subject"] = $"{Name}'s Record is Deleted";
            Emails["description"] = Body;

            Emails["directioncode"] = true;

            service.Create(Emails);

        }
    }
}