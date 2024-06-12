
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace MazikRACImplementation
{
    public class Case : IPlugin
    {
        IOrganizationService service;
        ITracingService tracingService;
        IPluginExecutionContext context;
        Entity entity;
        public void Execute(IServiceProvider serviceProvider)
        {
            context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                entity = (Entity)context.InputParameters["Target"];
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                service = serviceFactory.CreateOrganizationService(context.UserId);
                tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                try
                {
                    switch (context.Stage)
                    {
                        case 10: // pre-validation
                            SetCaseOwnerWhenCaseProcessingLocationIsEmpty();
                            break;

                        case 20: // pre-operation

                            if (context.MessageName.ToLower() == "update")
                            {
                                if (entity.Contains("mzkah_serviceprovider") && entity.Attributes["mzkah_serviceprovider"] != null && entity.Contains("mzkahrac_geomatchedcase") == false)
                                {
                                    entity["mzkahrac_geomatchedcase"] = false;
                                }
                            }

                            SetTriageOutcomeOnManualCaseProcessingLocationAssignment();
                            CreateTimeLinePostWhenCaseAcceptedOrDeclined();

                            break;
                        case 40: // post-operation

                            SetOwnerOnCasePatientFromDefaultTeam();

                            ShareUnSharePatientRecordWithDefaultTeam();

                            ShareCaseWithReferral();

                            var waitTimeHelper = new CaseWaitTimeHelper(service, context, tracingService, entity);
                            waitTimeHelper.WaitTimePLConsultToSurgeonConsultNew();
                            waitTimeHelper.SetWaitTimeFromDateOfSpineSurgeonConsult();

                            CreateDocumentLocationOnCreateCase();
                            break;
                    }

                }
                catch (Exception ex)
                {
                    if (entity != null)
                        tracingService.Trace("entity id: " + entity.Id.ToString());
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        private void CreateDocumentLocationOnCreateCase()
        {
            if (context.MessageName.ToLower() == "create" && entity.Contains("customerid") && entity.Attributes["customerid"] != null && entity.Contains("mzk_referral") && entity.Attributes["mzk_referral"] != null)
            {
                tracingService.Trace("-1");

                EntityReference patient = entity.GetAttributeValue<EntityReference>("customerid");
                EntityReference referral = entity.GetAttributeValue<EntityReference>("mzk_referral");

                // Patient document location
                QueryExpression qeDocumentLocation = new QueryExpression() { };
                qeDocumentLocation.EntityName = "sharepointdocumentlocation";
                qeDocumentLocation.ColumnSet = new ColumnSet("relativeurl");
                qeDocumentLocation.Criteria.AddCondition("relativeurl", ConditionOperator.NotNull);
                qeDocumentLocation.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, patient.Id.ToString());

                // Patient record
                LinkEntity lePatient = qeDocumentLocation.AddLink("contact", "regardingobjectid", "contactid", JoinOperator.Inner);
                lePatient.EntityAlias = "patient";
                lePatient.Columns = new ColumnSet("fullname");


                // It should be on the "contact" library
                LinkEntity leParentDocumentLocation = qeDocumentLocation.AddLink("sharepointdocumentlocation", "parentsiteorlocation", "sharepointdocumentlocationid", JoinOperator.Inner);
                leParentDocumentLocation.EntityAlias = "parentDocumentLocation";
                leParentDocumentLocation.LinkCriteria.AddCondition("relativeurl", ConditionOperator.Equal, "contact");

                // "opportunity" folder's document location, optional
                LinkEntity leOpportunity = qeDocumentLocation.AddLink("sharepointdocumentlocation", "sharepointdocumentlocationid", "parentsiteorlocation", JoinOperator.LeftOuter);
                leOpportunity.EntityAlias = "opportunity";
                leOpportunity.Columns = new ColumnSet("sharepointdocumentlocationid");
                leOpportunity.LinkCriteria.AddCondition("relativeurl", ConditionOperator.Equal, "opportunity");
                leOpportunity.LinkCriteria.AddCondition("regardingobjectid", ConditionOperator.Null);

                // Referral folder's document location, optional
                LinkEntity leReferral = leOpportunity.AddLink("sharepointdocumentlocation", "sharepointdocumentlocationid", "parentsiteorlocation", JoinOperator.LeftOuter);
                leReferral.EntityAlias = "referral";
                leReferral.Columns = new ColumnSet("sharepointdocumentlocationid", "relativeurl");
                leReferral.LinkCriteria.AddCondition("relativeurl", ConditionOperator.NotNull);
                leReferral.LinkCriteria.AddCondition("regardingobjectid", ConditionOperator.Equal, referral.Id.ToString());

                // "incident" folder's document location, optional
                LinkEntity leIncident = leReferral.AddLink("sharepointdocumentlocation", "sharepointdocumentlocationid", "parentsiteorlocation", JoinOperator.LeftOuter);
                leIncident.EntityAlias = "incident";
                leIncident.Columns = new ColumnSet("sharepointdocumentlocationid");
                leIncident.LinkCriteria.AddCondition("relativeurl", ConditionOperator.Equal, "incident");
                leIncident.LinkCriteria.AddCondition("regardingobjectid", ConditionOperator.Null);

                EntityCollection ecDocumentLocationResults = service.RetrieveMultiple(qeDocumentLocation);

                if (ecDocumentLocationResults.Entities.Count > 0)
                {
                    tracingService.Trace("0");

                    Helper helper = new Helper(service, context, tracingService, entity);

                    var cred = helper.SetUpCredentials();

                    if (cred == null)
                        return;
                    tracingService.Trace("0.0");

                    Entity document = ecDocumentLocationResults.Entities[0];
                    Guid opportunityId;
                    string patientName = document.GetAttributeValue<AliasedValue>("patient.fullname").Value.ToString();
                    string patientFolderName = document.GetAttributeValue<string>("relativeurl");
                    string referralFolderName = string.Empty;

                    Entity referralEntity = service.Retrieve("opportunity", referral.Id, new ColumnSet("name"));
                    string referralName = referralEntity.Contains("name") ? referralEntity.GetAttributeValue<string>("name") : "";

                    Guid incidentId = Guid.Empty, referralId = Guid.Empty;

                    tracingService.Trace("1");
                    if (document.Contains("referral.sharepointdocumentlocationid") && document.Attributes["referral.sharepointdocumentlocationid"] != null)
                    {
                        referralId = (Guid)document.GetAttributeValue<AliasedValue>("referral.sharepointdocumentlocationid").Value;
                        tracingService.Trace("2");

                        referralFolderName = (string)document.GetAttributeValue<AliasedValue>("referral.relativeurl").Value;
                        tracingService.Trace("3");
                    }
                    else
                    {
                        tracingService.Trace("4");

                        if (document.Contains("opportunity.sharepointdocumentlocationid") && document.Attributes["opportunity.sharepointdocumentlocationid"] != null)
                        {
                            tracingService.Trace("5");

                            opportunityId = (Guid)document.GetAttributeValue<AliasedValue>("opportunity.sharepointdocumentlocationid").Value;
                            tracingService.Trace("5.1");

                        }
                        else
                        {
                            tracingService.Trace("6");

                            opportunityId = helper.CreateDocumentLocation("opportunity of " + patientName, "opportunity", document.ToEntityReference(), null);
                            tracingService.Trace("7");

                            helper.CreateFolderNew($"{cred["SHAREPOINT_SITE_URL"]}/_api/Web/GetFolderByServerRelativeUrl(@v)/Folders/add('opportunity')?@v='{cred["SITE_RELATIVE_URL"]}/contact/{patientFolderName.Replace("'", "%27%27")}'", cred["accessToken"]);
                            tracingService.Trace("8");

                        }

                        tracingService.Trace(referralName + " " + referral.Id.ToString());
                        referralFolderName = helper.CreateFolderName(referralName, referral.Id.ToString());
                        tracingService.Trace("9");

                        referralId = helper.CreateDocumentLocation("Documents of " + referralName, referralFolderName, new EntityReference("sharepointdocumentlocation", opportunityId), referral);
                        tracingService.Trace("10");

                        helper.CreateFolderNew($"{cred["SHAREPOINT_SITE_URL"]}/_api/Web/GetFolderByServerRelativeUrl(@v)/Folders/add('{referralFolderName.Replace("'", "%27%27")}')?@v='{cred["SITE_RELATIVE_URL"]}/contact/{patientFolderName.Replace("'", "%27%27")}/opportunity'", cred["accessToken"]);
                        tracingService.Trace("11");

                    }


                    if (document.Contains("incident.sharepointdocumentlocationid") && document.Attributes["incident.sharepointdocumentlocationid"] != null)
                    {
                        incidentId = (Guid)document.GetAttributeValue<AliasedValue>("incident.sharepointdocumentlocationid").Value;
                    }
                    else
                    {
                        tracingService.Trace("12");

                        incidentId = helper.CreateDocumentLocation("incident of " + referralName, "incident", new EntityReference("sharepointdocumentlocation", referralId), null);
                        tracingService.Trace("13");

                        helper.CreateFolderNew($"{cred["SHAREPOINT_SITE_URL"]}/_api/Web/GetFolderByServerRelativeUrl(@v)/Folders/add('incident')?@v='{cred["SITE_RELATIVE_URL"]}/contact/{patientFolderName.Replace("'", "%27%27")}/opportunity/{referralFolderName.Replace("'", "%27%27")}'", cred["accessToken"]);
                    }

                    tracingService.Trace("14");

                    string caseNumber = entity.Contains("ticketnumber") ? entity.GetAttributeValue<string>("ticketnumber") : "";
                    tracingService.Trace("15");

                    string caseFolderName = helper.CreateFolderName(caseNumber, entity.Id.ToString());
                    tracingService.Trace("16");

                    helper.CreateDocumentLocation("Documents of " + caseNumber, caseFolderName, new EntityReference("sharepointdocumentlocation", incidentId), entity.ToEntityReference());
                    tracingService.Trace("17");

                    helper.CreateFolderNew($"{cred["SHAREPOINT_SITE_URL"]}/_api/Web/GetFolderByServerRelativeUrl(@v)/Folders/add('{caseFolderName.Replace("'", "%27%27")}')?@v='{cred["SITE_RELATIVE_URL"]}/contact/{patientFolderName.Replace("'", "%27%27")}/opportunity/{referralFolderName.Replace("'", "%27%27")}/incident'", cred["accessToken"]);
                    tracingService.Trace("18");

                }


            }
        }



        private void SetCaseOwnerWhenCaseProcessingLocationIsEmpty()
        {

            if ((!entity.Contains("mzkah_serviceprovider") && context.MessageName.ToLower() == "create") ||
                (entity.Contains("mzkah_serviceprovider") && entity.Attributes["mzkah_serviceprovider"] == null && context.MessageName.ToLower() == "update"
                && context.PreEntityImages.Contains("PreValidationCaseImage")))
            {

                EntityReference erReferral = null;
                EntityReference erCaseCreatedBy = null;

                if (context.MessageName.ToLower() == "create")
                {
                    if (entity.Contains("mzk_referral") && entity.Attributes["mzk_referral"] != null)
                    {
                        erReferral = entity.GetAttributeValue<EntityReference>("mzk_referral");
                    }

                    erCaseCreatedBy = entity.GetAttributeValue<EntityReference>("createdby");
                }
                else if (context.MessageName.ToLower() == "update")
                {
                    Entity preImage = context.PreEntityImages["PreValidationCaseImage"] as Entity;

                    if (preImage.Contains("mzk_referral") && preImage.Attributes["mzk_referral"] != null)
                    {
                        erReferral = preImage.GetAttributeValue<EntityReference>("mzk_referral");

                    }

                    erCaseCreatedBy = preImage.GetAttributeValue<EntityReference>("createdby");
                }

                if (erReferral != null)
                {
                    Entity referral = service.Retrieve(erReferral.LogicalName, erReferral.Id, new ColumnSet("ownerid", "owningteam"));

                    EntityReference erReferralOwner = referral.GetAttributeValue<EntityReference>("ownerid");

                    if (referral.GetAttributeValue<EntityReference>("ownerid").LogicalName == "team")
                    {

                        entity["ownerid"] = erReferralOwner;
                    }
                    else if (erCaseCreatedBy != null)
                    {

                        entity["ownerid"] = erCaseCreatedBy;
                    }

                }
            }
        }

        private void ShareCaseWithReferral()
        {
            if (context.MessageName.ToLower() == "create")
            {
                if (entity.Contains("mzk_referral") && entity.Attributes["mzk_referral"] != null)
                {
                    EntityReference erReferral = entity.GetAttributeValue<EntityReference>("mzk_referral");
                    Entity referral = service.Retrieve(erReferral.LogicalName, erReferral.Id, new ColumnSet("ownerid"));

                    if (referral.Contains("ownerid") && referral.Attributes["ownerid"] != null)
                    {
                        EntityReference erOwner = referral.GetAttributeValue<EntityReference>("ownerid");

                        GrantAccessRequest grantRequest = new GrantAccessRequest()
                        {
                            Target = new EntityReference(entity.LogicalName, entity.Id),
                            PrincipalAccess = new PrincipalAccess()
                            {
                                Principal = erOwner,
                                AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess
                            }
                        };

                        GrantAccessResponse granted = (GrantAccessResponse)service.Execute(grantRequest);

                    }
                }
            }
        }

        private void ShareUnSharePatientRecordWithDefaultTeam()
        {
            bool ownerIsTeam = false;
            bool previousOwnerWasTeam = false;
            Entity casePreImage = new Entity();
            Entity casePostImage = new Entity();
            string referralOwnerId = "00000000-0000-0000-0000-000000000000";

            if (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
            {
                if (entity.Contains("ownerid") && entity.Attributes["ownerid"] != null)
                {
                    EntityReference owner = entity.GetAttributeValue<EntityReference>("ownerid");
                    if (owner.LogicalName == "team")
                    {
                        ownerIsTeam = true;
                    }

                    EntityReference patient = null, referral = null;

                    if (context.MessageName.ToLower() == "create")
                    {
                        if (entity.Contains("customerid") && entity.Attributes["customerid"] != null)
                        {
                            patient = entity.GetAttributeValue<EntityReference>("customerid");
                        }
                        if (entity.Contains("mzk_referral") && entity.Attributes["mzk_referral"] != null)
                        {

                            referral = entity.GetAttributeValue<EntityReference>("mzk_referral");
                        }
                    }
                    else if (context.MessageName.ToLower() == "update")
                    {
                        if (!context.PreEntityImages.Contains("CasePreImageForShare"))
                            return;

                        if (!context.PostEntityImages.Contains("CasePostImageForShare"))
                            return;

                        casePreImage = context.PreEntityImages["CasePreImageForShare"] as Entity;
                        casePostImage = context.PostEntityImages["CasePostImageForShare"] as Entity;



                        if (casePostImage.Contains("customerid") && casePostImage.Attributes["customerid"] != null)
                        {
                            patient = casePostImage.GetAttributeValue<EntityReference>("customerid");
                        }
                        if (casePostImage.Contains("mzk_referral") && casePostImage.Attributes["mzk_referral"] != null)
                        {
                            referral = casePostImage.GetAttributeValue<EntityReference>("mzk_referral");
                        }

                        if (casePreImage.GetAttributeValue<EntityReference>("ownerid").LogicalName == "team")
                        {
                            previousOwnerWasTeam = true;
                        }
                    }


                    if (referral != null)
                    {
                        QueryExpression qeReferral = new QueryExpression()
                        {
                            EntityName = referral.LogicalName,
                            ColumnSet = new ColumnSet("ownerid", "owningteam", "owninguser")
                        };
                        qeReferral.Criteria.AddCondition(referral.LogicalName + "id", ConditionOperator.Equal, referral.Id);

                        LinkEntity leRACIntake = qeReferral.AddLink("mzkahrac_rapidaccessclinicintake", referral.LogicalName + "id", "mzkahrac_referral", JoinOperator.LeftOuter);
                        leRACIntake.EntityAlias = "racintake";
                        leRACIntake.Columns = new ColumnSet("mzkahrac_rapidaccessclinicintakeid", "ownerid");
                        EntityCollection ecReferral = service.RetrieveMultiple(qeReferral);

                        referralOwnerId = ecReferral.Entities[0].GetAttributeValue<EntityReference>("ownerid").Id.ToString();
                        if (ownerIsTeam)
                        {


                            // grant access to referral
                            GrantAccessRequest grantRequest = new GrantAccessRequest()
                            {
                                Target = referral,
                                PrincipalAccess = new PrincipalAccess()
                                {
                                    Principal = owner,
                                    AccessMask = AccessRights.ReadAccess
                                }
                            };



                            GrantAccessResponse granted = (GrantAccessResponse)service.Execute(grantRequest);


                            // grant access to rapid access clinic intake

                            //EntityCollection ecReferral = service.RetrieveMultiple(qeReferral);
                            if (ecReferral.Entities.Count > 0)
                            {

                                foreach (Entity eReferral in ecReferral.Entities)
                                {
                                    if (eReferral.Contains("racintake.mzkahrac_rapidaccessclinicintakeid") && eReferral.Attributes["racintake.mzkahrac_rapidaccessclinicintakeid"] != null)
                                    {
                                        Guid guidIntake = (Guid)eReferral.GetAttributeValue<AliasedValue>("racintake.mzkahrac_rapidaccessclinicintakeid").Value;
                                        EntityReference erIntake = new EntityReference("mzkahrac_rapidaccessclinicintake", guidIntake);
                                        GrantAccessRequest grantRequestIntake = new GrantAccessRequest()
                                        {
                                            Target = erIntake,
                                            PrincipalAccess = new PrincipalAccess()
                                            {
                                                Principal = owner,
                                                AccessMask = AccessRights.ReadAccess
                                            }
                                        };



                                        GrantAccessResponse intakeRequstGranted = (GrantAccessResponse)service.Execute(grantRequestIntake);

                                    }
                                }

                            }

                        }

                        if ((owner != casePreImage.GetAttributeValue<EntityReference>("ownerid")) && context.MessageName.ToLower() != "create")
                        {


                            if (previousOwnerWasTeam &&
                                ecReferral.Entities[0].GetAttributeValue<EntityReference>("ownerid").Id.ToString()
                                != casePreImage.GetAttributeValue<EntityReference>("ownerid").Id.ToString())
                            {


                                RevokeAccessRequest revokeRequest = new RevokeAccessRequest()
                                {
                                    Target = referral,
                                    Revokee = casePreImage.GetAttributeValue<EntityReference>("ownerid"),

                                };



                                RevokeAccessResponse revoked = (RevokeAccessResponse)service.Execute(revokeRequest);


                                //EntityCollection ecReferral = service.RetrieveMultiple(qeReferral);

                                if (ecReferral.Entities.Count > 0)
                                {

                                    foreach (Entity eReferral in ecReferral.Entities)
                                    {
                                        if (eReferral.Contains("racintake.mzkahrac_rapidaccessclinicintakeid") && eReferral.Attributes["racintake.mzkahrac_rapidaccessclinicintakeid"] != null)
                                        {
                                            if (((EntityReference)eReferral.GetAttributeValue<AliasedValue>("racintake.ownerid").Value).Id.ToString() != casePreImage.GetAttributeValue<EntityReference>("ownerid").Id.ToString())
                                            {
                                                Guid guidIntake = (Guid)eReferral.GetAttributeValue<AliasedValue>("racintake.mzkahrac_rapidaccessclinicintakeid").Value;
                                                EntityReference erIntake = new EntityReference("mzkahrac_rapidaccessclinicintake", guidIntake);
                                                RevokeAccessRequest revokeRequestIntake = new RevokeAccessRequest()
                                                {
                                                    Target = erIntake,
                                                    Revokee = casePreImage.GetAttributeValue<EntityReference>("ownerid"),

                                                };



                                                RevokeAccessResponse intakeRequestRevoked = (RevokeAccessResponse)service.Execute(revokeRequestIntake);
                                            }

                                        }
                                    }

                                }
                            }

                        }
                    }

                    if (patient != null)
                    {


                        Entity patientEntity = service.Retrieve(patient.LogicalName, patient.Id, new ColumnSet("ownerid"));
                        if (ownerIsTeam)
                        {


                            GrantAccessRequest grantRequest = new GrantAccessRequest()
                            {
                                Target = patient,
                                PrincipalAccess = new PrincipalAccess()
                                {
                                    Principal = owner,
                                    AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess
                                }
                            };


                            GrantAccessResponse granted = (GrantAccessResponse)service.Execute(grantRequest);


                        }


                        if ((owner != casePreImage.GetAttributeValue<EntityReference>("ownerid")) && context.MessageName.ToLower() != "create" &&
                            patientEntity.GetAttributeValue<EntityReference>("ownerid").Id.ToString() != casePreImage.GetAttributeValue<EntityReference>("ownerid").Id.ToString()
                        && casePreImage.GetAttributeValue<EntityReference>("ownerid").Id.ToString() != referralOwnerId)

                        {


                            if (previousOwnerWasTeam)
                            {


                                RevokeAccessRequest revokeRequest = new RevokeAccessRequest()
                                {
                                    Target = patient,
                                    Revokee = casePreImage.GetAttributeValue<EntityReference>("ownerid"),

                                };



                                RevokeAccessResponse revoked = (RevokeAccessResponse)service.Execute(revokeRequest);


                            }

                        }

                    }

                }

            }

        }

        private void CreateTimeLinePostWhenCaseAcceptedOrDeclined()
        {
            if (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
            {


                int triageOutcomeAccept = 432530001;
                int triageOutcomeDecline = 432530002;

                if (entity.Contains("mzkahrac_triageutcome") && entity.Attributes["mzkahrac_triageutcome"] != null)
                {

                    int triageOutcomeValue = entity.GetAttributeValue<OptionSetValue>("mzkahrac_triageutcome").Value;

                    if (!(triageOutcomeValue == triageOutcomeAccept || triageOutcomeValue == triageOutcomeDecline))
                        return;

                    Entity preImage = context.PreEntityImages["CasePreImage"] as Entity;
                    string caseProcessingLocation = string.Empty, requestedProvider = string.Empty;



                    if (preImage.Contains("mzkah_requestedprovider") && preImage.Attributes["mzkah_requestedprovider"] != null && preImage.Contains("mzkah_serviceprovider") && preImage.Attributes["mzkah_serviceprovider"] != null)
                    {
                        // providerEntity = service.Retrieve("contact", preImage.GetAttributeValue<EntityReference>("mzkah_requestedprovider").Id, new ColumnSet("fullname"));
                        //Entity caseProcessingLocationEntity = service.Retrieve("account", preImage.GetAttributeValue<EntityReference>("mzkah_serviceprovider").Id, new ColumnSet("name"));

                        //if (providerEntity.Contains("fullname") && caseProcessingLocationEntity.Contains("name"))
                        //{
                        requestedProvider = preImage.GetAttributeValue<EntityReference>("mzkah_requestedprovider").Name;
                        caseProcessingLocation = preImage.GetAttributeValue<EntityReference>("mzkah_serviceprovider").Name;
                        //}
                    }

                    string message = "[PROVIDER] from [CLINIC] has [ACTION] the case.";
                    if (requestedProvider != string.Empty && caseProcessingLocation != string.Empty)
                    {
                        message = message.Replace("[PROVIDER]", requestedProvider);
                        message = message.Replace("[CLINIC]", caseProcessingLocation);
                    }
                    else
                    {
                        message = "The case has been [ACTION].";
                    }
                    if (triageOutcomeValue == triageOutcomeAccept)
                    {

                        CreateTimelinePost(message.Replace("[ACTION]", "accepted"));
                    }
                    else if (triageOutcomeValue == triageOutcomeDecline)
                    {
                        CreateTimelinePost(message.Replace("[ACTION]", "declined"));
                    }
                }
            }
        }

        private void CreateTimelinePost(string message)
        {
            Entity timelinePost = new Entity("post");
            timelinePost["regardingobjectid"] = new EntityReference(entity.LogicalName, entity.Id);

            timelinePost["text"] = message;

            timelinePost["source"] = new OptionSetValue(1); // Auto-Post
            service.Create(timelinePost);
        }

        //private void SetOwnerOnCaseFromPreferredUser()
        //{
        //	if(context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
        //	{
        //		if (entity.Contains("mzkah_requestedprovider") && entity["mzkah_requestedprovider"] != null) // Requested Provider
        //		{
        //			Entity provider = service.Retrieve("contact", entity.GetAttributeValue<EntityReference>("mzkah_requestedprovider").Id, new ColumnSet("preferredsystemuserid"));
        //			if (provider.Contains("preferredsystemuserid"))
        //			{
        //				entity["ownerid"] = provider.GetAttributeValue<EntityReference>("preferredsystemuserid");
        //			}
        //		}
        //	}

        //}

        private void SetOwnerOnCasePatientFromDefaultTeam()
        {
            if (context.MessageName.ToLower() == "create" || context.MessageName.ToLower() == "update")
            {
                if (entity.Contains("mzkah_serviceprovider") && entity["mzkah_serviceprovider"] != null) // Case Processing Location
                {

                    Entity caseProcessingLocation = service.Retrieve("account", entity.GetAttributeValue<EntityReference>("mzkah_serviceprovider").Id, new ColumnSet("mzkah_defaultteam"));

                    if (caseProcessingLocation.Contains("mzkah_defaultteam") && caseProcessingLocation.Attributes["mzkah_defaultteam"] != null)
                    {

                        EntityReference defaultTeam = caseProcessingLocation.GetAttributeValue<EntityReference>("mzkah_defaultteam");
                        //Guid? patientId = null;
                        //if (context.MessageName.ToLower() == "create")
                        //{

                        //	if (entity.Contains("customerid"))
                        //	{

                        //		patientId = entity.GetAttributeValue<EntityReference>("customerid").Id;
                        //	}
                        //}
                        //else if (context.MessageName.ToLower() == "update")
                        //{

                        //	if (!context.PostEntityImages.Contains("CasePostImage"))
                        //		return;

                        //	Entity casePostImage = context.PostEntityImages["CasePostImage"] as Entity;
                        //	{

                        //		patientId = casePostImage.GetAttributeValue<EntityReference>("customerid").Id;
                        //	}
                        //}

                        //Entity _case = new Entity("incident", entity.Id);
                        //_case["ownerid"] = defaultTeam;
                        //service.Update(_case);

                        AssignRequest request = new AssignRequest();

                        request.Target = new EntityReference("incident", entity.Id);

                        request.Assignee = defaultTeam;

                        try

                        {
                            AssignResponse response = (AssignResponse)service.Execute(request);

                        }

                        catch (FaultException<OrganizationServiceFault> ex)

                        {



                            throw new InvalidPluginExecutionException(ex.Message);

                        }

                        //if (patientId != null)
                        //{

                        //	//Entity patient = new Entity("contact", (Guid)patientId);
                        //	//patient["ownerid"] = defaultTeam;
                        //	//service.Update(patient);

                        //	AssignRequest requestPatient = new AssignRequest();

                        //	requestPatient.Target = new EntityReference("contact", (Guid)patientId);

                        //	requestPatient.Assignee = defaultTeam;

                        //	try
                        //	{

                        //		AssignResponse response = (AssignResponse)service.Execute(requestPatient);

                        //	}

                        //	catch (FaultException<OrganizationServiceFault> ex)
                        //	{
                        //		 

                        //		throw new InvalidPluginExecutionException(ex.Message);
                        //	}
                        //}

                    }
                }
            }
        }

        private void SetTriageOutcomeOnManualCaseProcessingLocationAssignment()
        {
            if (entity.Contains("mzkah_serviceprovider") && entity.Attributes["mzkah_serviceprovider"] != null && context.PreEntityImages.Contains("CasePreImage"))
            {

                int caseTriagePendingValue = 432530000;
                Entity preImage = context.PreEntityImages["CasePreImage"] as Entity;
                if (entity.Contains("mzkahrac_triageutcome") && entity.Attributes["mzkahrac_triageutcome"] != null &&
                    entity.GetAttributeValue<OptionSetValue>("mzkahrac_triageutcome").Value != caseTriagePendingValue)
                {
                    entity.Attributes["mzkahrac_triageutcome"] = new OptionSetValue(caseTriagePendingValue);
                }
                else
                {
                    if (!preImage.Contains("mzkahrac_triageutcome") || (preImage.Contains("mzkahrac_triageutcome") &&
                        preImage.GetAttributeValue<OptionSetValue>("mzkahrac_triageutcome").Value != caseTriagePendingValue))
                    {
                        entity.Attributes["mzkahrac_triageutcome"] = new OptionSetValue(caseTriagePendingValue);
                    }
                }
                //if (entity.Contains("mzk_casestatus") && entity.Attributes["mzk_casestatus"] != null)
                //{
                //	entity.Attributes["mzk_casestatus"] = null;
                //}
                //else
                //{
                //	if (preImage.Contains("mzk_casestatus"))
                //	{
                //		entity.Attributes["mzk_casestatus"] = null;
                //	}
                //}

            }
        }


        [DataContract]
        private class GeoMatchResponse
        {
            [DataMember]
            public int geoMatchingStatus { get; set; }
            [DataMember]
            public string errorMessage { get; set; }
        }
    }

}
