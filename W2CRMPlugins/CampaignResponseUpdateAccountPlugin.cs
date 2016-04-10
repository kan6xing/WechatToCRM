
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;



namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    /// <summary>
    /// 根据车辆信息的变化更新客户状态
    /// </summary>
    public class CampaignResponseUpdateAccount : IPlugin
    {

        private const string C_EntityName = "campaignresponse";
        private const string C_ImageName = "Image";

        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService orgService = factory.CreateOrganizationService(null);

                if (ValidInput(context) == false)
                {
                    return;
                }

                if (context.MessageName == "Create")
                {
                    DoCreate(context, orgService);
                }
                else if (context.MessageName == "Update")
                {
                    DoUpdate(context, orgService);
                }
            }
            catch (FaultException<OrganizationServiceFault> excp)
            {
                throw new InvalidPluginExecutionException(excp.Message);
            }
            catch (Exception excp)
            {
                throw new InvalidPluginExecutionException(excp.Message);
            }
        }

        private void DoUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            //Entity vehicle = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("new_vehicleowner"));
            //Entity preImage = context.PreEntityImages[C_ImageName];
            //Entity postImage = context.PostEntityImages[C_ImageName];
            Entity postImage = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("customer"));

            //EntityReference preVehicleOwner = GetAmountValue(preImage);

            if (postImage.Contains("customer")!=false)
            {
                EntityCollection ec = (EntityCollection)postImage["customer"];
                            postImage["new_customerlookup"] = null;
                for (int i = 0; i < ec.Entities.Count; i++)
                {
                    if ((ec[i]["partyid"] as EntityReference).LogicalName == "account")
                    {
                        RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = C_EntityName,
                            LogicalName = "new_customerlookup",
                            RetrieveAsIfPublished = true
                        };

                        // Execute the request
                        RetrieveAttributeResponse attributeResponse =
                            (RetrieveAttributeResponse)orgService.Execute(attributeRequest);

                        if (attributeResponse != null)
                        {
                            postImage["new_customerlookup"] = ec[i]["partyid"];
                        }
                    }
                }
                            orgService.Update(postImage);
            }
        }


        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            DoUpdate(context, orgService);
        }


        private bool ValidInput(IPluginExecutionContext context)
        {
            if (context.PrimaryEntityName != C_EntityName)
            {
                return false;
            }

            if (context.Stage != 40)
            {
                return false;
            }

            if (context.Depth > 1)
            {
                return false;
            }

            if (context.MessageName != "Update" &&
                context.MessageName != "Create")
            {
                return false;
            }

            //if (context.MessageName == "Update")
            //{
            //    if (context.PostEntityImages.ContainsKey(C_ImageName) == false)
            //    {
            //        return false;
            //    }
            //}

            return true;
        }
    }
}
