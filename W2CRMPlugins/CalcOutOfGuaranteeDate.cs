using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    public class CalcOutOfGuaranteeDate : IPlugin
    {

        private const string C_EntityName = "new_vehiclefiles";
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
            Entity preImage = context.PreEntityImages["Image"];
            Entity postImage = context.PostEntityImages["Image"];

            DateTime? prePurchaseDate = GetPurchaseDate(preImage);
            DateTime? postPurchaseDate = GetPurchaseDate(postImage);

            if (prePurchaseDate == postPurchaseDate)
            {
                return;
            }

            if (postPurchaseDate.HasValue == false)
            {
                Entity vehicle = new Entity(context.PrimaryEntityName);
                vehicle.Id = context.PrimaryEntityId;
                vehicle["new_guaranteeenddate"] = null;

                orgService.Update(vehicle);
                return;
            }
            else
            {
                Entity vehicle = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
                int? guaranteeYears = GetGuaranteeYears(vehicle, orgService);
                if (guaranteeYears == null || guaranteeYears.HasValue == false)
                {
                    return;
                }
                DateTime warrantyEndDate = postPurchaseDate.Value.AddYears(guaranteeYears.Value);

                vehicle = new Entity(context.PrimaryEntityName);
                vehicle.Id = context.PrimaryEntityId;
                vehicle["new_guaranteeenddate"] = warrantyEndDate;
                orgService.Update(vehicle);
                return;
            }
        }

        private DateTime? GetPurchaseDate(Entity preImage)
        {
            if (preImage.Contains("new_purchasevehicledate") == true)
            {
                return (DateTime)preImage["new_purchasevehicledate"];
            }

            return null;
        }

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Guid vehicleId = context.PrimaryEntityId;
            Entity vehicle = (Entity)context.InputParameters["Target"];
            if (vehicle.Contains("new_purchasevehicledate") == false)
            {
                return;
            }

            if (vehicle.Contains("new_guaranteeenddate") == true)
            {
                return;
            }

            DateTime purchaseDate = (DateTime)vehicle["new_purchasevehicledate"];
            int? guaranteeYears = GetGuaranteeYears(vehicle, orgService);
            if (guaranteeYears == null || guaranteeYears.HasValue == false)
            {
                return;
            }
            DateTime warrantyEndDate = purchaseDate.AddYears(guaranteeYears.Value);

            vehicle = new Entity(context.PrimaryEntityName);
            vehicle.Id = vehicleId;
            vehicle["new_guaranteeenddate"] = warrantyEndDate;

            orgService.Update(vehicle);
        }

        private int? GetGuaranteeYears(Entity vehicle, IOrganizationService orgService)
        {
            if (vehicle.Contains("new_brand") == false)
            {
                return null;
            }

            EntityReference brandRef = vehicle["new_brand"] as EntityReference;

            Entity brand = orgService.Retrieve(brandRef.LogicalName, brandRef.Id, new ColumnSet(true));
            if (brand.Contains("new_guaranteeyears") == false)
            {
                return null;
            }

            return Convert.ToInt32((double)brand["new_guaranteeyears"]);
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

            if (context.MessageName == "Update")
            {
                if (context.PreEntityImages.ContainsKey(C_ImageName) == false ||
               context.PostEntityImages.ContainsKey(C_ImageName) == false)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
