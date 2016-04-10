using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins.Common;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    internal class FirstYearNpsPhoneCallPlugin:IPlugin
    {
        private const string C_EntityName = "new_vehiclefiles";
        private const string C_ImageName = "Image";
        private const string C_FirstYearNpsFlagFieldName = "new_yearnps";
        private const string C_VehicleOwnerFieldName = "new_vehicleowner";

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

                Entity preImage = context.PreEntityImages[C_ImageName];
                Entity postImage = context.PostEntityImages[C_ImageName];

                bool? preNpsFlag = GetNpsFlag(preImage);
                bool? postNpsFlag = GetNpsFlag(postImage);

                if (NeedProcess(preNpsFlag, postNpsFlag) == false)
                {
                    return;
                }

                CreatePhoneCall(postImage, orgService);
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

        private static void CreatePhoneCall(Entity vehicle, IOrganizationService orgService)
        {
            if (vehicle.Contains(C_VehicleOwnerFieldName) == false)
            {
                throw new Exception("车辆" + vehicle.Id.ToString() + "缺少车主信息，无法创建NPS电话联络");
            }
            
            bool ownerIsKR = CheckForKR(vehicle, orgService);
            if (ownerIsKR == true)
            {
                return;
            }

            Entity phone = new Entity("phonecall");
            phone["subject"] = "NPS_售后_" + (vehicle[C_VehicleOwnerFieldName] as EntityReference).Name;

            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = vehicle[C_VehicleOwnerFieldName] as EntityReference;
            phone["to"] = new EntityCollection(new List<Entity>() { toActivityParty });

            phone["directioncode"] = true;
            if (vehicle.Contains("new_dealership") == false)
            {
                throw new Exception("车辆" + vehicle.Id + "缺少所属经销商信息");
            }
            /*由工作流配置，根据经销商信息，获取呼叫方以及负责人
            EntityReference postSaleCsiOwner = QueryOwner.QueryPostSaleCsiOwner(vehicle, orgService);
            Entity fromActivityParty = new Entity("activityparty");
            fromActivityParty["partyid"] = postSaleCsiOwner;
            phone["from"] = new EntityCollection(new List<Entity>() { fromActivityParty });

            phone["ownerid"] = postSaleCsiOwner;
             */
           
            phone["regardingobjectid"] = vehicle.ToEntityReference();

            orgService.Create(phone);
        }

        private static bool CheckForKR(Entity vehicle, IOrganizationService orgService)
        {
            if (vehicle.Contains("new_brand") == false)
            {
                throw new Exception("车辆" + vehicle.Id.ToString() + "缺少品牌信息");
            }
            EntityReference brandRef = vehicle["new_brand"] as EntityReference;
            Entity brand = orgService.Retrieve(brandRef.LogicalName, brandRef.Id, new ColumnSet(true));
            if (brand.Contains("new_krflag") == false)
            {
                throw new Exception("品牌" + brand["new_name"].ToString() + "缺少KR Field In Acc Form属性信息");
            }
            string krField = brand["new_krflag"].ToString();

            Guid accountId = (vehicle["new_vehicleowner"] as EntityReference).Id;
            Entity account = orgService.Retrieve("account", accountId, new ColumnSet(true));
            if (account.Contains(krField) == false)
            {
                throw new Exception("客户实体中不包含" + krField + "字段，或该字段值为空");
            }
            return (bool)account[krField];
        }

        private bool NeedProcess(bool? preNpsFlag, bool? postNpsFlag)
        {
            if (preNpsFlag.HasValue && preNpsFlag.Value == true)
            {
                return false;
            }

            //pre: null or false
            return postNpsFlag.HasValue && postNpsFlag.Value == true;
        }

        private bool? GetNpsFlag(Entity vehicle)
        {
            if (vehicle.Contains(C_FirstYearNpsFlagFieldName) == true)
            {
                return (bool)vehicle[C_FirstYearNpsFlagFieldName];
            }
            else
            {
                return null;
            }
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

            if (context.MessageName != "Update")
            {
                return false;
            }

            if (context.PreEntityImages.ContainsKey(C_ImageName) == false ||
                context.PostEntityImages.ContainsKey(C_ImageName) == false)
            {
                return false;
            }

            return true;
        }
    }
}
