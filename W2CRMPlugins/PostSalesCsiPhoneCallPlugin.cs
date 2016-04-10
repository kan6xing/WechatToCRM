using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins.Common;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    public class PostSalesCsiPhoneCallPlugin : IPlugin
    {

        private const string C_EntityName = "new_vehiclefiles";
        private const string C_ImageName = "Image";
        private const string C_CsiFlagFieldName = "new_salescsi";
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

                bool? preCsiFlag = GetCsiFlag(preImage);
                bool? postCsiFlag = GetCsiFlag(postImage);

                if (NeedProcess(preCsiFlag, postCsiFlag) == false)
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
                throw new Exception("车辆" + vehicle.Id.ToString() + "缺少车主信息，无法创建CSI电话联络");
            }
            Entity account;
            bool ownerIsKR = CheckForKR(vehicle, orgService, out account);
            if (ownerIsKR == true)
            {
                return;
            }
            Entity phone = new Entity("phonecall");
            phone["subject"] = "CSI_新车交付_" + (vehicle[C_VehicleOwnerFieldName] as EntityReference).Name;

            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = vehicle[C_VehicleOwnerFieldName] as EntityReference;
            phone["to"] = new EntityCollection(new List<Entity>() { toActivityParty });

            phone["directioncode"] = true;
            if (account.Contains("telephone1") == false && account.Contains("telephone2") == false && account.Contains("telephone3") == false)
            {
                throw new Exception("客户" + account.Id.ToString() + "Mobile Phone、Office Phone、Home Phone均无信息");
            }
            string phoneNumber =
                account.Contains("telephone1") ? account["telephone1"].ToString() :
                (account.Contains("telephone2") ? account["telephone2"].ToString() :
                ((account.Contains("telephone3") ? account["telephone3"].ToString() : string.Empty)));
            phone["phonenumber"] = phoneNumber;
            phone["new_phonetasktype"] = new OptionSetValue(100000015);   //销售CSI电话
            if (vehicle.Contains("new_purchasevehicledate"))
                phone["new_taskendtime"] = ((DateTime)vehicle["new_purchasevehicledate"]).AddDays(3);
            if (vehicle.Contains("new_dealership") == false)
            {
                throw new Exception("车辆" + vehicle.Id.ToString() + "缺少所属经销商信息");
            }
            phone["new_dealer"] = vehicle["new_dealership"] as EntityReference;
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

        private static bool CheckForKR(Entity vehicle, IOrganizationService orgService, out Entity account)
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

            Guid accountId = (vehicle[C_VehicleOwnerFieldName] as EntityReference).Id;
            account = orgService.Retrieve("account", accountId, new ColumnSet(true));
            if (account.Contains(krField) == false)
            {
                throw new Exception("客户实体中不包含" + krField + "字段，或该字段值为空");
            }
            return (bool)account[krField];
        }

        private bool NeedProcess(bool? preCsiFlag, bool? postCsiFlag)
        {
            if (preCsiFlag.HasValue && preCsiFlag.Value == true)
            {
                return false;
            }

            //pre: null or false
            return postCsiFlag.HasValue && postCsiFlag.Value == true;
        }

        private bool? GetCsiFlag(Entity vehicle)
        {
            if (vehicle.Contains(C_CsiFlagFieldName) == true)
            {
                return (bool)vehicle[C_CsiFlagFieldName];
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
