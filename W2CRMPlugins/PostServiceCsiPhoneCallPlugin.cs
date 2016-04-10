using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins.Common;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    public class PostServiceCsiPhoneCallPlugin : IPlugin
    {

        private const string C_EntityName = "salesorder";
        private const string C_ImageName = "Image";
        private const string C_CsiFlagFieldName = "new_startcsiflag";
        private const string C_VehicleOwnerFieldName = "customerid";

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
                if (postImage.Contains("new_repaireddatetime"))
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

        private static void CreatePhoneCall(Entity order, IOrganizationService orgService)
        {
            if (order.Contains(C_VehicleOwnerFieldName) == false)
            {
                throw new Exception("订单" + order.Id.ToString() + "缺少客户信息，无法创建CSI电话联络");
            }
            if (order.Contains("new_vehicleid") == false)
            {
                throw new Exception("订单" + order.Id.ToString() + "缺少车辆信息");
            }

            /*判断是否G或F帐户，是则退出
             * To be added
            */

            /*是否P部门，是则退出
             * To be added
             */
            if (order.Contains("new_dmsdept"))
            {
                if (order["new_dmsdept"].ToString() == "P")
                    return;
            }
            
            EntityReference vehicleRef = order["new_vehicleid"] as EntityReference;
            Entity vehicle = orgService.Retrieve(vehicleRef.LogicalName, vehicleRef.Id, new ColumnSet(true));
            Entity account;
            bool ownerIsKR = CheckForKR(vehicle, orgService, out account);
            if (ownerIsKR == true)
            {
                return;
            }

            Entity phone = new Entity("phonecall");
            /*要增加性别、车型、车龄到标题上
             * To be added
             */
            string sSex = "",sVehicleType="",sThreeYearDate="";
            DateTime ThreeYearDate;
            if (account.Contains("new_sex"))
            {
                sSex = (((OptionSetValue)account["new_sex"]).Value == 100000000 ? "先生" : "女士");
            }
            if (vehicle.Contains("new_vehicletypetext"))
            {
                sVehicleType = vehicle["new_vehicletypetext"].ToString();
            }
            if (vehicle.Contains("new_actualregistrationdate"))
            {
                ThreeYearDate = ((DateTime)vehicle["new_actualregistrationdate"]).AddYears(3);
                sThreeYearDate = (ThreeYearDate >= DateTime.Today ? "<3年" : ">3年");
            }
            phone["subject"] = "CSI_维修后_" + (order[C_VehicleOwnerFieldName] as EntityReference).Name
                +sSex+"_"+sVehicleType+"_"+sThreeYearDate;
            
            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = order[C_VehicleOwnerFieldName] as EntityReference;
            phone["to"] = new EntityCollection(new List<Entity>() { toActivityParty });

            phone["directioncode"] = true;
            if (order.Contains("new_dealership") == false)
            {
                throw new Exception("订单" + order.Id + "缺少经销商信息");
            }

            if (account.Contains("telephone1") == false && account.Contains("telephone2") == false && account.Contains("telephone3") == false)
            {
                throw new Exception("客户" + account.Id.ToString() + "Mobile Phone、Office Phone、Home Phone均无信息");
            }
            string phoneNumber =
                account.Contains("telephone1") ? account["telephone1"].ToString() :
                (account.Contains("telephone2") ? account["telephone2"].ToString() :
                ((account.Contains("telephone3") ? account["telephone3"].ToString() : string.Empty)));
            phone["phonenumber"] = phoneNumber;
            phone["new_phonetasktype"] = new OptionSetValue(100000005); //售后服务CSI电话
            if (order.Contains("new_repaireddatetime"))
                phone["new_taskendtime"] = ((DateTime)order["new_repaireddatetime"]).AddDays(3);

            phone["new_dealer"] = order["new_dealership"] as EntityReference;
            /*由工作流配置，根据经销商信息，获取呼叫方以及负责人
            EntityReference postSaleCsiOwner = QueryOwner.QueryPostSaleCsiOwner(vehicle, orgService);
            Entity fromActivityParty = new Entity("activityparty");
            fromActivityParty["partyid"] = postSaleCsiOwner;
            phone["from"] = new EntityCollection(new List<Entity>() { fromActivityParty });

            phone["ownerid"] = postSaleCsiOwner;
             */
            phone["regardingobjectid"] = order.ToEntityReference();

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

            Guid accountId = (vehicle["new_vehicleowner"] as EntityReference).Id;
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
