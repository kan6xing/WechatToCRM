
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    /// <summary>
    /// 车辆保养状态置为“取消保养”时，取消电话邀约
    /// </summary>
    public class VehicleCancelMaintenancePlugin : IPlugin
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
                    //DoCreate(context, orgService);
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
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            int? preMaintenanceStatus = GetMaintenanceStatus(preImage);
            int? postMaintenanceStatus = GetMaintenanceStatus(postImage);

            if (preMaintenanceStatus != postMaintenanceStatus)   //保养状态有变化
            {
                if (postMaintenanceStatus == 100000007)   //取消保养
                {
                    if (preMaintenanceStatus == 100000004 || preMaintenanceStatus == 100000005 || preMaintenanceStatus == 100000006)
                        throw new InvalidPluginExecutionException(
                            "取消保养时，车辆的保养状态不能处于“车辆进厂”、“完工状态”或“完成满意度回访”等状态之一，保存失败！");
                    else
                    {       //取消T-15电话任务
                        CancelT15Telephone(postImage,orgService);
                    }
                }
            }

        }

        private void CancelT15Telephone(Entity postImage,IOrganizationService orgService)
        {
            EntityCollection telephones;
            Entity phone;
            QueryExpression query = new QueryExpression
            {
                EntityName = "phonecall",
                //ColumnSet = new ColumnSet("new_accountcode",
                //    "totalamount"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "regardingobjectid",
                                Operator = ConditionOperator.Equal,
                                Values = { postImage.Id.ToString() }
                            },
                            new ConditionExpression
                            {    // 状态为开启
                                AttributeName = "statecode",
                                Operator = ConditionOperator.Equal,
                                Values = { 0 }
                            },
                            new ConditionExpression
                            {    // 电话类型为"首保"或"日保"
                                AttributeName = "new_phonetasktype",
                                Operator = ConditionOperator.In,
                                Values = { 100000009,100000010 }
                            }
                        }
                }
            };
            telephones = orgService.RetrieveMultiple(query);

            for (int i = 0; i < telephones.Entities.Count; i++)
            {
                phone = telephones[i];

                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = phone.Id,
                        LogicalName = phone.LogicalName
                    },
                    State = new OptionSetValue(2),
                    Status = new OptionSetValue(3)
                };

                orgService.Execute(setStateRequest);
            }
        }

        private int? GetMaintenanceStatus(Entity vehicle)
        {
            if (vehicle.Contains("new_maintenancestage"))
            {
                return ((OptionSetValue)vehicle["new_maintenancestage"]).Value;
            }
            else
            {
                return 0;
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

            if (context.Depth > 2)
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
