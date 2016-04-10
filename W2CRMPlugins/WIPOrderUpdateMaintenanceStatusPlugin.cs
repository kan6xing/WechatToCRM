
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
    /// 根据WIP订单的变化更新车辆的保养状态
    /// </summary>
    public class WIPOrderUpdateMaintenanceStatusPlugin : IPlugin
    {

        private const string C_EntityName = "salesorder";
        private const string C_ImageName = "Image";
        private const string C_VehicleEntityName = "new_vehiclefiles";

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
            Entity postOrder = context.PostEntityImages[C_ImageName]; //orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("customerid"));
            if (postOrder.Contains("new_ordertype") != false && postOrder.Contains("new_vehicleid") != false)
            {
                if (((OptionSetValue)postOrder["new_ordertype"]).Value == 100000001)   //如果是WIP订单则处理
                    if (postOrder.Contains("description"))
                    {
                        if (postOrder["description"].ToString().Contains("保养") ||
                            postOrder["description"].ToString().Contains("首保"))     //如果是保养或首保WIP
                        {
                            UpdateVehicleMaintenance_Update(context, orgService);
                        }
                    }
            }
        }

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity postOrder = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_ordertype", "description", "new_vehicleid", "new_repaireddatetime"));

            if (postOrder.Contains("new_ordertype") != false && postOrder.Contains("new_vehicleid") != false)
            {
                if (((OptionSetValue)postOrder["new_ordertype"]).Value == 100000001)   //如果是WIP订单则处理
                    if (postOrder.Contains("description"))
                    {
                        if (postOrder["description"].ToString().Contains("保养") || 
                            postOrder["description"].ToString().Contains("首保"))     //如果是保养或首保WIP
                        {
                            UpdateVehicleMaintenance_Create(postOrder, orgService);
                        }
                    }
            }

        }

        private void UpdateVehicleMaintenance_Create(Entity postOrder, IOrganizationService orgService)
        {
            Entity Vehicle = orgService.Retrieve(C_VehicleEntityName, ((EntityReference)postOrder["new_vehicleid"]).Id,
                new ColumnSet("new_maintenancestage", "new_firstinspection"));
            bool bIsExistsRepairDate = IsExistsRepairedDate(postOrder);  //是否存在维修日期

            if (Vehicle.Contains("new_maintenancestage"))
            {
                OptionSetValue statusValue = (OptionSetValue)Vehicle["new_maintenancestage"];

                if (bIsExistsRepairDate)   //存在维修日期
                {
                    //保养状态为"完工状态"
                    if (statusValue.Value == 100000005)
                    {
                        return;
                    }
                    else
                    {
                        Vehicle["new_maintenancestage"] = new OptionSetValue(100000005);//保养状态置为“完工状态”
                        ToCloseServiceActivity(Vehicle,orgService);                //关闭服务活动

                        if (postOrder["description"].ToString().Contains("首保"))     //如果是首保WIP
                            Vehicle["new_firstinspection"] = new OptionSetValue(100000002);//首保状态置为“完成”
                    }
                }
                else
                {
                    //保养状态为"车辆进厂"
                    if (statusValue.Value == 100000004)
                    {
                        return;
                    }
                    else                        //保养状态置为“车辆进厂”
                        Vehicle["new_maintenancestage"] = new OptionSetValue(100000004);

                }
            }
            else      //保养状态为空时
            {
                if (bIsExistsRepairDate)   //存在维修日期，保养状态置为“完工状态”
                {
                    Vehicle["new_maintenancestage"] = new OptionSetValue(100000005);
                    ToCloseServiceActivity(Vehicle, orgService);                //关闭服务活动
                    
                    if (postOrder["description"].ToString().Contains("首保"))     //如果是首保WIP
                        Vehicle["new_firstinspection"] = new OptionSetValue(100000002);//首保状态置为“完成”
                }
                else           //不存在维修日期，保养状态置为“车辆进厂”
                    Vehicle["new_maintenancestage"] = new OptionSetValue(100000004);
            }

            orgService.Update(Vehicle);
        }

        private void ToCloseServiceActivity(Entity Vehicle, IOrganizationService orgService)
        {
            EntityCollection ServActs;
            QueryExpression query = new QueryExpression
            {
                EntityName = "serviceappointment",
                //ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "regardingobjectid",
                                Operator = ConditionOperator.Equal,
                                Values = { Vehicle.Id.ToString() }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "statecode",
                                Operator = ConditionOperator.NotEqual,
                                Values = { 1 }
                            },

                        }
                }
            };
            ServActs = orgService.RetrieveMultiple(query);

            if (ServActs.Entities.Count > 0)
            {
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = ServActs[0].Id,
                        LogicalName = ServActs[0].LogicalName
                    },
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(8)
                };

                orgService.Execute(setStateRequest);
            }
        }

        private void UpdateVehicleMaintenance_Update(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            DateTime? preRepairedDate = GetRepairedDate(preImage);
            DateTime? postRepairedDate = GetRepairedDate(postImage);

            if (preRepairedDate == postRepairedDate)
                return;
            else
            {
                Entity Vehicle = orgService.Retrieve(C_VehicleEntityName, ((EntityReference)postImage["new_vehicleid"]).Id,
                    new ColumnSet("new_maintenancestage", "new_firstinspection"));

                bool bIsExistsRepairDate = IsExistsRepairedDate(postImage);  //是否存在维修日期

                if (Vehicle.Contains("new_maintenancestage"))
                {
                    OptionSetValue statusValue = (OptionSetValue)Vehicle["new_maintenancestage"];

                    if (bIsExistsRepairDate)   //存在维修日期
                    {
                        //保养状态为"完工状态"
                        if (statusValue.Value == 100000005)
                        {
                            return;
                        }
                        else                     //保养状态置为“完工状态”
                        {
                            Vehicle["new_maintenancestage"] = new OptionSetValue(100000005);
                            ToCloseServiceActivity(Vehicle, orgService);                //关闭服务活动
                        }
                        
                        if (postImage.Contains("description"))
                            if (postImage["description"].ToString().Contains("首保"))     //如果是首保WIP
                                Vehicle["new_firstinspection"] = new OptionSetValue(100000002);//首保状态置为“完成”
                    }
                    else
                    {
                        return;
                    }
                }
                else      //保养状态为空时
                {
                    if (bIsExistsRepairDate)   //存在维修日期，保养状态置为“完工状态”
                    {
                        Vehicle["new_maintenancestage"] = new OptionSetValue(100000005);
                        ToCloseServiceActivity(Vehicle, orgService);                //关闭服务活动

                        if (postImage.Contains("description"))
                            if (postImage["description"].ToString().Contains("首保"))     //如果是首保WIP
                                Vehicle["new_firstinspection"] = new OptionSetValue(100000002);//首保状态置为“完成”
                    }
                    else           //不存在维修日期
                        return;
                }

                orgService.Update(Vehicle);
            }
        }

        private DateTime? GetRepairedDate(Entity order)
        {
            if (order.Contains("new_repaireddatetime"))
            {
                return (DateTime?)order["new_repaireddatetime"];
            }
            else
            {
                return null;
            }
        }

        private bool IsExistsRepairedDate(Entity postOrder)
        {
            if (postOrder.Contains("new_repaireddatetime") == false)
            {
                return false;
            }
            else
            {
                return true;
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
