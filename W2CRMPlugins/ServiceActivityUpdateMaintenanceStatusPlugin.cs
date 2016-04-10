
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    /// <summary>
    /// 根据服务活动的创建,更新车辆的保养状态
    /// </summary>
    public class ServiceActivityUpdateMaintenanceStatusPlugin : IPlugin
    {

        private const string C_EntityName = "serviceappointment";
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
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            //邀约成功
            string preVehicleID = GetVehichleID(preImage);
            string postVehicleID = GetVehichleID(postImage);

            if (preVehicleID != postVehicleID)
            {
                UpdateVehicleStatus_Success(postImage, orgService);
            }

            //取消保养处理
            int? preServActStatus = GetServActStatus(preImage);
            int? postServActStatus = GetServActStatus(postImage);

            if (preServActStatus != postServActStatus)
            {
                UpdateVehicleStatus_Cancel(postImage, orgService);
            }
        }

        private int? GetServActStatus(Entity ServAct)
        {
            if (ServAct.Contains("new_serviceactivitystatus"))
            {
                return ((OptionSetValue)ServAct["new_serviceactivitystatus"]).Value;
            }
            else
            {
                return 0;
            }
        }

        private string GetVehichleID(Entity seviceAct)
        {
            if (seviceAct.Contains("new_vehicleid"))
            {
                return ((EntityReference)seviceAct["new_vehicleid"]).Id.ToString();
            }
            else
            {
                return "";
            }
        }

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity postServAct = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, 
                new ColumnSet("new_vehicleid"));

            UpdateVehicleStatus_Success(postServAct, orgService);
        }

        private void UpdateVehicleStatus_Success(Entity ServAct, IOrganizationService orgService)
        {
            if (ServAct.Contains("new_vehicleid") != false)
            {
                Entity Vehicle = orgService.Retrieve(C_VehicleEntityName, ((EntityReference)ServAct["new_vehicleid"]).Id,
                    new ColumnSet("new_maintenancestage"));

                if (Vehicle == null) return;

                if (Vehicle.Contains("new_maintenancestage"))
                {
                    OptionSetValue statusValue = (OptionSetValue)Vehicle["new_maintenancestage"];
                    //保养状态为"T-15"
                    if (statusValue.Value == 100000002)
                    {
                        Vehicle["new_maintenancestage"] = new OptionSetValue(100000003);//保养状态置为“邀约成功”
                        orgService.Update(Vehicle);
                    }
                }
            }
        }

        private void UpdateVehicleStatus_Cancel(Entity ServAct, IOrganizationService orgService)
        {
            if (ServAct.Contains("new_vehicleid") != false)
            {
                if (ServAct.Contains("new_serviceactivitystatus") != false)
                    if (((OptionSetValue)ServAct["new_serviceactivitystatus"]).Value == 100000002) //活动为取消状态
                    {
                        Entity Vehicle = orgService.Retrieve(C_VehicleEntityName, ((EntityReference)ServAct["new_vehicleid"]).Id,
                            new ColumnSet("new_maintenancestage"));

                        if (Vehicle == null) return;

                        if (Vehicle.Contains("new_maintenancestage"))
                        {
                            OptionSetValue statusValue = (OptionSetValue)Vehicle["new_maintenancestage"];
                            //保养状态为"-车辆进厂"
                            //if (statusValue.Value == 100000004)
                            //{
                            //    throw new InvalidPluginExecutionException("相关车辆处于进厂状态，不能取消服务活动！");
                            //}
                            ////保养状态为"-完工状态"
                            //if (statusValue.Value == 100000005)
                            //{
                            //    throw new InvalidPluginExecutionException("相关车辆已完成保养，不能取消服务活动！");
                            //}
                            ////保养状态为"-完成满意度回访"
                            //if (statusValue.Value == 100000006)
                            //{
                            //    throw new InvalidPluginExecutionException("相关车辆已完成满意度回访，不能取消服务活动！");
                            //}

                            Vehicle["new_maintenancestage"] = new OptionSetValue(100000007);//保养状态置为“取消保养”
                            orgService.Update(Vehicle);
                        }
                    }
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
