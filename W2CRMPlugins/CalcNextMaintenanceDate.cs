using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    public class CalcNextMaintenanceDate : IPlugin
    {
        private const string C_EntityName = "salesorder";
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
                if (context.MessageName == "Create")
                {
                    DoCreate(context, orgService);
                }

                if (context.MessageName == "Update")
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

            DateTime? preRepairDatetime = GetRepairDatetime(preImage);
            DateTime? postRepairDatetime = GetRepairDatetime(postImage);

            if (CheckForWip(postImage) == false)
            {
                return;
            }

            if (preRepairDatetime.HasValue == true) return;
            if (postRepairDatetime.HasValue == false) return;

            DateTime currrentRepairDate = (DateTime)postImage["new_repaireddatetime"];
            double? currentMileAge = null;
            if (postImage.Contains("new_intofactorymileage") == false)
            {
                currentMileAge = null;
            }
            else
            {
                currentMileAge = (double)postImage["new_intofactorymileage"];
            }

            if (postImage.Contains("new_vehicleid") == false) return;
            EntityReference vehicleRef = postImage["new_vehicleid"] as EntityReference;

            Entity vehicle = orgService.Retrieve(vehicleRef.LogicalName, vehicleRef.Id, new ColumnSet(true));
            double? maintenanceMileInterval = GetMaintenanceMileInterval(vehicle, orgService);
            double? lastMileAge = GetMileAgeFromVehicle(vehicle);
            DateTime? lastRepairTime = getRepairDateFromVehicle(vehicle);

            Entity updateEntity = new Entity("new_vehiclefiles");
            updateEntity.Id = vehicle.Id;

            DateTime nextMaintenanceDate = DateTime.Today.AddDays(180);
            DateTime T180 = nextMaintenanceDate;

            if (maintenanceMileInterval != null && maintenanceMileInterval.HasValue == true &&
                currentMileAge != null && currentMileAge.HasValue == true)//&&
            //lastMileAge != null && lastMileAge.HasValue == true &&
            //lastRepairTime != null && lastRepairTime.HasValue == true
            {
                if (postImage.Contains("description"))          //判断是否为保养或首保WIP
                    if (postImage["description"].ToString().Contains("保养") ||
                        postImage["description"].ToString().Contains("首保"))
                    {
                        if (lastMileAge != null && lastMileAge.HasValue == true &&
                            lastRepairTime != null && lastRepairTime.HasValue == true)  //存在上一次保养时间和里程时,正常计算
                        {
                            if (currentMileAge.Value != lastMileAge.Value)
                            {
                                int temp = Convert.ToInt32(maintenanceMileInterval.Value * currrentRepairDate.Subtract(lastRepairTime.Value).Days / (currentMileAge.Value - lastMileAge.Value));
                                nextMaintenanceDate = currrentRepairDate.AddDays(temp);

                                if (nextMaintenanceDate > T180) nextMaintenanceDate = T180;
                            }
                        }
                        else //不存在上一次保养时间和里程时,直接设为180天后
                        {
                            nextMaintenanceDate = T180;
                        }

                        //更新车辆状态字段
                        updateEntity["new_lastdate"] = currrentRepairDate;
                        if (currentMileAge.HasValue)
                        {
                            updateEntity["new_lastkilometer"] = currentMileAge.Value;
                        }
                        updateEntity["new_estnextmaintenancedate"] = nextMaintenanceDate;
                    }
            }

            updateEntity["new_lastentereddate"] = currrentRepairDate;
            if (currentMileAge.HasValue)
            {
                updateEntity["new_lastenteredmileage"] = currentMileAge.Value;
            }

            orgService.Update(updateEntity);
        }

        private DateTime? GetRepairDatetime(Entity order)
        {
            if (order.Contains("new_repaireddatetime") == false)
            {
                return null;
            }
            else
            {
                return (DateTime)order["new_repaireddatetime"];
            }
        }

        private const int C_WIP_Order_Type = 100000001;
        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity order = (Entity)context.InputParameters["Target"];

            if (CheckForWip(order) == false) return;

            if (order.Contains("new_dmscreatedon") == false) return;  //modified： new_repaireddatetime
            DateTime currrentRepairDate = (DateTime)order["new_dmscreatedon"];  //modified： new_repaireddatetime

            double? currentMileAge = null;
            if (order.Contains("new_intofactorymileage") == false)
            {
                currentMileAge = null;
            }
            else
            {
                currentMileAge = (double)order["new_intofactorymileage"];
            }

            if (order.Contains("new_vehicleid") == false) return;
            EntityReference vehicleRef = order["new_vehicleid"] as EntityReference;

            Entity vehicle = orgService.Retrieve(vehicleRef.LogicalName, vehicleRef.Id, new ColumnSet(true));
            double? maintenanceMileInterval = GetMaintenanceMileInterval(vehicle, orgService);
            double? lastMileAge = GetMileAgeFromVehicle(vehicle);
            DateTime? lastRepairTime = getRepairDateFromVehicle(vehicle);

            Entity updateEntity = new Entity("new_vehiclefiles");
            updateEntity.Id = vehicle.Id;

            /*
            DateTime nextMaintenanceDate = DateTime.Today.AddDays(180);
            DateTime T180 = nextMaintenanceDate;
            if (maintenanceMileInterval != null && maintenanceMileInterval.HasValue == true &&
                currentMileAge != null && currentMileAge.HasValue == true &&
                lastMileAge != null && lastMileAge.HasValue == true &&
                lastRepairTime != null && lastRepairTime.HasValue == true)
            {
                if (currentMileAge.Value != lastMileAge.Value)
                {
                    if (order.Contains("description"))          //判断是否为保养或首保WIP
                        if (order["description"].ToString().Contains("保养") ||
                            order["description"].ToString().Contains("首保"))
                        {
                            int temp = Convert.ToInt32(maintenanceMileInterval.Value * currrentRepairDate.Subtract(lastRepairTime.Value).Days / (currentMileAge.Value - lastMileAge.Value));
                            nextMaintenanceDate = currrentRepairDate.AddDays(temp);

                            if (nextMaintenanceDate > T180) nextMaintenanceDate = T180;

                            updateEntity["new_lastdate"] = currrentRepairDate;
                            if (currentMileAge.HasValue)
                            {
                                updateEntity["new_lastkilometer"] = currentMileAge.Value;
                            }
                            updateEntity["new_estnextmaintenancedate"] = nextMaintenanceDate;
                        }
                }
            }*/

            updateEntity["new_lastentereddate"] = currrentRepairDate;
            if (currentMileAge.HasValue)
            {
                updateEntity["new_lastenteredmileage"] = currentMileAge.Value;
            }

            orgService.Update(updateEntity);
        }

        private DateTime? getRepairDateFromVehicle(Entity vehicle)
        {
            if (vehicle.Contains("new_lastdate") == false)
            {
                return null;
            }
            else
            {
                return (DateTime)vehicle["new_lastdate"];
            }
        }

        private double? GetMileAgeFromVehicle(Entity vehicle)
        {
            if (vehicle.Contains("new_lastkilometer") == false)
            {
                return null;
            }
            else
            {
                return (double)vehicle["new_lastkilometer"];
            }
        }

        private double? GetMaintenanceMileInterval(Entity vehicle, IOrganizationService orgService)
        {
            if (vehicle.Contains("new_brand") == false)
            {
                return null;
            }

            EntityReference brandRef = vehicle["new_brand"] as EntityReference;
            Entity brand = orgService.Retrieve(brandRef.LogicalName, brandRef.Id, new ColumnSet(true));
            if (brand.Contains("new_maintenanceinterval") == true)
            {
                return (double)brand["new_maintenanceinterval"];
            }

            return null;
        }

        private bool CheckForWip(Entity order)
        {
            if (order.Contains("new_ordertype") == false)
            {
                return false;
            }
            return (order["new_ordertype"] as OptionSetValue).Value == C_WIP_Order_Type;
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

            if (context.MessageName != "Update" && context.MessageName != "Create")
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
