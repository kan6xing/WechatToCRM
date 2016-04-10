
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
    /// 根据车辆信息的变化更新客户状态
    /// </summary>
    public class VehicleUpdateOwnerPlugin : IPlugin
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
            //Entity vehicle = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("new_vehicleowner"));
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            EntityReference preVehicleOwner = GetAmountValue(preImage);
            EntityReference postVehicleOwner = GetAmountValue(postImage);

            if (postImage.Contains("new_vehicleowner") != false)
            {
                if (preVehicleOwner != null && postVehicleOwner != null)
                {
                    if (preVehicleOwner.Id != postVehicleOwner.Id)   //车主变化
                    {
                        UpdateVehicleOwner(postImage, orgService);  //更新新车主状态
                        UpdatePreVehicleOwner(preImage, orgService);   //更新老车主状态
                    }
                }
                else if (preVehicleOwner == null && postVehicleOwner != null)
                {
                    UpdateVehicleOwner(postImage, orgService);  //更新新车主状态
                }
                else if (preVehicleOwner != null && postVehicleOwner == null)
                {
                    UpdatePreVehicleOwner(preImage, orgService);   //更新老车主状态
                }
            }

        }

        private EntityReference GetAmountValue(Entity vehicle)
        {
            if (vehicle.Contains("new_vehicleowner") == false)
            {
                return null;
            }
            else
            {
                return (EntityReference)vehicle["new_vehicleowner"];
            }
        }


        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity postVehicle = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_vehicleowner", "new_purchaseatyanjun"));

            if (postVehicle.Contains("new_vehicleowner") != false)
            {
                UpdateVehicleOwner(postVehicle, orgService);
            }

        }

        private void UpdateVehicleOwner(Entity postVehicle, IOrganizationService orgService)
        {
            Entity Acc = orgService.Retrieve("account", ((EntityReference)postVehicle["new_vehicleowner"]).Id,
                new ColumnSet("new_accountstatus"));

            if (Acc.Contains("new_accountstatus"))
            {
                OptionSetValue statusValue = (OptionSetValue)Acc["new_accountstatus"];
                //客户当前状态是"潜在客户"\"前车主"\"前车主+潜在客户"时,更新为车主
                if (statusValue.Value == 100000001 || statusValue.Value == 100000003 || statusValue.Value == 100000004)
                {
                    if (postVehicle.Contains("new_purchaseatyanjun"))
                    {
                        if ((bool)postVehicle["new_purchaseatyanjun"])
                            Acc["new_accountstatus"] = new OptionSetValue(100000000);
                        else
                            return;
                    }
                    else
                        return;
                }
            }
            else
            {
                if (postVehicle.Contains("new_purchaseatyanjun"))
                {
                    if ((bool)postVehicle["new_purchaseatyanjun"])
                        Acc["new_accountstatus"] = new OptionSetValue(100000000);
                    else
                        return;
                }
                else
                    return;
            }

            orgService.Update(Acc);
        }

        private void UpdatePreVehicleOwner(Entity preVehicle, IOrganizationService orgService)
        {
            EntityCollection vehicles;
            Entity Acc = orgService.Retrieve("account", ((EntityReference)preVehicle["new_vehicleowner"]).Id,
                new ColumnSet("new_accountstatus"));

            QueryExpression query = new QueryExpression
            {
                EntityName = C_EntityName,
                //ColumnSet = new ColumnSet("new_accountcode",
                //    "totalamount"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_vehicleowner",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference)preVehicle["new_vehicleowner"]).Id.ToString() }
                            },
                            new ConditionExpression
                            {    //在验骏购车
                                AttributeName = "new_purchaseatyanjun",
                                Operator = ConditionOperator.Equal,
                                Values = { true }
                            }
                        }
                }
            };
            vehicles = orgService.RetrieveMultiple(query);

            if (vehicles.Entities.Count == 0)   //原车主名下无车时才更新状态
            {
                if (IsHaveOppty(Acc, orgService))  //有商机为潜在客户
                {
                    if (IsPurchaseAtYanjun(preVehicle, orgService)) //在燕骏购车，为“前车主+潜在客户”
                        Acc["new_accountstatus"] = new OptionSetValue(100000004);
                    else                                //未在燕骏购车，为“潜在客户”
                        Acc["new_accountstatus"] = new OptionSetValue(100000001);
                }
                else
                { 
                    if (IsPurchaseAtYanjun(preVehicle, orgService)) //在燕骏购车，为“前车主”
                        Acc["new_accountstatus"] = new OptionSetValue(100000003);
                    else                                //未在燕骏购车，为“潜在客户”
                        Acc["new_accountstatus"] = null;
                }
            }

            orgService.Update(Acc);
        }

        private bool IsPurchaseAtYanjun(Entity preVehicle, IOrganizationService orgService)
        {
            if (preVehicle.Contains("new_purchaseatyanjun"))
            {
                if ((bool)preVehicle["new_purchaseatyanjun"])
                {
                    return true;
                }
                else
                    return false;
            }
            else
                return false;
        }

        private bool IsHaveOppty(Entity Acc, IOrganizationService orgService)
        {
            EntityCollection oppties;
            QueryExpression query = new QueryExpression
            {
                EntityName = "opportunity",
                //ColumnSet = new ColumnSet("new_accountcode",
                //    "totalamount"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "customerid",
                                Operator = ConditionOperator.Equal,
                                Values = { Acc.Id.ToString() }
                            },
                            new ConditionExpression
                            {    
                                AttributeName = "statecode",
                                Operator = ConditionOperator.Equal,
                                Values = { 0 }
                            }
                        }
                }
            };
            oppties = orgService.RetrieveMultiple(query);

            if (oppties.Entities.Count > 0)
                return true;
            else
                return false;
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
