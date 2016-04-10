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
    /// 根据发票变化计算客户贡献及车辆贡献
    /// </summary>
    public class InvoiceCalcConsecrationPlugin : IPlugin
    {

        private const string C_EntityName = "invoice";
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
            //Entity invoice = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet("customerid"));
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            decimal preAmount = GetAmountValue(preImage);
            decimal postAmount = GetAmountValue(postImage);

            Guid preCustomerID = GetCustomerID(preImage);
            Guid postCustomerID = GetCustomerID(postImage);

            Guid preVehicleID = GetVehicleID(preImage);
            Guid postVehicleID = GetVehicleID(postImage);

            if (preCustomerID == postCustomerID)
            {
                if (postImage.Contains("customerid") != false)//To calculate account consecration
                {
                    if (preAmount != postAmount)
                    {
                        CalculateAccConsecration(postImage, orgService);
                    }
                }
            }
            else
            {
                if (preImage.Contains("customerid") != false)//To calculate account consecration
                {
                        CalculateAccConsecration(preImage, orgService);
                }
                if (postImage.Contains("customerid") != false)//To calculate account consecration
                {
                        CalculateAccConsecration(postImage, orgService);
                }
            }

            if (preVehicleID == postVehicleID)
            {
                if (postImage.Contains("new_vehicle") != false)//To calculate vehicle consecration
                {
                    if (preAmount != postAmount)
                    {
                        CalculateVehicleConsecration(postImage, orgService);
                    }
                }
            }
            else
            {
                if (postImage.Contains("new_vehicle") != false)//To calculate vehicle consecration
                {
                    CalculateVehicleConsecration(postImage, orgService);
                }

                if (preImage.Contains("new_vehicle") != false)//To calculate vehicle consecration
                {
                    CalculateVehicleConsecration(preImage, orgService);
                }
            }
            //To be add Calculate Vehicle Consecration
        }

        private Guid GetVehicleID(Entity invoice)
        {
            if (invoice.Contains("new_vehicle") == false)
            {
                return new Guid();
            }
            else
            {
                return ((EntityReference)invoice["new_vehicle"]).Id;
            }
        }

        private Guid GetCustomerID(Entity invoice)
        {
            if (invoice.Contains("customerid") == false)
            {
                return new Guid();
            }
            else
            {
                return ((EntityReference)invoice["customerid"]).Id;
            }
        }

        private decimal GetAmountValue(Entity invoice)
        {
            if (invoice.Contains("new_netprice") == false)
            {
                return 0;
            }
            else
            {
                return ((Money)invoice["new_netprice"]).Value;
            }
        }

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity invoice = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, 
                new ColumnSet("customerid", "new_vehicle"));

            if (invoice.Contains("customerid") != false)//To calculate account consecration
            {
                CalculateAccConsecration(invoice, orgService);
            }

            if (invoice.Contains("new_vehicle") != false)//To calculate vehicle consecration
            {
                CalculateVehicleConsecration(invoice, orgService);
            }

            //To be add Calculate Vehicle Consecration
        }

        private void CalculateAccConsecration(Entity invoice,IOrganizationService orgService)
        {
            Entity Acc = orgService.Retrieve("account", ((EntityReference)invoice["customerid"]).Id,
                new ColumnSet("new_customercontributed"));
            EntityCollection invoices;
            Entity inv;
            decimal consecration=0;
            QueryExpression query = new QueryExpression
            {
                EntityName = C_EntityName,
                ColumnSet = new ColumnSet("new_accountcode",
                    "new_netprice"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "customerid",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference)invoice["customerid"]).Id.ToString() }
                            }
                        }
                }
            };
            invoices = orgService.RetrieveMultiple(query);

            for (int i = 0; i < invoices.Entities.Count; i++)
            {
                inv = invoices[i];
                if (inv.Contains("new_accountcode"))
                {
                    Entity accountCode = orgService.Retrieve("new_dmsaccountsetting", ((EntityReference)inv["new_accountcode"]).Id, new ColumnSet("new_code"));
                    if (accountCode.Contains("new_code"))
                    {
                        if (accountCode["new_code"].ToString().Substring(0, 1) == "C" ||
                            accountCode["new_code"].ToString().Substring(0, 1) == "K")
                        {
                            if (inv.Contains("new_netprice"))
                                consecration = consecration + ((Money)inv["new_netprice"]).Value;
                        }
                    }
                }
            }

            if (Acc.Contains("new_customercontributed"))
            {
                if (consecration != ((Money)Acc["new_customercontributed"]).Value)
                {
                    Acc["new_customercontributed"] = new Money(consecration);
                    orgService.Update(Acc);
                }
            }
            else
            {
                if (consecration != 0)
                {
                    Acc["new_customercontributed"] = new Money(consecration);
                    orgService.Update(Acc);
                }
            }

        }

        private void CalculateVehicleConsecration(Entity invoice, IOrganizationService orgService)
        {
            Entity Vehicle = orgService.Retrieve("new_vehiclefiles", ((EntityReference)invoice["new_vehicle"]).Id,
                new ColumnSet("new_cumulativesales"));
            EntityCollection invoices;
            Entity inv;
            decimal consecration = 0;
            QueryExpression query = new QueryExpression
            {
                EntityName = C_EntityName,
                ColumnSet = new ColumnSet("new_accountcode",
                    "new_netprice"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_vehicle",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference)invoice["new_vehicle"]).Id.ToString() }
                            }
                        }
                }
            };
            invoices = orgService.RetrieveMultiple(query);

            for (int i = 0; i < invoices.Entities.Count; i++)
            {
                inv = invoices[i];
                if (inv.Contains("new_accountcode"))
                {
                    Entity accountCode = orgService.Retrieve("new_dmsaccountsetting", ((EntityReference)inv["new_accountcode"]).Id, new ColumnSet("new_code"));
                    if (accountCode.Contains("new_code"))
                    {
                        if (accountCode["new_code"].ToString().Substring(0, 1) == "C" ||
                            accountCode["new_code"].ToString().Substring(0, 1) == "K")
                        {
                            if (inv.Contains("new_netprice"))
                                consecration = consecration + ((Money)inv["new_netprice"]).Value;
                        }
                    }
                }
            }

            if (Vehicle.Contains("new_cumulativesales"))
            {
                if (consecration != ((Money)Vehicle["new_cumulativesales"]).Value)
                {
                    Vehicle["new_cumulativesales"] = new Money(consecration);
                    orgService.Update(Vehicle);
                }
            }
            else
            {
                if (consecration != 0)
                {
                    Vehicle["new_cumulativesales"] = new Money(consecration);
                    orgService.Update(Vehicle);
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
