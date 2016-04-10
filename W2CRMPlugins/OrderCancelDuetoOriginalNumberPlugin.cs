
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
    /// 取消原始订单号所指订单
    /// </summary>
    public class OrderCancelDuetoOriginalNumberPlugin : IPlugin
    {

        private const string C_EntityName = "salesorder";
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
                    //DoUpdate(context, orgService);
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

        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity postEntity = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_originalcontractnumber", "new_ordertype"));
            EntityCollection originalContracts;

            if (postEntity.Contains("new_originalcontractnumber") && postEntity.Contains("new_ordertype"))
            {
                if (((OptionSetValue)postEntity["new_ordertype"]).Value == 100000000)   //如果是购车订单则处理
                {
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = C_EntityName,
                        ColumnSet = new ColumnSet(true),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_contractnumber",
                                Operator = ConditionOperator.Equal,
                                Values = { postEntity["new_originalcontractnumber"].ToString() }
                            }
                        }
                        }
                    };
                    originalContracts = orgService.RetrieveMultiple(query);

                    if (originalContracts.Entities.Count > 0)   //存在原合同，则取消原合同
                    {
                        //SetStateRequest setStateRequest = new SetStateRequest()
                        //{
                        //    EntityMoniker = new EntityReference
                        //    {
                        //        Id = originalContracts[0].Id,
                        //        LogicalName = C_EntityName
                        //    },
                        //    State = new OptionSetValue(2),
                        //    Status = 
                        //};

                        Entity orderclose = new Entity("orderclose");
                        orderclose["salesorderid"] = new EntityReference { Id = originalContracts[0].Id, LogicalName = C_EntityName };

                        CancelSalesOrderRequest setStateRequest = new CancelSalesOrderRequest();
                        setStateRequest.OrderClose = orderclose;
                        setStateRequest.Status = new OptionSetValue(4);

                        orgService.Execute(setStateRequest);
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
