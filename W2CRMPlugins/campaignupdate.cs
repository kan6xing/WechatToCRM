using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.GoldenHarvest.Plugins
{
    public class CampaignUpdate : IPlugin
    {
        private const string C_EntityName = "campaign";
        private const string C_ImageName = "Image";

        OrderCheck orderCheck = new OrderCheck();
        OrderStatus orderStatus = new OrderStatus();

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
                    //OnCreate(context,orgService);
                }
                else if (context.MessageName == "Update")
                {
                    OnUpdate(context, orgService);
                }
                else if (context.MessageName == "Delete")
                {
                    //OnDelete(context, orgService);
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

        private void OnUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity campaign = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_salesorderid", "new_startnumber", "new_endnumber", "new_startcode", "new_endcode","new_iscoupon"));
            if (campaign.Contains("new_salesorderid") == false)
            {
                    throw new Exception("无关联订单，不能保存！");
            }
            if (campaign.Contains("new_iscoupon") == false)
            {
                throw new Exception("无是否赠券标志，不能保存！");
            }
            else
            {
                if ((bool)campaign["new_iscoupon"])
                {
                    Entity trTmp = new Entity(campaign.LogicalName);
                    trTmp.Id = campaign.Id;
                    trTmp["new_closeprice"] = 1;
                    trTmp["new_pricebase"] = 1;
                    orgService.Update(trTmp);
                }
            }
            Entity so = orgService.Retrieve("salesorder", ((EntityReference)campaign["new_salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error", "new_cinema",
                    "new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate", 
                     "new_ordertype"));//"new_cactivateapprovestate", "new_corderfinapprovestate",

            if (((OptionSetValue)so["new_ordertype"]).Value != 100000002)   //判断订单类型不为C单时进行号段查重
            {
                CheckNumberValid(campaign, orgService);
            }

            orderCheck.CheckRangeQuantity(so, orgService);

            if (!((Entity)context.InputParameters["Target"]).Contains("new_isactive"))   //不包含激活成功标志时,检查订单的提交状态。包含激活成功标志时，表示是激活模块在更新号段
                if (orderStatus.IsSubmit(so))
                    throw new Exception("订单已经提交，不能修改！");
        }


        /// <summary>
        /// 目标市场营销列表
        /// </summary>
        /// <param name="CampaignID"></param>
        /// <param name="orgService"></param>
        /// <returns>返回市场营销列表集</returns>
        private EntityCollection GetList(Guid CampaignID, IOrganizationService orgService)
        {
            QueryExpression queryCampaign = new QueryExpression()
            {
                EntityName = "list",
                ColumnSet = new ColumnSet("listname"),
                LinkEntities = 
                        {
                            new LinkEntity
                            {
                                LinkFromEntityName = "list",
                                LinkFromAttributeName = "listid",
                                LinkToEntityName = "campaignitem",
                                LinkToAttributeName = "entityid",
                                LinkCriteria = new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions = 
                                    {
                                        new ConditionExpression
                                        {
                                            AttributeName = "campaignid",
                                            Operator = ConditionOperator.Equal,
                                            Values = { new Guid("793833FE-91F5-E211-9FA6-08002739A898") }
                                        }
                                    }
                                }
                            }
                        }
            };

            // Obtain results from the query expression.
            EntityCollection ec = orgService.RetrieveMultiple(queryCampaign);


            return ec;
        }

        /// <summary>
        /// 获取市场营销列表的成员
        /// </summary>
        /// <param name="ListID"></param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        private EntityCollection GetListAccounts(Guid ListID, IOrganizationService orgService)
        {
            QueryExpression query = new QueryExpression()
            {
                EntityName = "account",
                ColumnSet = new ColumnSet("name"),
                LinkEntities = 
                        {
                            new LinkEntity
                            {
                                LinkFromEntityName = "account",
                                LinkFromAttributeName = "accountid",
                                LinkToEntityName = "listmember",
                                LinkToAttributeName = "entityid",
                                LinkCriteria = new FilterExpression
                                {
                                    FilterOperator = LogicalOperator.And,
                                    Conditions = 
                                    {
                                        new ConditionExpression
                                        {
                                            AttributeName = "listid",
                                            Operator = ConditionOperator.Equal,
                                            Values = { ListID }
                                        }
                                    }
                                }
                            }
                        }
            };

            // Obtain results from the query expression.
            EntityCollection ecAccount = orgService.RetrieveMultiple(query);
            return ecAccount;
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
                context.MessageName != "Create" &&
                context.MessageName != "Delete")
            {
                return false;
            }

            //if (context.MessageName == "Update")
            //{
            //    if (context.PreEntityImages.ContainsKey(C_ImageName) == false ||
            //   context.PostEntityImages.ContainsKey(C_ImageName) == false)
            //    {
            //        return false;
            //    }
            //}

            return true;
        }
    }
}
