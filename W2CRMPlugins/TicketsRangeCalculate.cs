using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class TicketsRangeCalculate : IPlugin
    {
        private const string C_EntityName = "new_ticketsrange";
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
                    OnCreate(context,orgService);
                }
                else if (context.MessageName == "Update")
                {
                    OnUpdate(context, orgService);
                }
                else if (context.MessageName == "Delete")
                {
                    OnDelete(context, orgService);
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
        private void OnCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity tr = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_salesorderid", "new_startnumber", "new_endnumber","new_startcode","new_endcode"));
            if (tr.Contains("new_salesorderid") == false)
            {
                return;
            }

            Entity so = orgService.Retrieve("salesorder", ((EntityReference)tr["new_salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error", "new_cinema", "new_ordertype"));

            if (((OptionSetValue)so["new_ordertype"]).Value != 100000002)   //判断订单类型不为C单时进行号段查重
            {
                CheckNumberValid(tr, orgService);
            }

            orderCheck.CheckRangeQuantity(so, orgService);
        }

        private void OnUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity tr = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("new_salesorderid", "new_startnumber", "new_endnumber", "new_startcode", "new_endcode","new_iscoupon"));
            if (tr.Contains("new_salesorderid") == false)
            {
                    throw new Exception("无关联订单，不能保存！");
            }
            if (tr.Contains("new_iscoupon") == false)
            {
                throw new Exception("无是否赠券标志，不能保存！");
            }
            else
            {
                if ((bool)tr["new_iscoupon"])
                {
                    Entity trTmp = new Entity(tr.LogicalName);
                    trTmp.Id = tr.Id;
                    trTmp["new_closeprice"] = 1;
                    trTmp["new_pricebase"] = 1;
                    orgService.Update(trTmp);
                }
            }
            Entity so = orgService.Retrieve("salesorder", ((EntityReference)tr["new_salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error", "new_cinema",
                    "new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate", 
                     "new_ordertype"));//"new_cactivateapprovestate", "new_corderfinapprovestate",

            if (((OptionSetValue)so["new_ordertype"]).Value != 100000002)   //判断订单类型不为C单时进行号段查重
            {
                CheckNumberValid(tr, orgService);
            }

            orderCheck.CheckRangeQuantity(so, orgService);

            if (!((Entity)context.InputParameters["Target"]).Contains("new_isactive"))   //不包含激活成功标志时,检查订单的提交状态。包含激活成功标志时，表示是激活模块在更新号段
                if (orderStatus.IsSubmit(so))
                    throw new Exception("订单已经提交，不能修改！");
        }

        private void OnDelete(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity tr = (Entity)context.PreEntityImages["Image"];// orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                 //new ColumnSet("new_salesorderid", "new_startnumber", "new_endnumber", "new_startcode", "new_endcode"));
            if (tr.Contains("new_salesorderid") == false)
            {
                return;
            }

            //CheckNumberValid(tr, orgService);

            Entity so = orgService.Retrieve("salesorder", ((EntityReference)tr["new_salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error", "new_cinema",
                    "new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate", 
                     "new_ordertype"));//"new_cactivateapprovestate", "new_corderfinapprovestate",

            orderCheck.CheckRangeQuantity(so, orgService);

            //if (orderStatus.IsSubmit(so))
            //    throw new Exception("订单已经提交，不能修改！");
        }

        /// <summary>
        /// 检测号段是否有重复
        /// </summary>
        /// <param name="tr">号段实体,其中必须包含起止号码字段</param>
        /// <param name="orgService"></param>
        /// <returns>号段有重复时,返回false</returns>
        private bool CheckNumberValid(Entity tr, IOrganizationService orgService)
        {
            QueryExpression queryTR = new QueryExpression
            {
                EntityName = tr.LogicalName,
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_salesorderid", "new_startcode", "new_endcode"),
                
                Criteria = new FilterExpression
                { 
                    FilterOperator =LogicalOperator.Or,
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_startnumber",
                                Operator = ConditionOperator.Between,
                                Values = { tr["new_startnumber"],tr["new_endnumber"] }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_endnumber",
                                Operator = ConditionOperator.Between,
                                Values = { tr["new_startnumber"],tr["new_endnumber"] }
                            }
                        }
                }
            };

            EntityCollection ECReturnTR = orgService.RetrieveMultiple(queryTR);
            if (ECReturnTR.Entities.Count > 1)  //号段有重复
            {
                foreach (Entity entity in ECReturnTR.Entities)
                {
                    if (entity.Id.CompareTo(tr.Id) != 0)
                    {
                        throw new Exception("号段已经存在（" + entity["new_startcode"] + "-" + entity["new_endcode"] + "）,请检查并重新输入!");
                    }
                }
                return false;
            }

            #region 号段包含在已有号段中间
            queryTR = new QueryExpression
            {
                EntityName = tr.LogicalName,
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_salesorderid", "new_startcode", "new_endcode"),

                Criteria = new FilterExpression
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_startnumber",
                                Operator = ConditionOperator.LessEqual,
                                Values = { tr["new_startnumber"]}
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_endnumber",
                                Operator = ConditionOperator.GreaterEqual,
                                Values = { tr["new_startnumber"] }
                            }
                        }
                }
            };

            ECReturnTR = orgService.RetrieveMultiple(queryTR);
            if (ECReturnTR.Entities.Count > 1)  //号段有重复
            {
                foreach (Entity entity in ECReturnTR.Entities)
                {
                    if (entity.Id.CompareTo(tr.Id) != 0)
                    {
                        throw new Exception("号段已经存在（" + entity["new_startcode"] + "-" + entity["new_endcode"] + "）,请检查并重新输入!");
                    }
                }
                return false;
            }
            #endregion

            return true;
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
