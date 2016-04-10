using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class ActionCreate : IPlugin
    {
        private const string C_EntityName = "new_sharehistory";
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

                if (context.MessageName == "Update")
                {
                    OnUpdate(context, orgService);
                }
                if (context.MessageName == "Create")
                {
                    OnUpdate(context, orgService);
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
            //
            Entity action = orgService.Retrieve(C_EntityName, context.PrimaryEntityId,
                new ColumnSet("new_actionobject", "new_type", "new_imagetextinfoid", "new_productid", "new_accountid", "new_wechataccountid"));

            if (action != null)
            {

                switch (((OptionSetValue)action["new_actionobject"]).Value.ToString())
                {
                    case "100000000":    //图文信息
                        switch (((OptionSetValue)action["new_type"]).Value.ToString())
                        {
                            case "100000001":   //分享
                            case "100000002":   //收藏
                            case "100000004":   //阅读
                                UpdateAttentionResultByRelation(action, orgService);
                                break;
                        }
                        break;
                    case "100000002":    //商品
                        switch (((OptionSetValue)action["new_type"]).Value.ToString())
                        {
                            case "100000000":   //打开
                            case "100000001":   //分享
                            case "100000002":   //收藏
                            case "100000003":   //加入购物车
                            case "100000005":   //购买
                                UpdateAttentionResult(action, orgService);
                                break;
                        }
                        break;
                }
            }

            //orderCheck.CheckValidPrice(so, orgService);
        }

        private void UpdateAttentionResult(Entity action, IOrganizationService orgService)
        {
            Entity ar;

            action["new_attentionvalue"] = CalculateValue((EntityReference)action["new_wechataccountid"], (EntityReference)action["new_productid"], ((OptionSetValue)action["new_type"]).Value.ToString(), orgService);

            QueryExpression queryAR = new QueryExpression
            {
                EntityName = "new_attentionresult",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_productid", "new_wxaccountid", "new_accountid",
                    "new_value","new_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference)action["new_productid"]).Id }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_accountid",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference)action["new_accountid"]).Id  }
                            }
                        }
                }
            };

            EntityCollection ECReturnAR = orgService.RetrieveMultiple(queryAR);

            if (ECReturnAR.Entities.Count > 0)
            {
                ar = ECReturnAR.Entities[0];
                ar["new_value"] = Convert.ToInt32(Convert.ToInt32(ar["new_value"]) + Convert.ToInt32(action["new_attentionvalue"]));
                orgService.Update(ar);
            }
            else
            {
                ar = new Entity("new_attentionresult");
                ar["new_productid"] = action["new_productid"];
                ar["new_wxaccountid"] = action["new_wechataccountid"];
                ar["new_accountid"] = action["new_accountid"];
                ar["new_name"] = ((EntityReference)action["new_accountid"]).Name + "->" + ((EntityReference)action["new_productid"]).Name;
                ar["new_value"] = action["new_attentionvalue"];

                ar.Id= orgService.Create(ar);
                //throw new Exception("!!!!!!!!!!");
            }

            CreateLead(ar, orgService);

            action["new_attentionresultid"]= ar.ToEntityReference();
            orgService.Update(action);
        }

        private void UpdateAttentionResultByRelation(Entity action, IOrganizationService orgService)
        {
            QueryExpression queryRelation = new QueryExpression
            {
                EntityName = "new_prd2img",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_productid",  "new_imgid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_imgid",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference) action["new_imagetextinfoid"]).Id }
                            }
                        }
                }
            };

            EntityCollection ECReturnRelation = orgService.RetrieveMultiple(queryRelation);

            if (ECReturnRelation.Entities.Count > 0)
            {
                foreach (Entity et in ECReturnRelation.Entities)
                {
                    action["new_productid"]=(EntityReference)et["new_productid"];
                    UpdateAttentionResult( action, orgService);

                }
            }
        }

        private void CreateLead(Entity AttentionResult, IOrganizationService orgService)
        {
            int iKey=0;

            QueryExpression queryLC = new QueryExpression
            {
                EntityName = "new_new_leadscontrol",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_productid", "new_lowvalue"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Equal,
                                Values = { ((EntityReference) AttentionResult["new_productid"]).Id }
                            }
                        }
                }
            };

            EntityCollection ECReturnLC = orgService.RetrieveMultiple(queryLC);

            if (ECReturnLC.Entities.Count > 0)
            {
                iKey = (int)ECReturnLC.Entities[0]["new_lowvalue"];
            }

            if (iKey == 0)
            {
                queryLC = new QueryExpression
                {
                    EntityName = "new_new_leadscontrol",
                    //ColumnSet = new ColumnSet(true),
                    ColumnSet = new ColumnSet("new_productid", "new_lowvalue"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Null 
                            }
                        }
                    }
                };

                ECReturnLC = orgService.RetrieveMultiple(queryLC);

                if (ECReturnLC.Entities.Count > 0)
                {
                    iKey = (int)ECReturnLC.Entities[0]["new_lowvalue"];
                }
            }

            if (iKey == 0)
            {
                return;
            }
            else
            {
                if (Convert.ToInt32(AttentionResult["new_value"]) >= iKey)
                {    // Create Lead
                    Entity lead = new Entity("lead");
                    lead["new_productid"] = AttentionResult["new_productid"];
                    lead["new_wxaccountid"] = AttentionResult["new_wxaccountid"];
                    lead["new_attentionresultid"] = AttentionResult.ToEntityReference();
                    lead["customerid"] = AttentionResult["new_accountid"];
                    lead["lastname"] = ((EntityReference)AttentionResult["new_accountid"]).Name;
                    lead["subject"] = AttentionResult["new_name"];
                    lead["leadsourcecode"] = new OptionSetValue(100000000);

                    orgService.Create(lead);
                }
            }

        }

        private int CalculateValue(EntityReference token, EntityReference product, string action, IOrganizationService orgService)
        {
            int iValue = 0;

            QueryExpression queryAP = new QueryExpression
{
    EntityName = "new_attentionpara",
    //ColumnSet = new ColumnSet(true),
    ColumnSet = new ColumnSet("new_productid", "new_wxaccountid", "new_cartvalue", 
        "new_clickvalue","new_collectvalue","new_ordervalue","new_readvalue","new_sharevalue"),
    Criteria = new FilterExpression
    {
        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Equal,
                                Values = { product.Id }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_wxaccountid",
                                Operator = ConditionOperator.Equal,
                                Values = { token.Id }
                            }
                        }
    }
};

            EntityCollection ECReturnAP = orgService.RetrieveMultiple(queryAP);

            if (ECReturnAP.Entities.Count == 0)
            {
                queryAP = new QueryExpression
                {
                    EntityName = "new_attentionpara",
                    //ColumnSet = new ColumnSet(true),
                    ColumnSet = new ColumnSet("new_productid", "new_wxaccountid", "new_cartvalue",
                        "new_clickvalue", "new_collectvalue", "new_ordervalue", "new_readvalue", "new_sharevalue"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Null
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_wxaccountid",
                                Operator = ConditionOperator.Equal,
                                Values = { token.Id }
                            }
                        }
                    }
                };

                ECReturnAP = orgService.RetrieveMultiple(queryAP);
            }

            if (ECReturnAP.Entities.Count > 0)
            {
                switch (action)
                {
                    case "100000000":   //打开
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_clickvalue"]);
                        break;
                    case "100000001":   //分享
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_sharevalue"]);
                        break;
                    case "100000002":   //收藏
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_collectvalue"]);
                        break;
                    case "100000003":   //加入购物车
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_cartvalue"]);
                        break;
                    case "100000004":   //阅读
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_readvalue"]);
                        break;
                    case "100000005":   //购买
                        iValue = Convert.ToInt32(ECReturnAP[0]["new_ordervalue"]);
                        break;
                }
            }

            return iValue;
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
