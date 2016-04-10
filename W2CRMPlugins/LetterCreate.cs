using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;

using Microsoft.Crm.Sdk.Messages;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class LetterCreate : IPlugin
    {
        private const string C_EntityName = "letter";
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
            Entity entityCompaignActivity,entityImage,entityAccount,entityWXAccount;
            
            Entity letter = orgService.Retrieve(C_EntityName, context.PrimaryEntityId,
                new ColumnSet("regardingobjectid", "new_imgid", "new_token", "to", "new_openid", "new_imgsourceid"));

            if (letter != null)
            {
                EntityReference erRegarding,erImage,erAccount,erWXAccount ;

                #region 处理图文
                if (letter.Contains("regardingobjectid"))
                {
                    erRegarding = (EntityReference)letter["regardingobjectid"];
                }
                else
                {
                    return;
                }

                if (erRegarding.LogicalName == "campaignactivity")
                {
                    entityCompaignActivity = orgService.Retrieve("campaignactivity", erRegarding.Id,
                        new ColumnSet("new_imgid","new_wxaccountid"));

                    //赋值图文
                    if (entityCompaignActivity.Contains("new_imgid"))
                    {
                        erImage = (EntityReference)entityCompaignActivity["new_imgid"];
                        entityImage = orgService.Retrieve("new_imagetextinfo", erImage.Id,
                            new ColumnSet("new_sourceid"));

                        if (entityImage.Contains("new_sourceid"))
                        {
                            letter["new_imgsourceid"] = entityImage["new_sourceid"];
                            letter["new_imgid"] = erImage;
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }

                    //赋值token
                    if (entityCompaignActivity.Contains("new_wxaccountid"))
                    {
                        erWXAccount = (EntityReference)entityCompaignActivity["new_wxaccountid"];
                        entityWXAccount = orgService.Retrieve("new_wechataccount", erWXAccount.Id,
                            new ColumnSet("new_token"));

                        if (entityWXAccount.Contains("new_token"))
                        {
                            letter["new_token"] = entityWXAccount["new_token"];
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
                #endregion

                #region 处理客户
                if (letter.Contains("to"))
                {
                    //Entity toActivityParty = new Entity("activityparty");
                    //toActivityParty["partyid"] = new EntityReference(_accountType, _accountId);
                    //EntityCollection to = new EntityCollection(new List<Entity>() { toActivityParty });
                    erAccount = (EntityReference)((EntityCollection)letter["to"]).Entities[0]["partyid"];

                    entityAccount = GetAccountWechatAccount(letter["new_token"].ToString(), erAccount,
                        orgService);


                    if (entityAccount.Contains("new_openid") && entityAccount.Contains("new_token"))
                    {
                        letter["new_openid"] = entityAccount["new_openid"];
                        //letter["new_token"] = entityAccount["new_token"];
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

                #endregion
                //throw new Exception("token:" + letter["new_token"].ToString());
                //letter["new_sendstate"] = new OptionSetValue(100000001);
                orgService.Update(letter);

                //#region 向微信中间表写入待发送任务日志

                //System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding();
                //Service1Client client = new Service1Client(binding, new System.ServiceModel.EndpointAddress(@"http://115.28.81.79:7001/mds/Service1.svc"));
                ////throw new Exception(letter["new_token"].ToString() + ":" + entityAccount.Id);
                
                //string strResult = client.CreateLog("letter","3",letter.Id.ToString(),"","True",letter["new_token"].ToString());

                //#endregion

            }

            //orderCheck.CheckValidPrice(so, orgService);
        }

        private Entity GetAccountWechatAccount(string token, EntityReference account, IOrganizationService orgService)
        {
            QueryExpression queryAW = new QueryExpression
{
    EntityName = "new_accountwechataccount",
    //ColumnSet = new ColumnSet(true),
    ColumnSet = new ColumnSet("new_openid", "new_token"),
    Criteria = new FilterExpression
    {
        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_accountid",
                                Operator = ConditionOperator.Equal,
                                Values = { account.Id }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token  }
                            }
                        }
    }
};

            EntityCollection ECReturnAW = orgService.RetrieveMultiple(queryAW);

            if (ECReturnAW.Entities.Count > 0)
            {
                return ECReturnAW.Entities[0];
            }
            else
                return null;
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
