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
    public class CampaignActivityUpdate : IPlugin
    {
        private const string C_EntityName = "campaignactivity";
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
                //if (context.MessageName == "Create")
                //{
                //    OnUpdate(context, orgService);
                //}
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
            Entity entityImage,entityCusGroup;

            Entity compaignactivity = orgService.Retrieve(C_EntityName, context.PrimaryEntityId,
                new ColumnSet("new_imgid", "new_cusgroupid", "new_sendstate", "new_wxaccountid"));

            if (compaignactivity != null)
            {
                EntityReference  erImage,  erCusGroup;

                if (compaignactivity.Contains("new_sendstate"))
                {
                    if (((OptionSetValue)compaignactivity["new_sendstate"]).Value != 100000001)
                        return;
                }
                else
                {
                    return;
                }

                #region 处理图文

                //赋值图文
                if (compaignactivity.Contains("new_imgid"))
                {
                    erImage = (EntityReference)compaignactivity["new_imgid"];
                    entityImage = orgService.Retrieve("new_imagetextinfo", erImage.Id,
                        new ColumnSet("new_sourceid"));

                    if (entityImage.Contains("new_sourceid"))
                    {
                        compaignactivity["new_imgsourceid"] = entityImage["new_sourceid"];
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
                if (compaignactivity.Contains("new_cusgroupid"))
                {
                    erCusGroup = (EntityReference)compaignactivity["new_cusgroupid"];
                    entityCusGroup = orgService.Retrieve("new_cusgroup", erCusGroup.Id,
                        new ColumnSet("new_token", "new_sourceid"));

                    if (entityCusGroup.Contains("new_token") && entityCusGroup.Contains("new_sourceid"))
                    {
                        compaignactivity["new_token"] = entityCusGroup["new_token"];
                        compaignactivity["new_cusgroupsourceid"] = entityCusGroup["new_sourceid"];
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

                //compaignactivity["new_sendstate"] = new OptionSetValue(100000001);
                orgService.Update(compaignactivity);

                //#region 向微信中间表写入待发送任务日志

                //System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding();
                //Service1Client client = new Service1Client(binding, new System.ServiceModel.EndpointAddress("http://115.28.81.79:7001/mds/Service1.svc"));

                //string strResult = client.CreateLog("lettergroup", "3", compaignactivity.Id.ToString(), "", "True", compaignactivity["new_token"].ToString());

                //#endregion

            }

            //orderCheck.CheckValidPrice(so, orgService);
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
