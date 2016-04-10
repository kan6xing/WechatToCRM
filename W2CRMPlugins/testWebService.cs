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
    public class testWebService : IPlugin
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
            //DI_Wechat.WebService.WechatService.URL = "http://test.ylxrm.com:80/serverSoap.php?WSDL";
            //string strReturn = DI_Wechat.WebService.WechatService.InvokeWebMethod("serviceMessageSendInterface",
            //    new object[] { "274", "pinvyp1416366518", "oHrm1jq8e4Br20gzgPNTsXQvD9mw" });

            www.jseasy.com.cn.crm2011.Wechat2CRMPlugin.Plugins.SendWechatService.CustomerPortClient sendWechat = new Wechat2CRMPlugin.Plugins.SendWechatService.CustomerPortClient();
            sendWechat.serviceMessageSendInterface("274", "pinvyp1416366518", "oHrm1jq8e4Br20gzgPNTsXQvD9mw");
 
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
