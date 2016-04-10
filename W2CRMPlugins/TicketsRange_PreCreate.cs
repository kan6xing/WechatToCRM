using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class TicketsRange_PreCreate:IPlugin
    {
        private const string C_EntityName = "new_ticketsrange";
        private const string C_ImageName = "Image";

        OrderStatus os = new OrderStatus();
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
                    OnCreate(context, orgService);
                }
                if (context.MessageName == "Delete")
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
            Entity tr = (Entity)context.InputParameters["Target"];// orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));

            //检查是否存在已提交状态
            Entity so = orgService.Retrieve("salesorder", ((EntityReference)tr["new_salesorderid"]).Id,
                new ColumnSet("new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate"));
                    //,"new_cactivateapprovestate", "new_corderfinapprovestate"

            if (os.IsSubmit(so))
                throw new Exception("订单已经提交，不可创建");
        }

        private void OnDelete(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity tr = orgService.Retrieve(context.PrimaryEntityName, ((EntityReference)context.InputParameters["Target"]).Id, new ColumnSet(true));

            Entity so = orgService.Retrieve("salesorder", ((EntityReference)tr["new_salesorderid"]).Id,
                new ColumnSet("new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate"));
                    //,"new_cactivateapprovestate", "new_corderfinapprovestate"

            if (os.IsSubmit(so))
                throw new Exception("订单已经提交，不可删除");
        }
        
        private bool ValidInput(IPluginExecutionContext context)
        {
            if (context.PrimaryEntityName != C_EntityName)
            {
                return false;
            }

            if (context.Stage != 20)
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
