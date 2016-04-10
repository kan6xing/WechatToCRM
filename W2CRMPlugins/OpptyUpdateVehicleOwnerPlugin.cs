
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
    /// 根据商机的变化更新客户状态,创建商机时更新成潜在客户
    /// </summary>
    public class OpptyUpdateVehicleOwnerPlugin : IPlugin
    {

        private const string C_EntityName = "opportunity";
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
            Entity postOppty = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet("customerid", "statecode"));

            if (postOppty.Contains("customerid") != false && ((OptionSetValue) postOppty["statecode"]).Value==0)
            {
                UpdateVehicleOwner(postOppty, orgService);
            }

        }

        private void UpdateVehicleOwner(Entity postOppty, IOrganizationService orgService)
        {
            Entity Acc = orgService.Retrieve("account", ((EntityReference)postOppty["customerid"]).Id, 
                new ColumnSet("new_accountstatus"));

            if (Acc.Contains("new_accountstatus"))
            {
                OptionSetValue statusValue = (OptionSetValue)Acc["new_accountstatus"];
                //客户当前状态是"车主"时,更新为“二次购车”
                if (statusValue.Value == 100000000)
                {
                    Acc["new_accountstatus"] = new OptionSetValue(100000002);
                }

                //客户当前状态是"前车主"时,更新为“前车主+潜在客户”
                if (statusValue.Value == 100000003)
                {
                    Acc["new_accountstatus"] = new OptionSetValue(100000004);
                }

            }
            else     //客户状态为空时，置为潜在客户
                            Acc["new_accountstatus"] = new OptionSetValue(100000001);

            orgService.Update(Acc);
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
