using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class AccountPostSave:IPlugin
    {
        private const string C_EntityName = "account";
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
                    OnCreate(context, orgService);
                }
                else if (context.MessageName == "Update")
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
        private void OnCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity ac = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet(true));

            if(ac.Contains("new_customertype"))
                if(((OptionSetValue)ac["new_customertype"]).Value==100000000)
            CheckValid(ac, orgService);

        }

        private void OnUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity ac = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId,
                new ColumnSet(true));

            if (ac.Contains("new_customertype"))
                if (((OptionSetValue)ac["new_customertype"]).Value == 100000000)
                    CheckValid(ac, orgService);

        }

        /// <summary>
        /// 检测客户是否有重复
        /// </summary>
        /// <param name="ac">客户实体,其中必须包含起止号码字段</param>
        /// <param name="orgService"></param>
        /// <returns>客户有重复时,返回false</returns>
        private bool CheckValid(Entity ac, IOrganizationService orgService)
        {
            QueryExpression queryTR = new QueryExpression
            {
                EntityName = ac.LogicalName,
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name","new_department"),

                Criteria = new FilterExpression
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = { ac["name"] }
                            },
                            (ac.Contains("new_department") ?
                            new ConditionExpression
                            {
                                AttributeName = "new_department",
                                Operator = ConditionOperator.Equal,
                                Values = { ac["new_department"] }
                            }
                            :
                            new ConditionExpression
                            {
                                AttributeName = "new_department",
                                Operator = ConditionOperator.Null
                            }
                            )
                        }
                }
            };

            EntityCollection ECReturnTR = orgService.RetrieveMultiple(queryTR);
            if (ECReturnTR.Entities.Count > 1)  //客户有重复
            {
                foreach (Entity entity in ECReturnTR.Entities)
                {
                    if (entity.Id.CompareTo(ac.Id) != 0)
                    {
                        throw new Exception("客户名称（" + entity["name"] + "）部门（" + (entity.Contains("new_department") ? entity["new_department"] : "") + "）已经存在,请检查并重新输入!");
                    }
                }
                return false;
            }

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
