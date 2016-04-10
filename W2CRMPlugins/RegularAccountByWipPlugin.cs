using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.yanjun.AutoTask.Plugins
{
    //本Plugin被移到自动任务代码中，当前代码被放弃。 By Daemon Lin On 2011-10-16 17:45
    internal class RegularAccountByWipPlugin : IPlugin
    {
        private const string C_EntityName = "invoice";
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
                    DoUpdate(context, orgService);
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

        private void DoUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity preImage = context.PreEntityImages[C_ImageName];
            Entity postImage = context.PostEntityImages[C_ImageName];

            EnumInvoiceType? preInvoiceType = GetInvoiceType(preImage);
            EnumInvoiceType? postInvoiceType = GetInvoiceType(postImage);

            EntityReference preAccount = GetAccountReference(preImage);
            EntityReference postAccount = GetAccountReference(postImage);

            DateTime? preInvoiceDate = GetInvoiceDate(preImage);
            DateTime? postInvoiceDate = GetInvoiceDate(postImage);

            if (preInvoiceType.Value == EnumInvoiceType.NewBuy && postInvoiceType.Value == EnumInvoiceType.NewBuy)
            {
                return;
            }

            if (preInvoiceType != postInvoiceType ||
                preAccount != postAccount ||
                preInvoiceDate != postInvoiceDate)
            {
                if (preAccount.Id == postAccount.Id)
                {
                    UpdateAccountRegularFlag(orgService, postAccount, IsRegularAccount(postAccount, orgService));
                }
                else
                {
                    UpdateAccountRegularFlag(orgService, preAccount, IsRegularAccount(preAccount, orgService));
                    UpdateAccountRegularFlag(orgService, postAccount, IsRegularAccount(postAccount, orgService));
                }
            }
        }

        private DateTime? GetInvoiceDate(Entity invoice)
        {
            if (invoice.Contains("new_invoicedate") == false)
            {
                return null;
            }
            else
            {
                return (DateTime)invoice["new_invoicedate"];
            }
        }

        private EntityReference GetAccountReference(Entity invoice)
        {
            if (invoice.Contains("customerid") == false)
            {
                return null;
            }
            else
            {
                return (EntityReference)invoice["customerid"];
            }
        }

        private EnumInvoiceType? GetInvoiceType(Entity invoice)
        {
            if (invoice.Contains("new_invoicetype") == false)
            {
                return null;
            }
            else
            {
                return (EnumInvoiceType)((invoice["new_invoicetype"] as OptionSetValue).Value);
            }
        }

        private enum EnumInvoiceType
        {
            NewBuy = 100000000,
            Wip = 100000001
        }
        private void DoCreate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Guid invoiceId = context.PrimaryEntityId;
            Entity invoice = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
            if (invoice.Contains("new_invoicetype") == false)
            {
                return;
            }

            EnumInvoiceType invoiceType = (EnumInvoiceType)((invoice["new_invoicetype"] as OptionSetValue).Value);
            if (invoiceType == EnumInvoiceType.NewBuy)
            {
                return;
            }

            if (invoice.Contains("customerid") == false)
            {
                return;
            }
            EntityReference accRef = invoice["customerid"] as EntityReference;
            Entity acc = orgService.Retrieve(accRef.LogicalName, accRef.Id, new ColumnSet(true));
            if (acc.Contains("new_regularflag") == true && (bool)acc["new_regularflag"] == true)
            {
                return;
            }

            UpdateAccountRegularFlag(orgService, accRef, IsRegularAccount(accRef, orgService));
        }

        private static void UpdateAccountRegularFlag(IOrganizationService orgService, EntityReference accRef, bool regularFlag)
        {
            if (accRef == null)
            {
                return;
            }
            Entity updateEntity = new Entity(accRef.LogicalName);
            updateEntity.Id = accRef.Id;
            updateEntity["new_regularflag"] = regularFlag;

            orgService.Update(updateEntity);
        }

        private bool IsRegularAccount(EntityReference accRef, IOrganizationService orgService)
        {
            return GetLast15MonthesWipAmount(accRef, orgService) > 2;
        }

        private int GetLast15MonthesWipAmount(EntityReference accRef, IOrganizationService orgService)
        {
            QueryExpression query = new QueryExpression(C_EntityName);
            query.Criteria = new FilterExpression(LogicalOperator.And);
            query.Criteria.AddCondition("customerid", ConditionOperator.Equal, accRef.Id);
            query.Criteria.AddCondition("new_invoicetype", ConditionOperator.Equal, (int)EnumInvoiceType.Wip);
            query.Criteria.AddCondition("new_invoicedate", ConditionOperator.LastXMonths, 15);

            return orgService.RetrieveMultiple(query).Entities.Count;
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
