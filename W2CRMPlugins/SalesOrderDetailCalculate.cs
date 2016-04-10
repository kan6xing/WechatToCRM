using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class SalesOrderDetailCalculate : IPlugin
    {
        private const string C_EntityName = "salesorderdetail";
        private const string C_ImageName = "Image";

        OrderCheck orderCheck = new OrderCheck();
        PubClass pulicClass = new PubClass();
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
            Entity sod = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
            if (sod.Contains("priceperunit") == false)
            {
                return;
            }

            decimal dQuantity=(decimal)sod["quantity"];                  //数量
            Money mStandardPrice = ((Money)sod["priceperunit"]);                  //正价
            Money mDiscount = ((Money)sod["volumediscountamount"]);               //批发折扣
            Money mBasePrice =new Money( mStandardPrice.Value-mDiscount.Value);    //基础价
            Money mClosePrice;
            if (sod.Contains("new_closeprice") == false)
                mClosePrice = new Money(mBasePrice.Value);                        //售价
            else
                mClosePrice = (Money)sod["new_closeprice"];

            Entity sod_update = new Entity(context.PrimaryEntityName);
            sod_update.Id = context.PrimaryEntityId;
            sod_update["new_baseprice"] = mBasePrice;
            if (sod.Contains("new_closeprice") == false)    //售价为空时赋值，不用计算零售折扣
                sod_update["new_closeprice"] = mClosePrice;
            else                                            //售价不为空时，计算零售折扣
            {
                sod_update["manualdiscountamount"] = new Money((mBasePrice.Value - mClosePrice.Value) * dQuantity);
            }

            orgService.Update(sod_update);
            
            //检查号段与团票匹配性
            Entity so = orgService.Retrieve("salesorder", ((EntityReference)sod["salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error", "new_cinema", "new_ordertype"));

            orderCheck.CheckRangeQuantity(so, orgService);

            //orderCheck.CheckValidPrice(so, orgService);
        }

        private void OnUpdate(IPluginExecutionContext context, IOrganizationService orgService)
        {
            Entity sod = orgService.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));

            decimal dQuantity = (decimal)sod["quantity"];                  //数量
            Money mStandardPrice = ((Money)sod["priceperunit"]);                  //正价
            Money mDiscount = ((Money)sod["volumediscountamount"]);               //批发折扣

            Entity sod_update = new Entity(context.PrimaryEntityName);
            sod_update.Id = context.PrimaryEntityId;

                Money mBasePrice = new Money(mStandardPrice.Value - mDiscount.Value);    //基础价
                sod_update["new_baseprice"] = mBasePrice;

                Money mClosePrice;
            if (pulicClass.IsGroupTicketType(sod,orgService) == false)    //产品不为团票类时，给售价赋值
            {
                mClosePrice = new Money(mStandardPrice.Value - mDiscount.Value);          //售价
                sod_update["new_closeprice"] = mClosePrice;
            }
            else
            {
                mClosePrice = (Money)sod["new_closeprice"];
            }
            sod_update["manualdiscountamount"] = new Money((mBasePrice.Value - mClosePrice.Value) * dQuantity);


            orgService.Update(sod_update);

            //检查号段与团票匹配性
            Entity so = orgService.Retrieve("salesorder", ((EntityReference)sod["salesorderid"]).Id,
                new ColumnSet("new_meetthespecificationforgift", "new_error","new_cinema",
                    "new_bordermktapprovestate", "new_borderfinapprovestate", "new_aorderapprovestate",
                    "new_ordertype"));//"new_cactivateapprovestate", "new_corderfinapprovestate", 

            orderCheck.CheckRangeQuantity(so, orgService);

            //orderCheck.CheckValidPrice(so, orgService);

            if (orderStatus.IsSubmit(so))
                throw new Exception("订单已经提交，不能修改！");

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
