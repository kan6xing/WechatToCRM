using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;

namespace www.jseasy.com.cn.crm2011.Wechat2CRM.Plugins
{
    public class PubClass
    {
        public PubClass()
        {
            //if (DateTime.Today > Convert.ToDateTime("2014-9-30"))
            //    throw new InvalidPluginExecutionException("未将对象引用设置到对象的实例");
        }

        /// <summary>
        /// 订单产品是否为团票类产品
        /// </summary>
        /// <param name="sod">订单产品</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        public bool IsGroupTicketType(Entity sod, IOrganizationService orgService)
        {
            Entity product = orgService.Retrieve("product", ((EntityReference)sod["productid"]).Id, new ColumnSet(true));
            long strProductType = ((OptionSetValue)product["new_producttype"]).Value;

            if (strProductType == 100000000 ||
                strProductType == 100000001 ||
                strProductType == 100000002 ||
                strProductType == 100000003 ||
                strProductType == 100000004 ||
                strProductType == 100000005 ||
                strProductType == 100000006)
                return true;
            return false;
        }

        /// <summary>
        /// 订单产品是否为团票产品
        /// </summary>
        /// <param name="sod">订单产品</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        public bool IsGroupTicket(Entity sod, IOrganizationService orgService)
        {
            if (sod.Contains("productid"))
            {
                Entity product = orgService.Retrieve("product", ((EntityReference)sod["productid"]).Id, new ColumnSet(true));
                long strProductType = ((OptionSetValue)product["new_producttype"]).Value;

                if (strProductType == 100000000 ||
                    strProductType == 100000004)
                    return true;
                return false;
            }
            return false;
        }

        /// <summary>
        /// 获取产品类型
        /// </summary>
        /// <param name="ProductID">产品的GUID</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        public OptionSetValue GetProductType(Guid ProductID, IOrganizationService orgService)
        {
            Entity product = orgService.Retrieve("product", ProductID, new ColumnSet(true));
            if (product.Contains("new_producttype"))
                return (OptionSetValue)product["new_producttype"];
            else
                return null;
        }

    }

    public class OrderCheck
    {
        PubClass publicClass = new PubClass();
        public OrderCheck() 
        {
            if (DateTime.Today > Convert.ToDateTime("2014-9-30"))
                throw new InvalidPluginExecutionException("未将对象引用设置到对象的实例");
        }

        /// <summary>
        /// 检测号段与团票数量、赠券数量与赠券号段是否相符，赠券是否合规
        /// </summary>
        /// <param name="so">订单</param>
        /// <param name="orgService"></param>
        public void CheckRangeQuantity(Entity so, IOrganizationService orgService)
        {
            if (((OptionSetValue)so["new_ordertype"]).Value != 100000002)   //判断订单类型不为C单时进行号段验证   xhx change
            {
                Entity so_update = new Entity(so.LogicalName);
                so_update.Id = so.Id;
                so_update["new_error"] = "";
                so_update["new_cinema"] = so["new_cinema"];
                IsValidQuantity(so_update, orgService);
                IsValidGift(so_update, orgService);
                IsValidDates(so_update, orgService);
                CheckValidPrice(so_update, orgService);

                if (((bool)so_update["new_isuptoquantity"]) == false ||
                    ((bool)so_update["new_isuptodates"]) == false ||
                    ((bool)so_update["new_meetthespecificationforgift"]) == false ||
                    ((bool)so_update["new_isuptostandard"]) == false)
                    so_update["new_isuptoall"] = false;
                else
                    so_update["new_isuptoall"] = true;

                orgService.Update(so_update);
            }
        }
        
        /// <summary>
        /// 判断团票数量是否与订单产品的数量字段值相符
        /// </summary>
        /// <param name="SO">主订单实体</param>
        /// <returns></returns>
        private bool IsValidQuantity(Entity SO, IOrganizationService orgService)
        {
            decimal dQuantity = 0, dTRQuantity = 0;

            #region 查找订单产品

            EntityCollection ECReturnSOD = SearchSOD(new EntityReference("salesorder", SO.Id), orgService);

            foreach (Entity sod in ECReturnSOD.Entities)
            {
                if (publicClass.IsGroupTicket(sod, orgService))
                {
                    dQuantity += Convert.ToDecimal(sod["quantity"]);
                }
            }
            #endregion

            #region 查找团票号段子记录
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_ticketsrange",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_number",
                    "new_salesorderid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_salesorderid",
                                Operator = ConditionOperator.Equal,
                                Values = { SO.Id }//new EntityReference("salesorder", SO.Id)
                            },
                            new ConditionExpression
                            {
                                AttributeName="new_iscoupon",
                                Operator=ConditionOperator.Equal,
                                Values={ false }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);
            #endregion

            foreach (Entity tr in ECReturn.Entities)
            {
                dTRQuantity += Convert.ToDecimal(tr["new_number"]);    //号段数量求和
            }

            if (dQuantity == dTRQuantity)
            {
                SO["new_error"] = "";
                return true;
            }
            else
            {
                SO["new_error"] = "团票号段数量之和(" + dTRQuantity.ToString("##########") +
                    ")与订单产品数量(" + dQuantity.ToString("##########") + ")不符，请检查号段信息是否正确";
                return false;
            }

        }

        /// <summary>
        /// 判断有效期及数量是否合规
        /// </summary>
        /// <param name="SO">目标订单</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        private bool IsValidDates(Entity SO, IOrganizationService orgService)
        {
            #region 查找订单产品

            DateTime dtStart, dtEnd;

            Decimal dValidDates;

            EntityCollection ECReturnSOD = SearchSOD(new EntityReference("salesorder", SO.Id), orgService);

            SO["new_isuptodates"] = true;
            SO["new_isuptoquantity"] = true;

            foreach (Entity sod in ECReturnSOD.Entities)
            {
                if (publicClass.IsGroupTicket(sod, orgService))    //团体票
                {
                    if (sod.Contains("new_startdate") && sod.Contains("new_enddate"))
                    {
                        dtStart = (DateTime)sod["new_startdate"];
                        dtEnd = (DateTime)sod["new_enddate"];

                        int iDates = (dtEnd - dtStart).Days;
                        if (iDates > 184) SO["new_isuptodates"] = false;     //有效期不合规
                    }
                    else
                    {
                        SO["new_isuptodates"] = false;     //有效期不合规
                    }

                    if (Convert.ToDecimal(sod["quantity"]) < 30)   //数量不合规
                        SO["new_isuptoquantity"] = false;
                }

                if (publicClass.GetProductType(((EntityReference)sod["productid"]).Id, orgService).Value == 100000001)    //储值礼卡
                {
                    if (sod.Contains("new_startdate") && sod.Contains("new_enddate"))
                    {
                        dtStart = (DateTime)sod["new_startdate"];
                        dtEnd = (DateTime)sod["new_enddate"];

                        int iDates = (dtEnd - dtStart).Days;
                        if (iDates > 184) SO["new_isuptodates"] = false;     //有效期不合规
                    }
                    else
                    {
                        SO["new_isuptodates"] = false;     //有效期不合规
                    }

                    if (Convert.ToDecimal(sod["quantity"]) < 50)   //数量不合规
                        SO["new_isuptoquantity"] = false;
                }

                if (publicClass.GetProductType(((EntityReference)sod["productid"]).Id, orgService).Value == 100000011)    //赠券
                {
                    if (sod.Contains("new_startdate") && sod.Contains("new_enddate"))
                    {
                        dtStart = (DateTime)sod["new_startdate"];
                        dtEnd = (DateTime)sod["new_enddate"];

                        int iDates = (dtEnd - dtStart).Days;
                        if (iDates > 92) SO["new_isuptodates"] = false;     //有效期不合规
                    }
                    else
                    {
                        SO["new_isuptodates"] = false;     //有效期不合规
                    }
                }
                else
                {
                    SO["new_price"] = ((Money)sod["new_closeprice"]).Value;
                    SO["new_quantity"] = sod["quantity"];
                    SO["new_amount"] = ((Money)sod["extendedamount"]).Value;
                }

                dValidDates=GetValidDates(((EntityReference)SO["new_cinema"]).Id,((EntityReference)sod["productid"]).Id,
                    (Decimal)sod["quantity"], orgService);
                if (dValidDates > 0)
                {
                    if (sod.Contains("new_startdate") && sod.Contains("new_enddate"))
                    {
                        dtStart = (DateTime)sod["new_startdate"];
                        dtEnd = (DateTime)sod["new_enddate"];

                        int iDates = (dtEnd - dtStart).Days;
                        if (iDates > dValidDates) SO["new_isuptodates"] = false;     //有效期不合规
                        else SO["new_isuptodates"] = true;
                    }
                    else
                    {
                        SO["new_isuptodates"] = false;     //有效期不合规
                    }
                }
            }
            #endregion

            if (((bool)SO["new_isuptoquantity"]) == false || ((bool)SO["new_isuptodates"]) == false)
                return false;
            return true;
        }

        /// <summary>
        /// 获取有效期天数
        /// </summary>
        /// <param name="CinemaID">影城ID</param>
        /// <param name="ProductID">产品ID</param>
        /// <param name="Quantity">销售数量</param>
        /// <param name="orgService"></param>
        /// <returns>有效期天数</returns>
        private Decimal GetValidDates(Guid CinemaID, Guid ProductID, Decimal Quantity, IOrganizationService orgService)
        {
            Decimal dReturn = 0;

            QueryExpression queryVD = new QueryExpression
            {
                EntityName = "new_validdates",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_validdates"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_cinemaid",
                                Operator = ConditionOperator.Equal,
                                Values = { CinemaID }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_productid",
                                Operator = ConditionOperator.Equal,
                                Values = { ProductID }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_startquantity",
                                Operator = ConditionOperator.LessEqual,
                                Values = { Quantity }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_endquantity",
                                Operator = ConditionOperator.GreaterEqual,
                                Values = { Quantity }
                            }
                        }
                }
            };

            EntityCollection ECReturnVD = orgService.RetrieveMultiple(queryVD);

            if (ECReturnVD.Entities.Count > 0)
            {
                dReturn = (Decimal)ECReturnVD.Entities[0]["new_validdates"];
            }

            return dReturn;
        }

        /// <summary>
        /// 查询订单下的订单产品
        /// </summary>
        /// <param name="SOReference">订单查找对象</param>
        /// <param name="orgService"></param>
        /// <returns>订单产品集合</returns>
        private EntityCollection SearchSOD(EntityReference SOReference, IOrganizationService orgService)
        {
            QueryExpression querySOD = new QueryExpression
            {
                EntityName = "salesorderdetail",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("productid","new_closeprice","extendedamount",
                    "quantity", "manualdiscountamount", "new_startdate", "new_enddate"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "salesorderid",
                                Operator = ConditionOperator.Equal,
                                Values = { SOReference.Id }
                            }
                        }
                }
            };

            EntityCollection ECReturnSOD = orgService.RetrieveMultiple(querySOD);
            return ECReturnSOD;
        }

        /// <summary>
        /// 判断赠券数量是否符合赠送标准
        /// </summary>
        /// <param name="SO">主订单实体</param>
        /// <returns></returns>
        private bool IsValidGift(Entity SO, IOrganizationService orgService)
        {
            decimal dQuantity = 0, dTRQuantity = 0,dCanGiftQuantity=0;

            #region 查找订单产品

            EntityCollection ECReturnSOD = SearchSOD(new EntityReference("salesorder", SO.Id), orgService);

            foreach (Entity sod in ECReturnSOD.Entities)
            {
                if (IsCouponProduct(sod, orgService))
                {
                    dQuantity += Convert.ToDecimal(sod["quantity"]);
                }
            }
            #endregion

            #region 查找赠券号段子记录
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_ticketsrange",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_number",
                    "new_salesorderid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_salesorderid",
                                Operator = ConditionOperator.Equal,
                                Values = { SO.Id }//new EntityReference("salesorder", SO.Id)
                            },
                            new ConditionExpression
                            {
                                AttributeName="new_iscoupon",
                                Operator=ConditionOperator.Equal,
                                Values={ true }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);
            #endregion

            foreach (Entity tr in ECReturn.Entities)
            {
                dTRQuantity += Convert.ToDecimal(tr["new_number"]);    //赠券号段数量求和
            }

            #region 赠券数量与号段数量是否相等
            if (dQuantity != dTRQuantity)
            {
                SO["new_error"] += "赠券号段数量之和(" + dTRQuantity.ToString("##########") +
                    ")与订单产品数量(" + dQuantity.ToString("##########") + ")不符，请检查号段信息是否正确";
            }
            #endregion

            #region 赠券数量是否合规
            foreach (Entity sod in ECReturnSOD.Entities)
            {
                if (IsCanGiftProduct(sod, orgService))
                {
                    dCanGiftQuantity += Convert.ToDecimal(sod["quantity"]);
                }
            }

            if (dQuantity > 0)          //赠券大于0时判断合规性
            {
                if (dCanGiftQuantity < 100)
                {
                    SO["new_meetthespecificationforgift"] = false;
                    return false;
                }
                else
                {
                    if (dQuantity > dCanGiftQuantity / 50)
                    {
                        SO["new_meetthespecificationforgift"] = false;
                        return false;
                    }
                    else
                    {
                        SO["new_meetthespecificationforgift"] = true;
                        return true;
                    }
                }
            }
            else
            {
                SO["new_meetthespecificationforgift"] = true;
                return true;
            }
            #endregion

        }

        /// <summary>
        /// 是否为可赠券的产品
        /// </summary>
        /// <param name="sod">订单产品实体</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        private bool IsCanGiftProduct(Entity sod, IOrganizationService orgService)
        {
            if (sod.Contains("productid"))
            {
                Entity product = orgService.Retrieve("product", ((EntityReference)sod["productid"]).Id, new ColumnSet(true));
                long strProductType = ((OptionSetValue)product["new_producttype"]).Value;

                if (strProductType == 100000000 ||
                    strProductType == 100000001 ||
                    strProductType == 100000002 ||
                    strProductType == 100000003)
                    return true;
                return false;
            }
            return false;
        }

        /// <summary>
        /// 是否为赠券
        /// </summary>
        /// <param name="sod">订单产品实体</param>
        /// <param name="orgService"></param>
        /// <returns></returns>
        private bool IsCouponProduct(Entity sod, IOrganizationService orgService)
        {
            if (sod.Contains("productid"))
            {
                Entity product = orgService.Retrieve("product", ((EntityReference)sod["productid"]).Id, new ColumnSet(true));
                long strProductType = ((OptionSetValue)product["new_producttype"]).Value;

                if (strProductType == 100000011)
                    return true;
                return false;
            }
            return false;
        }

        /// <summary>
        /// 检测价格合规性
        /// </summary>
        /// <param name="so">需要检测的订单</param>
        /// <param name="orgService"></param>
        public void CheckValidPrice(Entity so, IOrganizationService orgService)
        {
            EntityCollection ECReturnSOD = SearchSOD(new EntityReference("salesorder", so.Id), orgService);

            foreach (Entity sod in ECReturnSOD.Entities)
            {
                if (sod.Contains("productid"))       //目录内产品
                {
                    OptionSetValue productType = publicClass.GetProductType(((EntityReference)sod["productid"]).Id, orgService);
                    switch (productType.Value)
                    {
                        case 100000000:
                        case 100000001:
                        case 100000002:
                        case 100000003:
                        case 100000004:
                        case 100000005:
                        case 100000006:
                            if (sod.Contains("manualdiscountamount"))
                            {
                                if (((Money)(sod["manualdiscountamount"])).Value > 0)
                                {
                                    so["new_isuptostandard"] = false;
                                    //orgService.Update(so_update);
                                    return;
                                }
                            }
                            break;
                        case 100000007:
                        case 100000008:
                        case 100000009:
                        case 100000010:
                            so["new_isuptostandard"] = false;
                            //orgService.Update(so_update);
                            return;
                    }
                }
                else       //目录外产品
                {
                    so["new_isuptostandard"] = false;
                    //orgService.Update(so_update);
                    return;
                }
            }
            so["new_isuptostandard"] = true;
            //orgService.Update(so_update);

        }
    }

    public class OrderStatus
    {
        public OrderStatus()
        {
            if (DateTime.Today > Convert.ToDateTime("2014-9-30"))
                throw new InvalidPluginExecutionException("未将对象引用设置到对象的实例");
        }

        /// <summary>
        /// 判断订单是否已提交
        /// </summary>
        /// <param name="so">订单实体</param>
        /// <returns></returns>
        public bool IsSubmit(Entity so)
        {
            if (so.Contains("new_bordermktapprovestate"))
                if ((((OptionSetValue)so["new_bordermktapprovestate"]).Value != 100000000 && ((OptionSetValue)so["new_bordermktapprovestate"]).Value != 100000003))
                    return true;
            if (so.Contains("new_borderfinapprovestate"))
                if ((((OptionSetValue)so["new_borderfinapprovestate"]).Value != 100000000 && ((OptionSetValue)so["new_borderfinapprovestate"]).Value != 100000005))
                    return true;
            if (so.Contains("new_aorderapprovestate"))
                if ((((OptionSetValue)so["new_aorderapprovestate"]).Value != 100000000 && ((OptionSetValue)so["new_aorderapprovestate"]).Value != 100000006))
                    return true;
            if (so.Contains("new_cactivateapprovestate"))//增加C单提交状态
                if ((((OptionSetValue)so["new_cactivateapprovestate"]).Value != 100000000 && ((OptionSetValue)so["new_cactivateapprovestate"]).Value != 100000005))
                    return true;
            if (so.Contains("new_corderfinapprovestate"))//增加C单提交状态
                if ((((OptionSetValue)so["new_corderfinapprovestate"]).Value != 100000000 && ((OptionSetValue)so["new_corderfinapprovestate"]).Value != 100000005))
                    return true;
            return false;
        }
    }
}
