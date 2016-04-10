using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Crm;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel;
using System.ServiceModel.Description;


namespace AccessCRMForWechat
{
    public class AccessCRM
    {
        IOrganizationService orgService;

        public AccessCRM(string Uri, string User, string Password)
        {
            try
            {
                Uri orgServiceUri = new Uri(Uri);//"https://wechatint.api.crm5.dynamics.com/XRMServices/2011/Organization.svc"
                var clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = User; //"wechatadmin@wechatint.onmicrosoft.com"
                clientCredentials.UserName.Password = Password; //"pass@word1"

                orgService = new OrganizationServiceProxy(orgServiceUri, null, clientCredentials, null);
            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "登陆CRM错误：" + ex.Message.ToString());
            }
        }

        #region 客户
        /// <summary>
        /// 写客户实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="email"></param>
        /// <param name="mobile"></param>
        /// <param name="isassociator">是否为会员，1是，0否</param>
        /// <param name="name"></param>
        /// <param name="isattention">是否关注，1是，0否</param>
        /// <param name="attentiontime"></param>
        /// <param name="cancelattentiontime"></param>
        /// <param name="sex"></param>
        /// <param name="country"></param>
        /// <param name="province"></param>
        /// <param name="city"></param>
        /// <param name="memberscore"></param>
        /// <returns>客户GUID</returns>
        public string WriteAccount(string openid, string token, string email, string mobile, string isassociator, string name,
            string isattention, string attentiontime, string cancelattentiontime, string sex, string country, string province,
            string city, string memberscore, string cusgroupid)
        {
            Entity acc,accWechatAcc;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }

            //todo：判断是否已经存在关联关系，存在则同时修改关系和客户，不存在则创建关系
            accWechatAcc = GetAccountWechatRelation(token, openid);
            if (accWechatAcc != null)
            {
                //修改客户
                acc = new Entity("account");
                acc.Id = ((EntityReference)accWechatAcc["new_accountid"]).Id;
                if (name.Length > 0) acc["new_nickname"] = name;
                if (email.Length > 0) acc["emailaddress1"] = email;
                if (mobile.Length > 0) acc["telephone1"] = mobile;
                if (sex.Length > 0) acc["new_sex"] = sex;
                if (country.Length > 0) acc["new_country"] = country;
                if (province.Length > 0) acc["new_province"] = province;
                if (city.Length > 0) acc["new_city"] = city;
                if (isattention.Length > 0) acc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                if (isassociator.Length > 0) acc["new_isassociator"] = isassociator == "1" ? true : false;
                if (attentiontime.Length > 0) acc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                if (cancelattentiontime.Length > 0) acc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                if (memberscore.Length > 0) acc["new_memberscore"] = Convert.ToDecimal(memberscore);
                if (cusgroupid.Length > 0) acc["new_cusgroupid"] = GetRefCusGroupEntity(token,cusgroupid);
                orgService.Update(acc);

                //修改关系
                if (name.Length > 0) accWechatAcc["new_nickname"] = name;
                if (sex.Length > 0) accWechatAcc["new_sex"] = sex;
                if (country.Length > 0) accWechatAcc["new_country"] = country;
                if (province.Length > 0) accWechatAcc["new_province"] = province;
                if (city.Length > 0) accWechatAcc["new_city"] = city;
                if (isattention.Length > 0) accWechatAcc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                if (attentiontime.Length > 0) accWechatAcc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                if (cancelattentiontime.Length > 0) accWechatAcc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                if (cusgroupid.Length > 0) accWechatAcc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                orgService.Update(accWechatAcc);
            }
            else   //不存在则创建关系
            {
                if (mobile.Length > 0)    //存在电话号码时的处理
                {    //
                    #region 以电话号码判断客户是否存在
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "account",
                        //ColumnSet = new ColumnSet(true),
                        ColumnSet = new ColumnSet("name",
                            "new_token", "new_openid"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "telephone1",
                                Operator = ConditionOperator.Equal,
                                Values = { mobile }
                            }
                        }
                        }
                    };

                    EntityCollection ECReturn = orgService.RetrieveMultiple(query);

                    if (ECReturn.Entities.Count > 0)   //客户存在，则只创建关系
                    {
                        acc = ECReturn.Entities[0];
                        accWechatAcc = new Entity("new_accountwechataccount");
                        if (name.Length > 0)
                        {
                            accWechatAcc["new_nickname"] = name;
                            accWechatAcc["new_name"] = name;
                        }
                        if (sex.Length > 0) accWechatAcc["new_sex"] = sex;
                        if (country.Length > 0) accWechatAcc["new_country"] = country;
                        if (province.Length > 0) accWechatAcc["new_province"] = province;
                        if (city.Length > 0) accWechatAcc["new_city"] = city;
                        if (isattention.Length > 0) accWechatAcc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (attentiontime.Length > 0) accWechatAcc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) accWechatAcc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (cusgroupid.Length > 0) accWechatAcc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        accWechatAcc["new_wechataccount"] = GetRefWechatAccountEntity(token);
                        accWechatAcc["new_accountid"] = acc.ToEntityReference();
                        accWechatAcc["new_token"] = token;
                        accWechatAcc["new_openid"] = openid;

                        orgService.Create(accWechatAcc);
                    }
                    else   //客户不存在，则创建客户、关系
                    {
                        //创建客户
                        acc = new Entity("account");
                        if (name.Length > 0)
                        {
                            acc["new_nickname"] = name;
                            acc["name"] = name;
                        }
                        else
                        {
                            acc["new_nickname"] = openid;
                            acc["name"] = openid;
                        }
                        if (email.Length > 0) acc["emailaddress1"] = email;
                        if (mobile.Length > 0) acc["telephone1"] = mobile;
                        if (sex.Length > 0) acc["new_sex"] = sex;
                        if (country.Length > 0) acc["new_country"] = country;
                        if (province.Length > 0) acc["new_province"] = province;
                        if (city.Length > 0) acc["new_city"] = city;
                        if (isattention.Length > 0) acc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (isassociator.Length > 0) acc["new_isassociator"] = isassociator == "1" ? true : false;
                        if (attentiontime.Length > 0) acc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) acc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (memberscore.Length > 0) acc["new_memberscore"] = Convert.ToDecimal(memberscore);
                        if (cusgroupid.Length > 0) acc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        acc["new_openid"] = openid;
                        acc["new_token"] = token;
                        acc["new_wechataccount"] = GetRefWechatAccountEntity(token);

                        acc.Id = orgService.Create(acc);

                        //创建关系
                        accWechatAcc = new Entity("new_accountwechataccount");
                        if (name.Length > 0)
                        {
                            accWechatAcc["new_nickname"] = name;
                            accWechatAcc["new_name"] = name;
                        }
                        if (sex.Length > 0) accWechatAcc["new_sex"] = sex;
                        if (country.Length > 0) accWechatAcc["new_country"] = country;
                        if (province.Length > 0) accWechatAcc["new_province"] = province;
                        if (city.Length > 0) accWechatAcc["new_city"] = city;
                        if (isattention.Length > 0) accWechatAcc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (attentiontime.Length > 0) accWechatAcc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) accWechatAcc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (cusgroupid.Length > 0) accWechatAcc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        accWechatAcc["new_wechataccount"] = GetRefWechatAccountEntity(token);
                        accWechatAcc["new_accountid"] = acc.ToEntityReference();
                        accWechatAcc["new_token"] = token;
                        accWechatAcc["new_openid"] = openid;

                        orgService.Create(accWechatAcc);
                    }
                    #endregion
                }
                else      //不存在电话号码时的处理
                {
                    #region 以token\openid判断客户是否存在
                    QueryExpression query = new QueryExpression
                    {
                        EntityName = "account",
                        //ColumnSet = new ColumnSet(true),
                        ColumnSet = new ColumnSet("name",
                            "new_token", "new_openid"),
                        Criteria = new FilterExpression
                        {
                            Conditions =
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "new_token",
                                        Operator = ConditionOperator.Equal,
                                        Values = { token }
                                    },
                                    new ConditionExpression
                                    {
                                        AttributeName = "new_openid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { openid }
                                    }
                                }
                        }
                    };

                    EntityCollection ECReturn = orgService.RetrieveMultiple(query);

                    if (ECReturn.Entities.Count > 0)   //客户存在，则只创建关系
                    {
                        acc = ECReturn.Entities[0];
                        accWechatAcc = new Entity("new_accountwechataccount");
                        if (name.Length > 0)
                        {
                            accWechatAcc["new_nickname"] = name;
                            accWechatAcc["new_name"] = name;
                        }
                        if (sex.Length > 0) accWechatAcc["new_sex"] = sex;
                        if (country.Length > 0) accWechatAcc["new_country"] = country;
                        if (province.Length > 0) accWechatAcc["new_province"] = province;
                        if (city.Length > 0) accWechatAcc["new_city"] = city;
                        if (isattention.Length > 0) accWechatAcc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (attentiontime.Length > 0) accWechatAcc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) accWechatAcc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (cusgroupid.Length > 0) accWechatAcc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        accWechatAcc["new_wechataccount"] = GetRefWechatAccountEntity(token);
                        accWechatAcc["new_accountid"] = acc.ToEntityReference();
                        accWechatAcc["new_token"] = token;
                        accWechatAcc["new_openid"] = openid;

                        orgService.Create(accWechatAcc);
                    }
                    else   //客户不存在，则创建客户、关系
                    {
                        //创建客户
                        acc = new Entity("account");
                        if (name.Length > 0)
                        {
                            acc["new_nickname"] = name;
                            acc["name"] = name;
                        }
                        else
                        {
                            acc["new_nickname"] = openid;
                            acc["name"] = openid;
                        }
                        if (email.Length > 0) acc["emailaddress1"] = email;
                        if (mobile.Length > 0) acc["telephone1"] = mobile;
                        if (sex.Length > 0) acc["new_sex"] = sex;
                        if (country.Length > 0) acc["new_country"] = country;
                        if (province.Length > 0) acc["new_province"] = province;
                        if (city.Length > 0) acc["new_city"] = city;
                        if (isattention.Length > 0) acc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (isassociator.Length > 0) acc["new_isassociator"] = isassociator == "1" ? true : false;
                        if (attentiontime.Length > 0) acc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) acc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (memberscore.Length > 0) acc["new_memberscore"] = Convert.ToDecimal(memberscore);
                        if (cusgroupid.Length > 0) acc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        acc["new_openid"] = openid;
                        acc["new_token"] = token;
                        acc["new_wechataccount"] = GetRefWechatAccountEntity(token);

                        acc.Id = orgService.Create(acc);

                        //创建关系
                        accWechatAcc = new Entity("new_accountwechataccount");
                        if (name.Length > 0)
                        {
                            accWechatAcc["new_nickname"] = name;
                            accWechatAcc["new_name"] = name;
                        }
                        if (sex.Length > 0) accWechatAcc["new_sex"] = sex;
                        if (country.Length > 0) accWechatAcc["new_country"] = country;
                        if (province.Length > 0) accWechatAcc["new_province"] = province;
                        if (city.Length > 0) accWechatAcc["new_city"] = city;
                        if (isattention.Length > 0) accWechatAcc["new_attentionstatus"] = isattention == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                        if (attentiontime.Length > 0) accWechatAcc["new_attentiontime"] = Convert.ToDateTime(attentiontime);
                        if (cancelattentiontime.Length > 0) accWechatAcc["new_cancelattentiontime"] = Convert.ToDateTime(cancelattentiontime);
                        if (cusgroupid.Length > 0) accWechatAcc["new_cusgroupid"] = GetRefCusGroupEntity(token, cusgroupid);
                        accWechatAcc["new_wechataccount"] = GetRefWechatAccountEntity(token);
                        accWechatAcc["new_accountid"] = acc.ToEntityReference();
                        accWechatAcc["new_token"] = token;
                        accWechatAcc["new_openid"] = openid;

                        orgService.Create(accWechatAcc);
                    }
                    #endregion
                }
            }

            return acc.Id.ToString();
        }

        public Entity GetAccountWechatRelation(string token, string openid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_accountwechataccount",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name","new_accountid",
                    "new_token", "new_openid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_openid",
                                Operator = ConditionOperator.Equal,
                                Values = { openid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);
            if (ECReturn.Entities.Count > 0)
            {
                return ECReturn.Entities[0];
            }
            else
                return null;

        }

        public void StopAccount(string AccountID)
        {
            Entity acc = orgService.Retrieve("account", new Guid(AccountID), new ColumnSet(("name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = acc.ToEntityReference(); ;

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 会员总积分
        /// <summary>
        /// 写客户实体会员总积分
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="email"></param>
        /// <param name="mobile"></param>
        /// <param name="isassociator">是否为会员，1是，0否</param>
        /// <param name="name"></param>
        /// <param name="sex"></param>
        /// <param name="memberscore"></param>
        /// <returns>客户GUID</returns>
        public string WriteScore(string openid, string token, string email, string mobile, string isassociator, string name,
            string memberscore)
        {
            Entity acc;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "account",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name",
                    "new_token", "new_openid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_openid",
                                Operator = ConditionOperator.Equal,
                                Values = { openid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                acc = ECReturn.Entities[0];
                if (name.Length > 0) acc["name"] = name; else acc["name"] = openid;
                if (email.Length > 0) acc["emailaddress1"] = email;
                if (mobile.Length > 0) acc["telephone1"] = mobile;
                if (isassociator.Length > 0) acc["new_isassociator"] = isassociator == "1" ? true : false;
                if (memberscore.Length > 0) acc["new_memberscore"] = Convert.ToDecimal(memberscore);
                orgService.Update(acc);
            }
            else
            {
                acc = new Entity("account");
                if (name.Length > 0) acc["name"] = name; else acc["name"] = openid;
                if (email.Length > 0) acc["emailaddress1"] = email;
                if (mobile.Length > 0) acc["telephone1"] = mobile;
                if (isassociator.Length > 0) acc["new_isassociator"] = isassociator == "1" ? true : false;
                if (memberscore.Length > 0) acc["new_memberscore"] = Convert.ToDecimal(memberscore);
                acc["new_openid"] = openid;
                acc["new_token"] = token;
                acc["new_wechataccount"] = GetRefWechatAccountEntity(token);
                return orgService.Create(acc).ToString();
            }
            return acc.Id.ToString();
        }

        #endregion

        #region 会员卡
        /// <summary>
        /// 写会员卡实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="number">卡号</param>
        /// <param name="status">1为开启，0为关闭</param>
        /// <param name="registtime"></param>
        /// <param name="endtime"></param>
        /// <returns>会员卡GUID</returns>
        public string WriteAssociator(string openid, string token, string number, string status, string registtime, string endtime)
        {
            Entity associator, acc;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }

            #region 更新关联客户的会员标志
            QueryExpression query = new QueryExpression
            {
                EntityName = "account",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name",
                    "new_token", "new_openid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_openid",
                                Operator = ConditionOperator.Equal,
                                Values = { openid }
                            }
                            //,
                            //new ConditionExpression
                            //{
                            //    AttributeName = "new_isassociator",
                            //    Operator = ConditionOperator.Equal,
                            //    Values = { new OptionSetValue(0) }
                            //}
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                acc = ECReturn.Entities[0];
                acc["new_isassociator"] = true;
                orgService.Update(acc);
            }
            #endregion

            query = new QueryExpression
            {
                EntityName = "new_member",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_openid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_name",
                                Operator = ConditionOperator.Equal,
                                Values = { number }
                            }
                        }
                }
            };

            ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                associator = ECReturn.Entities[0];
                associator["new_openid"] = openid;
                if (status.Length > 0) associator["new_memberstatus"] = status == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                if (registtime.Length > 0) associator["new_registemembertime"] = Convert.ToDateTime(registtime);
                if (endtime.Length > 0) associator["new_endtime"] = Convert.ToDateTime(endtime);
                associator["new_accountid"] = GetRefAccountEntity(openid, token);
                orgService.Update(associator);
            }
            else
            {
                associator = new Entity("new_member");
                if (number.Length > 0) associator["new_name"] = number; else associator["new_name"] = openid;
                if (status.Length > 0) associator["new_memberstatus"] = status == "1" ? new OptionSetValue(100000000) : new OptionSetValue(100000001);
                if (registtime.Length > 0) associator["new_registemembertime"] = Convert.ToDateTime(registtime);
                if (endtime.Length > 0) associator["new_endtime"] = Convert.ToDateTime(endtime);
                associator["new_openid"] = openid;
                associator["new_token"] = token;
                associator["new_accountid"] = GetRefAccountEntity(openid, token);
                associator["new_wechataccountid"] = GetRefWechatAccountEntity(token);
                return orgService.Create(associator).ToString();
            }
            return associator.Id.ToString();
        }

        public void StopAssociator(string AssociatorID)
        {
            Entity acc = orgService.Retrieve("new_member", new Guid(AssociatorID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = acc.ToEntityReference(); ;

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 图文
        /// <summary>
        /// 写图文实体
        /// </summary>
        /// <param name="openid">图文在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="title"></param>
        /// <returns>图文GUID</returns>
        public string WriteIMG(string sourceid, string token, string title)
        {
            Entity img;

            if (sourceid.Length == 0 || token.Length == 0)
            {
                throw new Exception("sourceid 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_imagetextinfo",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_title",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                img = ECReturn.Entities[0];
                if (title.Length > 0)
                {
                    img["new_title"] = title;
                    img["new_name"] = title;
                }
                orgService.Update(img);
            }
            else
            {
                img = new Entity("new_imagetextinfo");
                if (title.Length > 0)
                {
                    img["new_title"] = title;
                    img["new_name"] = title;
                }
                img["new_sourceid"] = sourceid;
                img["new_token"] = token;
                return orgService.Create(img).ToString();
            }
            return img.Id.ToString();
        }

        public void StopIMG(string IMGID)
        {
            Entity img = orgService.Retrieve("new_imagetextinfo", new Guid(IMGID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 微信关注者分组
        /// <summary>
        /// 写微信关注者分组实体
        /// </summary>
        /// <param name="openid">微信关注者分组在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="name"></param>
        /// <returns>微信关注者分组GUID</returns>
        public string WriteCusGroup(string sourceid, string token, string name)
        {
            Entity cusgroup;

            if (sourceid.Length == 0 || token.Length == 0)
            {
                throw new Exception("sourceid 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_cusgroup",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                cusgroup = ECReturn.Entities[0];
                if (name.Length > 0)
                {
                    cusgroup["new_name"] = name;
                }
                orgService.Update(cusgroup);
            }
            else
            {
                cusgroup = new Entity("new_cusgroup");
                if (name.Length > 0)
                {
                    cusgroup["new_name"] = name;
                }
                cusgroup["new_sourceid"] = sourceid;
                cusgroup["new_token"] = token;
                cusgroup["new_wxaccountid"] = GetRefWechatAccountEntity(token);
                return orgService.Create(cusgroup).ToString();
            }
            return cusgroup.Id.ToString();
        }

        public void StopCusGroup(string CusGroupID)
        {
            Entity img = orgService.Retrieve("new_cusgroup", new Guid(CusGroupID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 微信互动
        /// <summary>
        /// 写微信互动实体
        /// </summary>
        /// <param name="sourceid">微信互动在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="name"></param>
        /// <returns>微信互动GUID</returns>
        public string WriteLottery(string sourceid, string token, string name,string type)
        {
            Entity lottery;

            if (sourceid.Length == 0 || token.Length == 0)
            {
                throw new Exception("sourceid 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_lottery",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                lottery = ECReturn.Entities[0];
                if (name.Length > 0)
                {
                    lottery["new_name"] = name;
                }
                switch (type)
                {
                    case "1":
                        lottery["new_type"] = new OptionSetValue(100000001);
                        break;
                    case "2":
                        lottery["new_type"] = new OptionSetValue(100000002);
                        break;
                    case "3":
                        lottery["new_type"] = new OptionSetValue(100000003);
                        break;
                }
                orgService.Update(lottery);
            }
            else
            {
                lottery = new Entity("new_imagetextinfo");
                if (name.Length > 0)
                {
                    lottery["new_name"] = name;
                }
                lottery["new_sourceid"] = sourceid;
                lottery["new_token"] = token;
                lottery["new_wxaccountid"] = GetRefWechatAccountEntity(token);
                switch (type)
                {
                    case "1":
                        lottery["new_type"] = new OptionSetValue(100000001);
                        break;
                    case "2":
                        lottery["new_type"] = new OptionSetValue(100000002);
                        break;
                    case "3":
                        lottery["new_type"] = new OptionSetValue(100000003);
                        break;
                }
                return orgService.Create(lottery).ToString();
            }
            return lottery.Id.ToString();
        }

        public void StopLottery(string LotteryID)
        {
            Entity lottery = orgService.Retrieve("new_lottery", new Guid(LotteryID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = lottery.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 门店
        /// <summary>
        /// 写门店实体
        /// </summary>
        /// <param name="sourceid">门店在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="name"></param>
        /// <param name="address"></param>
        /// <param name="isbranch">是否分支机构</param>
        /// <returns>门店GUID</returns>
        public string WriteShop(string sourceid, string token, string name, string address, Boolean isbranch)
        {
            Entity shop;

            if (sourceid.Length == 0 || token.Length == 0)
            {
                throw new Exception("sourceid 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_shop",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                shop = ECReturn.Entities[0];
                shop["new_name"] = name;
                shop["new_address"] = address;
                shop["new_isbranch"] = isbranch;
                orgService.Update(shop);
            }
            else
            {
                shop = new Entity("new_shop");
                shop["new_name"] = name;
                shop["new_address"] = address;
                shop["new_isbranch"] = isbranch;
                shop["new_sourceid"] = sourceid;
                shop["new_token"] = token;
                shop["new_wxaccountid"] = GetRefWechatAccountEntity(token);
                return orgService.Create(shop).ToString();
            }
            return shop.Id.ToString();
        }

        public void StopShop(string ShopID)
        {
            Entity img = orgService.Retrieve("new_shop", new Guid(ShopID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 商品关联记录
        /// <summary>
        /// 写商品关联记录实体
        /// </summary>
        /// <param name="openid">商品关联记录在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="title"></param>
        /// <returns>商品关联记录GUID</returns>
        public string WritePrd2prd(string token, string productid1, string productid2)
        {
            Entity prd2prd;
            EntityReference refProductid1, refProductid2;

            if (productid1.Length == 0 || productid2.Length == 0 || token.Length == 0)
            {
                throw new Exception("productid1 或 productid1 或 token不能为空");
            }

            refProductid1 = GetRefProductEntity(productid1);
            refProductid2 = GetRefProductEntity(productid2);

            prd2prd = new Entity("new_prd2prd");
            prd2prd["new_productid1"] = refProductid1;
            prd2prd["new_name"] = refProductid1.Name + "->" + refProductid2.Name;
            prd2prd["new_productid2"] = refProductid2;
            prd2prd["new_wxaccountid"] = GetRefWechatAccountEntity(token);
            return orgService.Create(prd2prd).ToString();
        }

        public void StopPrd2prd(string prd2prdID)
        {
            Entity img = orgService.Retrieve("new_prd2prd", new Guid(prd2prdID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 商品图文关联记录
        /// <summary>
        /// 写商品图文关联记录实体
        /// </summary>
        /// <param name="openid">商品图文关联记录在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="title"></param>
        /// <returns>商品图文关联记录GUID</returns>
        public string WritePrd2img(string token, string productid, string imgid)
        {
            Entity prd2img;
            EntityReference refProductid, refImgid;

            if (productid.Length == 0 || imgid.Length == 0 || token.Length == 0)
            {
                throw new Exception("productid 或 imgid 或 token不能为空");
            }

            refProductid = GetRefProductEntity(productid);
            refImgid = GetRefImageTextInfoEntity(token, imgid);
            prd2img = new Entity("new_prd2img");
            prd2img["new_productid"] = refProductid;
            prd2img["new_imgid"] = refImgid;
            prd2img["new_name"] = refProductid.Name + "<->" + refImgid.Name;
            prd2img["new_wxaccountid"] = GetRefWechatAccountEntity(token);
            return orgService.Create(prd2img).ToString();
        }

        public void StopPrd2img(string prd2prdID)
        {
            Entity img = orgService.Retrieve("new_prd2img", new Guid(prd2prdID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 报名
        /// <summary>
        /// 写报名实体
        /// </summary>
        /// <param name="GUID"></param>
        /// <param name="openid">报名在微信中的ID</param>
        /// <param name="token"></param>
        /// <param name="name"></param>
        /// <param name="note"></param>
        /// <param name="enrolltime"></param>
        /// <param name="status"></param>
        /// <returns>报名GUID</returns>
        public string WriteEnroll(string GUID, string openid, string token, string name, string note, string enrolltime, string status)
        {
            Entity enroll;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }

            if (GUID.Length > 0)
            {
                enroll = orgService.Retrieve("new_enroll", new Guid(GUID), new ColumnSet(("new_name")));

                if (enroll == null)
                    throw new Exception("CRM系统中不存在ID为：" + GUID + "的报名记录。");

                if (status.Length > 0)
                {
                    enroll["new_handlestatus"] = status;
                    enroll["new_enrolltime"] = Convert.ToDateTime(enrolltime);
                }
                orgService.Update(enroll);
            }
            else
            {
                enroll = new Entity("new_enroll");
                if (name.Length > 0)
                {
                    enroll["new_name"] = name;
                }
                if (note.Length > 0) enroll["new_note"] = note;
                enroll["new_openid"] = openid;
                enroll["new_token"] = token;
                enroll["new_accountid"] = GetRefAccountEntity(openid, token);
                if (status.Length > 0) enroll["new_handlestatus"] = status;
                if (enrolltime.Length > 0) enroll["new_enrolltime"] = Convert.ToDateTime(enrolltime);
                enroll["new_wechataccountid"] = GetRefWechatAccountEntity(token);
                return orgService.Create(enroll).ToString();
            }
            return enroll.Id.ToString();
        }

        public void StopEnroll(string EnrollID)
        {
            Entity enroll = orgService.Retrieve("new_enroll", new Guid(EnrollID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = enroll.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 行为历史
        /// <summary>
        /// 写行为历史实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="imgid">图文ID</param>
        /// <param name="occuredtime"></param>
        /// <param name="actiontype">动作类型:0：打开；1：分享；2：收藏；3：加入购物车；4：阅读；5：购买；6：扫描二维码</param>
        /// <param name="objecttype">0：图文信息、1：LBS、2：商品、3：门店</param>
        /// <returns>行为历史GUID</returns>
        public string WriteAction(string openid, string token, string imgid, string occuredtime, string actiontype, string objecttype)
        {
            Entity action;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }


            action = new Entity("new_sharehistory");
            action["new_accountid"] = GetRefAccountEntity(openid, token);
            if(action["new_accountid"]==null)
            {
                throw new Exception("公众号token:"+token+"中不存在openid为"+openid+"的客户");
            }
            switch (objecttype)
            {
                case "0":   //0：图文信息、1：LBS、2：商品、3：门店
                    action["new_imagetextinfoid"] = GetRefImageTextInfoEntity(token, imgid);
                    action["new_name"] = ((EntityReference)action["new_imagetextinfoid"]).Name;
                    action["new_actionobject"] = new OptionSetValue(100000000);
                    break;
                case "1":
                    //action["new_imagetextinfoid"] = GetRefImageTextInfoEntity(token, imgid);
                    //action["new_name"] = ((EntityReference)action["new_imagetextinfoid"]).Name;
                    action["new_actionobject"] = new OptionSetValue(100000001);
                    break;
                case "2":
                    action["new_productid"] = GetRefProductEntity(imgid);
                    action["new_name"] = ((EntityReference)action["new_productid"]).Name;
                    action["new_actionobject"] = new OptionSetValue(100000002);
                    break;
                case "3":
                    action["new_shopid"] = GetRefShopEntity(token, imgid);
                    action["new_name"] = ((EntityReference)action["new_shopid"]).Name;
                    action["new_actionobject"] = new OptionSetValue(100000003);
                    break;
            }
            action["new_occuredtime"] = Convert.ToDateTime(occuredtime);
            switch (actiontype)
            {
                case "0":    //0：打开；1：分享；2：收藏；3：加入购物车；4：阅读；5：购买；6：扫描二维码
                    action["new_type"] = new OptionSetValue(100000000);
                    break;
                case "1":
                    action["new_type"] = new OptionSetValue(100000001);
                    break;
                case "2":
                    action["new_type"] = new OptionSetValue(100000002);
                    break;
                case "3":
                    action["new_type"] = new OptionSetValue(100000003);
                    break;
                case "4":
                    action["new_type"] = new OptionSetValue(100000004);
                    break;
                case "5":
                    action["new_type"] = new OptionSetValue(100000005);
                    break;
                case "6":
                    action["new_type"] = new OptionSetValue(100000006);
                    break;
            }
            action["new_token"] = token;
            action["new_wechataccountid"] = GetRefWechatAccountEntity(token);
            return orgService.Create(action).ToString();
        }

        #endregion

        #region 公众号
        /// <summary>
        /// 写公众号实体
        /// </summary>
        /// <param name="token"></param>
        /// <param name="name"></param>
        /// <returns>公众号GUID</returns>
        public string WriteWechatAccount(string token, string name)
        {
            Entity wa;

            if (name.Length == 0 || token.Length == 0)
            {
                throw new Exception("wxname 或 token不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_wechataccount",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                wa = ECReturn.Entities[0];
                if (name.Length > 0)
                {
                    wa["new_name"] = name;
                }
                orgService.Update(wa);
            }
            else
            {
                wa = new Entity("new_wechataccount");
                if (name.Length > 0)
                {
                    wa["new_name"] = name;
                }
                wa["new_token"] = token;
                return orgService.Create(wa).ToString();
            }
            return wa.Id.ToString();
        }

        public void StopWechatAccount(string WAID)
        {
            Entity img = orgService.Retrieve("new_wechataccount", new Guid(WAID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 公众号管理员
        /// <summary>
        /// 写公众号管理员实体
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="parentid"></param>
        /// <returns>公众号GUID</returns>
        public string WriteWechatAccountAdmin(string userid,string parentid,string email,Boolean issuper,string id)
        {
            Entity wa;
            EntityReference er;

            if (userid.Length == 0)
            {
                throw new Exception("公众号管理员userid不能为空");
            }

            QueryExpression query = new QueryExpression
            {
                EntityName = "new_wechataccountuser",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_email"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_id",
                                Operator = ConditionOperator.Equal,
                                Values = { id }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                wa = ECReturn.Entities[0];
                if (email.Length > 0)
                {
                    wa["new_email"] = email;
                }
                if (parentid.Length > 0)
                {
                    er = GetRefWechatAccountAdminEntity(parentid);
                    if (er != null)
                        wa["new_parentadmin"] = er;
                }
                orgService.Update(wa);
            }
            else
            {
                wa = new Entity("new_wechataccountuser");
                if (email.Length > 0)
                {
                    wa["new_email"] = email;
                }
                if (parentid.Length > 0)
                {
                    er = GetRefWechatAccountAdminEntity(parentid);
                    if (er != null)
                        wa["new_parentadmin"] = er;
                }
                wa["new_issuperadmin"] = issuper;
                wa["new_userid"] = userid;
                wa["new_name"] = userid;
                wa["new_id"] = id;
                return orgService.Create(wa).ToString();
            }
            return wa.Id.ToString();
        }

        public void StopWechatAccountAdmin(string WAID)
        {
            Entity img = orgService.Retrieve("new_wechataccountuser", new Guid(WAID), new ColumnSet(("new_name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = img.ToEntityReference();

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 订单
        /// <summary>
        /// 写订单实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="amount">金额</param>
        /// <param name="orderdate"></param>
        /// <param name="name"></param>
        /// <returns>订单GUID</returns>
        public string WriteOrder(string openid, string token, string amount, string orderdate, string name)
        {
            Entity order;
            EntityReference refAccount, refPriceLevel;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }
            else
            {
                refAccount = GetRefAccountEntity(openid, token);
                if (refAccount == null)
                {
                    throw new Exception("标识为openid=" + openid + ",token=" + token + "的客户CRM中不存在。");
                }
            }

            refPriceLevel = GetRefPriceLevelEntity();
            if (refPriceLevel == null)
            {
                throw new Exception("CRM中不存在价目表。");
            }

            order = new Entity("salesorder");
            order["requestdeliveryby"] = Convert.ToDateTime(orderdate);
            order["totalamount"] = new Money(Convert.ToDecimal(amount));
            order["totalamount_base"] = new Money(Convert.ToDecimal(amount));
            order["name"] = name;
            order["customerid"] = refAccount;
            order["pricelevelid"] = refPriceLevel;
            order["new_wechataccountid"] = GetRefWechatAccountEntity(token);
            return orgService.Create(order).ToString();
        }

        public void StopOrder(string OrderID)
        {
            Entity acc = orgService.Retrieve("salesorder", new Guid(OrderID), new ColumnSet(("name")));
            // Create the Request Object
            SetStateRequest state = new SetStateRequest();

            // Set the Request Object's Properties
            state.State = new OptionSetValue(1);
            state.Status = new OptionSetValue(2);

            // Point the Request to the case whose state is being changed
            state.EntityMoniker = acc.ToEntityReference(); ;

            // Execute the Request
            SetStateResponse stateSet = (SetStateResponse)orgService.Execute(state);

        }
        #endregion

        #region 订单明细
        /// <summary>
        /// 写订单明细实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="cardnumber">商品的微信ID</param>
        /// <param name="type"></param>
        /// <param name="scoretype"></param>
        /// <param name="score">金额</param>
        /// <returns>订单明细GUID</returns>
        public string WriteOrderDetail(string OrderGUID, string token, string productid, string quantity, string price, string amount)
        {
            Entity order;
            EntityReference refOrder, refProduct, refUom;

            if (OrderGUID.Length == 0 || token.Length == 0)
            {
                throw new Exception("OrderGUID 或 token不能为空");
            }
            else
            {
                refOrder = GetRefOrderEntity(OrderGUID);
                if (refOrder == null)
                {
                    throw new Exception("标识为OrderGUID=" + OrderGUID + "的订单CRM中不存在。");
                }
            }

            refProduct = GetRefProductEntity(productid);
            if (refProduct == null)
            {
                throw new Exception("CRM中不存在ID为" + productid + "的产品。");
            }

            refUom = GetRefUomEntity();
            if (refUom == null)
            {
                throw new Exception("CRM中不存在计量单位。");
            }

            order = new Entity("salesorderdetail");
            order["priceperunit"] = new Money(Convert.ToDecimal(price));
            order["quantity"] = Convert.ToDecimal(quantity);
            order["baseamount"] = new Money(Convert.ToDecimal(amount));
            order["ispriceoverridden"] = true;    //自定义价格
            order["isproductoverridden"] = false;  //目录内
            order["salesorderid"] = refOrder;
            order["productid"] = refProduct;
            order["uomid"] = refUom;
            return orgService.Create(order).ToString();
        }

        #endregion

        #region 积分增减明细
        /// <summary>
        /// 写积分明细实体
        /// </summary>
        /// <param name="openid"></param>
        /// <param name="token"></param>
        /// <param name="cardnumber">会员卡号</param>
        /// <param name="type">积分增减方式</param>
        /// <param name="scoretype">积分增减原因</param>
        /// <param name="score">积分</param>
        /// <param name="scoretime">积分时间</param>
        /// <param name="note"></param>
        /// <returns>订单明细GUID</returns>
        public string WriteScoreDetail(string openid, string token, string cardnumber, string type, string scoretype, string score,
            string scoretime, string note)
        {
            Entity scoredetail;
            EntityReference refAccount, refMember = null;

            if (openid.Length == 0 || token.Length == 0)
            {
                throw new Exception("openid 或 token不能为空");
            }
            else
            {
                refAccount = GetRefAccountEntity(openid, token);
                if (refAccount == null)
                {
                    throw new Exception("CRM中不存在公众号为token=" + token + "，openid=" + openid + "的客户。");
                }
            }

            if (cardnumber.Length > 0)
            {
                refMember = GetRefMemberEntity(cardnumber);
                if (refMember == null)
                {
                    throw new Exception("CRM中不存在卡号为" + cardnumber + "的会员卡。");

                }
            }


            scoredetail = new Entity("new_memberscore");
            scoredetail["new_accountid"] = refAccount;
            if (refMember != null) scoredetail["new_memeberid"] = refMember;
            if (note.Length > 0) scoredetail["new_note"] = note;
            scoredetail["new_openid"] = openid;    //
            scoredetail["new_token"] = token;  //
            scoredetail["new_score"] = Convert.ToDecimal(score);
            scoredetail["new_scoretime"] = Convert.ToDateTime(scoretime);
            if (scoretype.Length > 0) scoredetail["new_scoretype"] = new OptionSetValue(Convert.ToInt32(scoretype));
            if (type.Length > 0) scoredetail["new_type"] = new OptionSetValue(Convert.ToInt32(type));
            return orgService.Create(scoredetail).ToString();
            scoredetail["new_wechataccountid"] = GetRefWechatAccountEntity(token);
        }

        #endregion


        #region 查找参照实体

        private EntityReference GetRefWechatAccountAdminEntity(string parentid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_wechataccountuser",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_id",
                                Operator = ConditionOperator.Equal,
                                Values = { parentid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                EntityReference er = ECReturn.Entities[0].ToEntityReference();
                er.Name = ECReturn.Entities[0]["new_name"].ToString();
                return er;
            }
            else
                return null;

        }

        private EntityReference GetRefMemberEntity(string cardnumber)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_member",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_name",
                                Operator = ConditionOperator.Equal,
                                Values = { cardnumber }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                EntityReference er = ECReturn.Entities[0].ToEntityReference();
                er.Name = ECReturn.Entities[0]["new_name"].ToString();
                return er;
            }
            else
                return null;

        }

        EntityReference GetRefProductEntity(string productid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "product",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "productnumber",
                                Operator = ConditionOperator.Equal,
                                Values = { productid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                EntityReference er = ECReturn.Entities[0].ToEntityReference();
                er.Name = ECReturn.Entities[0]["name"].ToString();
                return er;
            }
            else
                return null;

        }

        private EntityReference GetRefOrderEntity(string OrderGUID)
        {
            Entity order = orgService.Retrieve("salesorder", new Guid(OrderGUID), new ColumnSet(("name")));

            return order.ToEntityReference();
        }

        private EntityReference GetRefWechatAccountEntity(string token)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_wechataccount",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
                return ECReturn.Entities[0].ToEntityReference();
            else
                return null;

        }

        private EntityReference GetRefCusGroupEntity(string token,string cusgroupid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_cusgroup",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { cusgroupid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
                return ECReturn.Entities[0].ToEntityReference();
            else
                return null;

        }

        private EntityReference GetRefAccountEntity(string openid, string token)
        {
            QueryExpression query;
            EntityCollection ECReturn;

            query = new QueryExpression
            {
                EntityName = "account",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name",
                    "new_token", "new_openid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_openid",
                                Operator = ConditionOperator.Equal,
                                Values = { openid }
                            }
                        }
                }
            };

            ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
                return ECReturn.Entities[0].ToEntityReference();
            else
            {
                query = new QueryExpression
                {
                    EntityName = "new_accountwechataccount",
                    //ColumnSet = new ColumnSet(true),
                    ColumnSet = new ColumnSet("new_name", "new_accountid",
                        "new_token", "new_openid"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_openid",
                                Operator = ConditionOperator.Equal,
                                Values = { openid }
                            }
                        }
                    }
                };

                ECReturn = orgService.RetrieveMultiple(query);

                if (ECReturn.Entities.Count > 0)
                    return (EntityReference)ECReturn.Entities[0]["new_accountid"]; //ECReturn.Entities[0].ToEntityReference();
                else
                    return null;
            }

        }

        private EntityReference GetRefPriceLevelEntity()
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "pricelevel",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = { "默认价目表" }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
                return ECReturn.Entities[0].ToEntityReference();
            else
                return null;

        }

        private EntityReference GetRefUomEntity()
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "uom",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("name"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "name",
                                Operator = ConditionOperator.Equal,
                                Values = { "基本计价单位" }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
                return ECReturn.Entities[0].ToEntityReference();
            else
                return null;

        }

        private EntityReference GetRefImageTextInfoEntity(string token, string sourceid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_imagetextinfo",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                EntityReference er = ECReturn.Entities[0].ToEntityReference();
                er.Name = ECReturn.Entities[0]["new_name"].ToString();
                return er;
            }
            else
                return null;

        }

        private EntityReference GetRefShopEntity(string token, string sourceid)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "new_shop",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_name",
                    "new_token", "new_sourceid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values = { token }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_sourceid",
                                Operator = ConditionOperator.Equal,
                                Values = { sourceid }
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);

            if (ECReturn.Entities.Count > 0)
            {
                EntityReference er = ECReturn.Entities[0].ToEntityReference();
                er.Name = ECReturn.Entities[0]["new_name"].ToString();
                return er;
            }
            else
                return null;

        }

        #endregion

        #region 获取待发微信实体
        
        public void WriteSendLog(string token)
        {
            AccessCRMForWechat.AccessMiddleDB m_accessMiddleDB = new AccessCRMForWechat.AccessMiddleDB();

            EntityCollection ecLetters = ToBeSentEntities(token);
            if (ecLetters != null)
            {
                foreach (Entity letter in ecLetters.Entities)
                {
                    letter["new_sendstate"] = new OptionSetValue(100000001);
                    orgService.Update(letter);

                    #region 向微信中间表写入待发送任务日志

                    //System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding();
                    //Service1Client client = new Service1Client(binding, new System.ServiceModel.EndpointAddress(@"http://115.28.81.79:7001/mds/Service1.svc"));
                    ////throw new Exception(letter["new_token"].ToString() + ":" + entityAccount.Id);

                    //string strResult = client.CreateLog("letter", "3", letter.Id.ToString(), "", "True", letter["new_token"].ToString());
                    m_accessMiddleDB.CreateLog("letter", "3", letter.Id.ToString(), "", "True", letter["new_token"].ToString());

                    #endregion

                }
            }

            EntityCollection ecCampaignActivities = ToBeSentEntities_Group(token);
            if (ecCampaignActivities != null)
            {
                foreach (Entity campaignactivity in ecCampaignActivities.Entities)
                {
                    campaignactivity["new_sendstate"] = new OptionSetValue(100000001);
                    orgService.Update(campaignactivity);

                    #region 向微信中间表写入待发送任务日志

                    //System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding();
                    //Service1Client client = new Service1Client(binding, new System.ServiceModel.EndpointAddress("http://115.28.81.79:7001/mds/Service1.svc"));

                    //string strResult = client.CreateLog("lettergroup", "3", campaignactivity.Id.ToString(), "", "True", campaignactivity["new_token"].ToString());

                    m_accessMiddleDB.CreateLog("lettergroup", "3", campaignactivity.Id.ToString(), "", "True", campaignactivity["new_token"].ToString());

                    #endregion

                }
            }
        }

        public EntityCollection ToBeSentEntities(string token)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "letter",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_token"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_sendstate",
                                Operator = ConditionOperator.Equal,
                                Values = { 100000000 }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values={token}
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);
            if (ECReturn.Entities.Count > 0)
            {
                return ECReturn;
            }
            else
                return null;
        }

        public EntityCollection ToBeSentEntities_Group(string token)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "campaignactivity",
                //ColumnSet = new ColumnSet(true),
                ColumnSet = new ColumnSet("new_token"),
                Criteria = new FilterExpression
                {
                    Conditions =
                        {
                            new ConditionExpression
                            {
                                AttributeName = "new_sendstate",
                                Operator = ConditionOperator.Equal,
                                Values = { 100000000 }
                            },
                            new ConditionExpression
                            {
                                AttributeName = "new_token",
                                Operator = ConditionOperator.Equal,
                                Values={token}
                            }
                        }
                }
            };

            EntityCollection ECReturn = orgService.RetrieveMultiple(query);
            if (ECReturn.Entities.Count > 0)
            {
                return ECReturn;
            }
            else
                return null;
        }

        #endregion

        public Entity GetLetter(string LetterGUID)
        {
            Entity letter = orgService.Retrieve("letter", new Guid(LetterGUID), 
                new ColumnSet("new_token", "from", "new_openid", "new_imgsourceid"));

            return letter;
        }

        /// <summary>
        /// 更新群发微信发送状态
        /// </summary>
        /// <param name="letter">letter实体或campaignactivity实体</param>
        /// <param name="statevalue"></param>
        public void UpdateWASendState(Entity letter, int statevalue)
        {
            letter["new_sendstate"] = new OptionSetValue(statevalue);
            orgService.Update(letter);
        }

        public Entity GetCampaignActivity(string CampaignActivityGUID)
        {
            Entity campaignactivity = orgService.Retrieve("campaignactivity", new Guid(CampaignActivityGUID),
                new ColumnSet("new_token", "new_imgsourceid", "new_cusgroupsourceid"));

            return campaignactivity;
        }

    }
}
