using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Collections;
using System.Configuration;
using MySql.Data;
using MySql.Data.MySqlClient;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Crm;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel;
using System.ServiceModel.Description;
using DataAccess.dbConnect;


namespace CRMToWechat
{
    public partial class Form1 : Form
    {
        AccessCRMForWechat.AccessCRM m_accessCrm;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Execute();

            //string constr = ConfigurationManager.AppSettings["con_mysql"].ToString();

            ////MessageBox.Show(constr);

            //MySqlConnection mycon = new MySqlConnection(constr);

            //mycon.Open();

            ////MySqlCommand mycmd = new MySqlCommand("insert into transferrull(entityname,transferflag,updatemethod) values('account',1,2)", mycon);

            ////DataTable dt=
            //MySqlDataReader reader = (new MySqlCommand("select * from intcrm_transferrull", mycon)).ExecuteReader();
            //while (reader.Read())
            //{
            //    MessageBox.Show(reader["entityname"] + " -- " + reader["transferflag"] + " -- " + reader[3]);
            //}


            ////if (mycmd.ExecuteNonQuery() > 0)
            //{


            //}


            //mycon.Close();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int iInterval = 60;

            lbCurrentInterval.Text = DataAccess.IniConfig.GetValue("timer_crm", "interval");
            int.TryParse(lbCurrentInterval.Text, out iInterval);

            if (iInterval == 0) iInterval = 60;
            timer1.Interval = iInterval * 1000;

            lbServiceState.BackColor = Color.Lime;
            lbServiceState.Text = "已启动";
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            Execute();
            if (lbServiceState.BackColor.Name == "Lime") timer1.Start();
        }

        private void Execute()
        {
            try
            {
                DataAccess.dbConnect.ConnectionPool_mysql.ConnectionString = ConfigurationManager.AppSettings["con_mysql"].ToString();
                DataAccess.dbConnect.ConnectionPool_mysql.OpenConnecion();

                DataTable dtwx_crm = DataAccess.dbConnect.ConnectionPool_mysql.GetQuery("select * from intcrm_wxuser");

                foreach (DataRow dr in dtwx_crm.Rows)  //遍历公众号
                {
                    m_accessCrm = new AccessCRMForWechat.AccessCRM(dr["CRMUri"].ToString(), dr["crmuser"].ToString(), dr["crmpassword"].ToString());

                    DataTable dttransferfull = TransferRull( "1", dr["wxtoken"].ToString());   //获取启用状态的传输规则

                    foreach (DataRow drtransferfull in dttransferfull.Rows)   //遍历传输规则
                    {
                        switch (drtransferfull["entityname"].ToString())
                        {
                            //case "account":
                            //    FollowersToCRM(drtransferfull["token"].ToString());   //将关注者更新至CRM
                            //    break;
                            //case "img":
                            //    IMGToCRM( drtransferfull["token"].ToString());   //将图文信息更新至CRM
                            //    break;
                            //case "wxuser":
                            //    WxuserToCRM( drtransferfull["token"].ToString());   //将公众号更新至CRM
                            //    break;
                            //case "wxuseraccounts":
                            //    WxusersToCRM( drtransferfull["token"].ToString());   //将公众号管理员更新至CRM
                            //    break;
                            //case "action":
                            //    ActionToCRM( drtransferfull["token"].ToString());   //将关注者行为记录更新至CRM
                            //    break;
                            //case "order":
                            //    OrderToCRM( drtransferfull["token"].ToString());   //将订单记录更新至CRM
                            //    break;
                            //case "orderlist":
                            //    OrderListToCRM( drtransferfull["token"].ToString());   //将订单明细记录更新至CRM
                            //    break;
                            //case "enroll":
                            //    EnrollToCRM( drtransferfull["token"].ToString());   //将报名记录更新至CRM
                            //    break;
                            //case "associator":
                            //    AssociatorToCRM( drtransferfull["token"].ToString());   //将会员卡记录更新至CRM
                            //    break;
                            //case "scoredetail":
                            //    MemberScoreDetailToCRM( drtransferfull["token"].ToString());   //将积分明细更新至CRM
                            //    break;
                            //case "score":
                            //    MemberScoreToCRM(drtransferfull["token"].ToString());   //将会员积分更新至CRM
                            //    break;
                            //case "prd2prd":
                            //    prd2prdToCRM(drtransferfull["token"].ToString());   //将商品关联记录更新至CRM
                            //    break;
                            //case "prd2img":
                            //    prd2imgToCRM(drtransferfull["token"].ToString());   //将商品图文关联记录更新至CRM
                            //    break;
                            case "letter":
                                LetterToWechat(drtransferfull["token"].ToString());   //向关注者发送微信
                                break;
                            default:
                                break;
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), ex.Message);
            }
        }

        private void LetterToWechat(string token)
        {
            try
            {
                //string strSendResult;

                Entity letter = null;
                DataTable dttransferlog = TransferLog("letter", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                break;
                            case "2":   //修改
                                break;
                            case "1":   //删除
                                break;
                            case "3":   //
                                letter = m_accessCrm.GetLetter(drtransferlog["crmrecordid"].ToString());
                                if (letter!= null)
                                {
                                    //调用微信平台 Webservice 发送消息    //strSendResult = 
                                    //m_accessCrm.WritePrd2img(letter.Rows[0]["token"].ToString(), letter.Rows[0]["productid"].ToString(), letter.Rows[0]["imgid"].ToString());
                                    //WriteBackLog(drtransferlog["id"].ToString(), strSendResult, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "发送微信消息错误：源数据id=" + drtransferlog["crmrecordid"].ToString() + "不存在！");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "发送微信消息错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "发送微信消息错误：" + ex.Message);
            }
            finally
            {
            }
        }

        #region 商品图文关联记录至CRM
        private void prd2imgToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable prd2img = null;
                DataTable dttransferlog = TransferLog("prd2img", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                prd2img = Prd2img(drtransferlog["wxrecordid"].ToString());
                                if (prd2img.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建商品图文关联记录方法
                                    strBackGUID = m_accessCrm.WritePrd2img(prd2img.Rows[0]["token"].ToString(), prd2img.Rows[0]["productid"].ToString(), prd2img.Rows[0]["imgid"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新商品图文关联记录至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "2":   //修改
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除商品图文关联记录方法
                                strBackGUID = FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopPrd2img(strBackGUID);
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除商品图文关联记录更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品图文关联记录更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品图文关联记录更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Prd2img(string p)
        {
            string strSql = "select * from wx_prd2img where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion


        #region 商品关联记录至CRM
        private void prd2prdToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable prd2prd = null;
                DataTable dttransferlog = TransferLog("prd2prd", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                prd2prd = Prd2prd(drtransferlog["wxrecordid"].ToString());
                                if (prd2prd.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建商品关联记录方法
                                    strBackGUID = m_accessCrm.WritePrd2prd(prd2prd.Rows[0]["token"].ToString(), prd2prd.Rows[0]["productid1"].ToString(), prd2prd.Rows[0]["productid2"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新商品关联记录至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "2":   //修改
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除商品关联记录方法
                                strBackGUID = FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopPrd2prd(strBackGUID);
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除商品关联记录更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品关联记录更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品关联记录更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Prd2prd(string p)
        {
            string strSql = "select * from wx_prd2prd where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 会员卡信息至CRM
        private void AssociatorToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable associator = null;
                DataTable dttransferlog = TransferLog( "associator", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                associator = Associator( drtransferlog["wxrecordid"].ToString());
                                if (associator.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建会员卡信息方法
                                    strBackGUID = m_accessCrm.WriteAssociator(associator.Rows[0]["wecha_id"].ToString(),
                                        associator.Rows[0]["token"].ToString(), associator.Rows[0]["number"].ToString(), "1", "", "");
                                    WriteBackLog( drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新会员卡信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除会员卡信息方法
                                strBackGUID = FindIDForDelete( drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopAssociator(strBackGUID);
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除会员卡信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将会员卡信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将会员卡信息更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Associator(string p)
        {
            string strSql = "select * from wx_member_card_create where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 报名信息至CRM
        private void EnrollToCRM( string token)
        {
            try
            {
                string strBackGUID;
                DataTable enroll = null;
                DataTable dttransferlog = TransferLog( "enroll", token);
                foreach(DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                enroll = Enroll(drtransferlog["wxrecordid"].ToString());
                                if (enroll.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建报名信息方法
                                    strBackGUID = FindIDForUpdate(drtransferlog);
                                    strBackGUID = m_accessCrm.WriteEnroll(strBackGUID,
                                        enroll.Rows[0]["wecha_id"].ToString(), enroll.Rows[0]["token"].ToString(), enroll.Rows[0]["name"].ToString(),
                                        enroll.Rows[0]["intro"].ToString(), enroll.Rows[0]["enrolltime"].ToString(), enroll.Rows[0]["status"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新报名信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除报名信息方法
                                strBackGUID = FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopEnroll(strBackGUID);
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除报名信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将报名信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    { }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将报名信息更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }

        }
        private DataTable Enroll(string p)
        {
            string strSql = "select v.*,en.token,en.name,en.intro,FROM_UNIXTIME(v.time,'%Y-%m-%d %T') as enrolltime from wx_enroll_value v LEFT JOIN wx_enroll en on en.id=v.formid where v.id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 关注者行为记录至CRM
        private void ActionToCRM( string token)
        {
            try
            {
                string strBackGUID;
                DataTable action = null;

                DataTable dttransferlog = TransferLog("action", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                action = Action( drtransferlog["wxrecordid"].ToString());
                                if (action.Rows.Count>0)
                                {
                                    //To do :调用CRM Webservice 创建公众号方法
                                    strBackGUID = m_accessCrm.WriteAction(action.Rows[0]["openid"].ToString(), action.Rows[0]["token"].ToString(),
                                        action.Rows[0]["objectid"].ToString(),
                                        action.Rows[0]["createdtime"].ToString(),
                                        action.Rows[0]["actiontype"].ToString(),
                                        action.Rows[0]["actionobject"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建关注者行为至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除关注者行为方法
                                break;
                            case "2":   //修改
                                //action = Action(mycon.Clone(), transferlog["wxrecordid"].ToString());
                                //if (action.Read())
                                //{
                                //    ;
                                //    //To do :调用CRM Webservice 更新关注者行为方法
                                //}
                                //else
                                //{
                                //    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "更新关注者行为至CRM错误：源数据id=" + transferlog["wxrecordid"].ToString() + "不存在！");
                                //}
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者行为更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog( drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者行为更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Action( string p)
        {
            string strSql = "select * from wx_action where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 公众号至CRM
        private void WxuserToCRM( string token)
        {
            try
            {
                string strBackGUID;

                DataTable wxuser = null;
                DataTable dttransferlog = TransferLog( "wxuser", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                wxuser = Wxuser( drtransferlog["wxrecordid"].ToString());
                                if (wxuser.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建公众号方法
                                    strBackGUID = m_accessCrm.WriteWechatAccount(wxuser.Rows[0]["token"].ToString(), wxuser.Rows[0]["wxname"].ToString());
                                    WriteBackLog( drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新公众号至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除公众号方法
                                strBackGUID=FindIDForDelete( drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopWechatAccount(strBackGUID);
                                    WriteBackLog( drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除公众号更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog( drtransferlog["id"].ToString(), "", "", "1");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog( drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable Wxuser( string p)
        {
            string strSql = "select * from wx_wxuser where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }

        #endregion

        #region 公众号帐户至CRM
        private void WxusersToCRM( string token)
        {
            try
            {
                string strBackGUID;

                DataTable wxuser = null;
                DataTable dttransferlog = TransferLog("wxuseraccounts", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                wxuser = Wxusers( drtransferlog["wxrecordid"].ToString(),token);
                                if (wxuser.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建公众号管理员方法
                                    strBackGUID = m_accessCrm.WriteWechatAccountAdmin(wxuser.Rows[0]["username"].ToString(),
                                        wxuser.Rows[0]["parentid"].ToString(),
                                        wxuser.Rows[0]["email"].ToString(),
                                        wxuser.Rows[0]["parentid"].ToString() == "0",
                                        wxuser.Rows[0]["id"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新公众号管理员至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除公众号管理员方法
                                strBackGUID=FindIDForDelete( drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopWechatAccountAdmin(strBackGUID);
                                    WriteBackLog( drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除公众号管理员更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog( drtransferlog["id"].ToString(), "", "", "1");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号管理员更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog( drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号管理员更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable Wxusers( string p,string token)
        {
            string strSql = @"select * from ((SELECT ur.*,wx.token FROM wx_users ur left join wx_wxuser wx on wx.uid=ur.id where wx.token is not NULL)
UNION
(SELECT u1.*,wx.token FROM wx_users u1 left join wx_users u2 on u1.parentId=u2.id left join wx_wxuser wx on wx.uid=u2.id where u1.parentId>0)) a
where a.token='"+token+"' and a.id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }

        #endregion

        #region 图文信息至CRM
        private void IMGToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable img = null;
                DataTable dttransferlog = TransferLog("img", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                img = Img(drtransferlog["wxrecordid"].ToString());
                                if (img.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建图文信息方法
                                    strBackGUID = m_accessCrm.WriteIMG(img.Rows[0]["id"].ToString(), img.Rows[0]["token"].ToString(), img.Rows[0]["title"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新图文信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除图文信息方法
                                strBackGUID = FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopIMG(strBackGUID);
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除图文信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将图文信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将图文信息更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Img(string p)
        {
            string strSql = "select * from wx_img where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 关注者至CRM
        /// <summary>
        ///  将关注者更新至CRM
        /// </summary>
        /// <param name="mycon">数据库链接</param>
        /// <param name="token"></param>
        private void FollowersToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable dtaccount = null;
                DataTable dttransferlog = TransferLog("account", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    //逐条执行数据同步
                    try
                    {
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                dtaccount = Account(drtransferlog["wxrecordid"].ToString());
                                if (dtaccount.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建更新客户方法
                                    strBackGUID = m_accessCrm.WriteAccount(dtaccount.Rows[0]["openid"].ToString(), dtaccount.Rows[0]["token"].ToString(), "", "",
                                        "", dtaccount.Rows[0]["NickName"].ToString(),
                                        dtaccount.Rows[0]["status"].ToString() == "关注中" ? "1" : "0",
                                        dtaccount.Rows[0]["_attentiontime"].ToString(),
                                        dtaccount.Rows[0]["_cattentiontime"].ToString(),
                                        dtaccount.Rows[0]["sex"].ToString(),
                                        dtaccount.Rows[0]["country"].ToString(),
                                        dtaccount.Rows[0]["province"].ToString(),
                                        dtaccount.Rows[0]["city"].ToString(),
                                        "");
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建关注者至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //调用CRM Webservice 删除客户方法
                                strBackGUID=FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopAccount(strBackGUID);
                                    WriteBackLog( drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除关注者更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog( drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable Account(string p)
        {
            string strSql = "select *, case when AttentionTime ='' then '' else FROM_UNIXTIME(AttentionTime,'%Y-%m-%d %T') end as _AttentionTime,case when cAttentionTime ='' then '' else FROM_UNIXTIME(cAttentionTime,'%Y-%m-%d %T') end as _cAttentionTime from wx_customer where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 订单至CRM
        /// <summary>
        ///  将订单更新至CRM
        /// </summary>
        /// <param name="mycon">数据库链接</param>
        /// <param name="token"></param>
        private void OrderToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable order = null;
                DataTable dttransferlog = TransferLog( "order", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    //逐条执行数据同步
                    try
                    {
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                order = Order( drtransferlog["wxrecordid"].ToString());
                                if (order.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建更新客户方法
                                    strBackGUID = m_accessCrm.WriteOrder(order.Rows[0]["wecha_id"].ToString(), order.Rows[0]["token"].ToString(),
                                        order.Rows[0]["price"].ToString(),
                                        order.Rows[0]["_Time"].ToString(),
                                        order.Rows[0]["truename"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建订单至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "2":   //修改
                            case "1":   //删除
                                //调用CRM Webservice 删除客户方法
                                //strBackGUID = FindIDForDelete(mycon.Clone(), transferlog);
                                //if (strBackGUID.Length > 0)
                                //{
                                //    m_accessCrm.StopOrder(strBackGUID);
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), strBackGUID, "", "1");
                                //}
                                //else
                                //{
                                //    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除订单更新至CRM错误：不存在CRM对应GUID，日志ID=" + transferlog["id"].ToString() + "。");
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), "", "", "1");
                                //}

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将订单更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将订单更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable Order( string p)
        {
            string strSql = "select *, case when Time ='' then '' else FROM_UNIXTIME(Time,'%Y-%m-%d %T') end as _Time from wx_product_cart where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 订单明细至CRM
        /// <summary>
        ///  将订单更新至CRM
        /// </summary>
        /// <param name="mycon">数据库链接</param>
        /// <param name="token"></param>
        private void OrderListToCRM( string token)
        {
            try
            {
                string strBackGUID;

                DataTable orderlist = null;
                DataTable dttransferlog = TransferLog( "orderlist", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    //逐条执行数据同步
                    try
                    {
                        WriteStartLog( drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                orderlist = OrderList(drtransferlog["wxrecordid"].ToString());
                                if (orderlist.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建更新订单明细方法
                                    strBackGUID = m_accessCrm.WriteOrderDetail(orderlist.Rows[0]["crmrecordid"].ToString(), orderlist.Rows[0]["token"].ToString(),
                                        orderlist.Rows[0]["productid"].ToString(),
                                        orderlist.Rows[0]["total"].ToString(),
                                        orderlist.Rows[0]["price"].ToString(),
                                        orderlist.Rows[0]["amount"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建订单明细至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "2":   //修改
                            case "1":   //删除
                                //调用CRM Webservice 删除订单明细方法
                                //strBackGUID = FindIDForDelete(mycon.Clone(), transferlog);
                                //if (strBackGUID.Length > 0)
                                //{
                                //    m_accessCrm.StopOrder(strBackGUID);
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), strBackGUID, "", "1");
                                //}
                                //else
                                //{
                                //    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除订单更新至CRM错误：不存在CRM对应GUID，日志ID=" + transferlog["id"].ToString() + "。");
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), "", "", "1");
                                //}

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将订单明细更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将订单明细更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable OrderList(string p)
        {
            string strSql = "select cl.*,price*total as amount,log.crmrecordid from wx_product_cart_list cl LEFT JOIN (select DISTINCT wxrecordid,crmrecordid from intcrm_transferlog where entityname='order' and issuccess=1) log on cl.cartid=log.wxrecordid where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 会员积分明细至CRM
        /// <summary>
        ///  将积分明细更新至CRM
        /// </summary>
        /// <param name="token"></param>
        private void MemberScoreDetailToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable memberscore = null;
                DataTable dttransferlog = TransferLog("scoredetail", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    //逐条执行数据同步
                    try
                    {
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                memberscore = MemberScoreDetail(drtransferlog["wxrecordid"].ToString());
                                if (memberscore.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建更新订单明细方法
                                    strBackGUID = m_accessCrm.WriteScoreDetail(memberscore.Rows[0]["wecha_id"].ToString(), memberscore.Rows[0]["token"].ToString(),
                                        "",
                                        memberscore.Rows[0]["type"].ToString(),
                                        memberscore.Rows[0]["score_type"].ToString(),
                                        memberscore.Rows[0]["expense"].ToString(),
                                        memberscore.Rows[0]["stime"].ToString(),"");
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建积分明细至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "2":   //修改
                            case "1":   //删除
                                //调用CRM Webservice 删除订单明细方法
                                //strBackGUID = FindIDForDelete(mycon.Clone(), transferlog);
                                //if (strBackGUID.Length > 0)
                                //{
                                //    m_accessCrm.StopOrder(strBackGUID);
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), strBackGUID, "", "1");
                                //}
                                //else
                                //{
                                //    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除订单更新至CRM错误：不存在CRM对应GUID，日志ID=" + transferlog["id"].ToString() + "。");
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), "", "", "1");
                                //}

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将积分明细更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将积分明细更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable MemberScoreDetail(string p)
        {
            string strSql = "select score.*,FROM_UNIXTIME(score.sign_time,'%Y-%m-%d %T') as stime from wx_member_card_sign score where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 会员总积分至CRM
        /// <summary>
        ///  将总积分更新至CRM
        /// </summary>
        /// <param name="token"></param>
        private void MemberScoreToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable userinfo = null;
                DataTable dttransferlog = TransferLog("score", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    //逐条执行数据同步
                    try
                    {
                        WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "2":   //修改
                                userinfo = MemberScore(drtransferlog["wxrecordid"].ToString());
                                if (userinfo.Rows.Count > 0)
                                {
                                    //调用CRM Webservice建更新方法
                                    strBackGUID = m_accessCrm.WriteScore(userinfo.Rows[0]["wecha_id"].ToString(), userinfo.Rows[0]["token"].ToString(),
                                        "",
                                        userinfo.Rows[0]["tel"].ToString(),
                                        "1",
                                        userinfo.Rows[0]["truename"].ToString(),
                                        userinfo.Rows[0]["total_score"].ToString());
                                    WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "积分至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "0":   //增加
                            case "1":   //删除
                                //调用CRM Webservice 删除订单明细方法
                                //strBackGUID = FindIDForDelete(mycon.Clone(), transferlog);
                                //if (strBackGUID.Length > 0)
                                //{
                                //    m_accessCrm.StopOrder(strBackGUID);
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), strBackGUID, "", "1");
                                //}
                                //else
                                //{
                                //    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除订单更新至CRM错误：不存在CRM对应GUID，日志ID=" + transferlog["id"].ToString() + "。");
                                //    WriteBackLog(mycon.Clone(), transferlog["id"].ToString(), "", "", "1");
                                //}

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将积分更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将积分更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }

        private DataTable MemberScore(string p)
        {
            string strSql = "select * from wx_userinfo where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        public DataTable TransferRull(string transferflag, string token)
        {
            try
            {
                string strSql = "select * from intcrm_transferrull where token='" + token + "'";

                if (transferflag.Length > 0)
                {
                    strSql += " and transferflag=" + transferflag;
                }
                DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

                return reader;
            }

            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "获取传输规则异常：" + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 取出待传输日志
        /// </summary>
        /// <param name="mycon"></param>
        /// <param name="entityname"></param>
        /// <param name="token"></param>
        /// <param name="direct">传输方向：0为微信->CRM，1为CRM->微信</param>
        /// <param name="isclosed"></param>
        /// <returns></returns>
        public DataTable TransferLog( string entityname, string token, string direct = "0", string isclosed = "0")
        {
            string strSql = "select * from intcrm_transferlog where token='" + token + "' and  (issuccess is NULL or issuccess=0) and direct=" + direct +
                " and isclosed=" + isclosed + " and entityname='" + entityname + "'";

            if (entityname.Length == 0)
            {
                throw new Exception("传输实体名不能为空！");
            }
            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }

        /// <summary>
        /// 获取公众号与CRM组织映射关系
        /// </summary>
        /// <param name="mycon"></param>
        /// <param name="state">启用状态，1：启用，0：停用</param>
        /// <returns></returns>
        public DataTable WXUserCRMOrg(MySqlConnection mycon, string state = "1")
        {
            //mycon.Open();
            DataTable dtTmp;
            string strSql = "select * from intcrm_wxuser";

            if (state.Length > 0)
            {
                strSql += " where state=" + state;
            }
            dtTmp = ConnectionPool_mysql.GetQuery(strSql);

            return dtTmp;
        }

        #region 日志操作
        public void WriteStartLog(string logID)
        {
            string strSql = "update intcrm_transferlog set transfertime='" + DateTime.Today.ToString("yyyy-MM-dd") + " " + DateTime.Now.ToLongTimeString() +
                "' where id=" + logID;
            ConnectionPool_mysql.UpdateQuery(strSql);

        }

        /// <summary>
        /// 执行任务后，回写Log
        /// </summary>
        /// <param name="mycon"></param>
        /// <param name="logID"></param>
        /// <param name="crmGUID">CRM中的记录ID，它与wxID参数只能其中一个有值</param>
        /// <param name="wxID">微信中的记录ID，它与crmGUID参数只能其中一个有值</param>
        /// <param name="IsSuccess">1 为成功 0为失败</param>
        public void WriteBackLog(string logID, string crmGUID, string wxID, string IsSuccess)
        {
            string strSql;
            if (IsSuccess == "1")
            {
                strSql = "update intcrm_transferlog set issuccess=1,isclosed=1,successtime='" +
                    DateTime.Today.ToString("yyyy-MM-dd") + " " + DateTime.Now.ToLongTimeString() + "'" +
                    (crmGUID.Length > 0 ? ", crmrecordid='" + crmGUID + "'" : "") +
                    (wxID.Length > 0 ? ", wxrecordid='" + wxID + "'" : "") +
                    " where id=" + logID;
            }
            else
            {
                strSql = "update intcrm_transferlog set issuccess=0 where id=" + logID;
            }
            ConnectionPool_mysql.UpdateQuery(strSql);
        }

        public void CloseLog(MySqlConnection mycon, string logID)
        {
            string strSql;
            mycon.Open();
            strSql = "update intcrm_transferlog set isclosed=1 where id=" + logID;
            MySqlCommand cmd = new MySqlCommand(strSql, mycon);
            cmd.ExecuteNonQuery();

            mycon.Close();
        }
        private string FindIDForDelete( DataRow transferlog)
        {
            string strSql, strReturn = "";
            DataTable tmpReader;
            if (transferlog["operatetype"].ToString() == "1")
            {
                if (transferlog["wxrecordid"].ToString().Length > 0)
                {
                    strSql = "select * from intcrm_transferlog where direct=0 and operatetype<>1 and wxrecordid='" +
                        transferlog["wxrecordid"].ToString() +
                        "' and entityname='" + transferlog["entityname"].ToString() +
                        "' and token='" + transferlog["token"].ToString() + "'";
                    tmpReader = ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count>0)
                    {
                        strReturn = tmpReader.Rows[0]["crmrecordid"].ToString();
                        return strReturn;
                    }
                    else
                    {
                        return strReturn;
                    }
                }
                if (transferlog["crmrecordid"].ToString().Length > 0)
                {
                    strSql = "select * from intcrm_transferlog where direct=1 and operatetype<>1 and crmrecordid='" + transferlog["crmrecordid"].ToString() + "'";
                    tmpReader = ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count>0)
                    {
                        strReturn = tmpReader.Rows[0]["wxrecordid"].ToString();
                        return strReturn;
                    }
                    else
                    {
                        return strReturn;
                    }
                }
            }
            return strReturn;
        }

        private string FindIDForUpdate( DataRow transferlog)
        {
            string strSql, strReturn = "";
            DataTable tmpReader;
            if (transferlog["operatetype"].ToString() == "2")
            {
                if (transferlog["wxrecordid"].ToString().Length > 0)
                {
                    strSql = "select * from intcrm_transferlog where direct=0 and operatetype<>2 and wxrecordid='" + transferlog["wxrecordid"].ToString() + "'";
                    tmpReader = ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count>0)
                    {
                        strReturn = tmpReader.Rows[0]["crmrecordid"].ToString();
                        return strReturn;
                    }
                    else
                    {
                        return strReturn;
                    }
                }
                if (transferlog["crmrecordid"].ToString().Length > 0)
                {
                    strSql = "select * from intcrm_transferlog where direct=1 and operatetype<>2 and crmrecordid='" + transferlog["crmrecordid"].ToString() + "'";
                    tmpReader =ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count>0)
                    {
                        strReturn = tmpReader.Rows[0]["wxrecordid"].ToString();
                        return strReturn;
                    }
                    else
                    {
                        return strReturn;
                    }
                }
            }
            return strReturn;
        }

        #endregion

        private void btnSave_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            int iInterval = 60;
            int.TryParse(txtInterval.Text,out iInterval);
            DataAccess.IniConfig.WriteValue("timer_crm", "interval", iInterval.ToString());

            lbCurrentInterval.Text = DataAccess.IniConfig.GetValue("timer_crm", "interval");
            timer1.Interval = Convert.ToInt32(lbCurrentInterval.Text) * 1000;

            if (lbServiceState.BackColor.Name == "Lime") timer1.Start();
        }

        private void lbServiceState_Click(object sender, EventArgs e)
        {
            if (lbServiceState.BackColor.Name == "Lime")
            {
                lbServiceState.BackColor = Color.Red;
                lbServiceState.Text = "已停止";
                timer1.Stop();
            }
            else
            {
                lbServiceState.BackColor = Color.Lime;
                lbServiceState.Text = "已启动";
                timer1.Start();
            }
        }

    }
}
