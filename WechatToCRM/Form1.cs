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

using DataAccess.dbConnect;


namespace WechatToCRM
{
    public partial class Form1 : Form
    {
        AccessCRMForWechat.AccessCRM m_accessCrm;
        AccessCRMForWechat.AccessMiddleDB m_accessMiddleDB;
        int m_iRullEditFlag = 2, m_iCrmWxEditFlag = 2;  //0:新增  1:删除  2:修改
        int m_iCurrentID_Rull = 0;    //记录传输规则当前ID
        string m_strCurrentWxtoken_CrmWx = "";    //记录系统映射关系当前主键值

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
            // TODO: 这行代码将数据加载到表“mytestDataSet1.intcrm_transferrull”中。您可以根据需要移动或删除它。
            this.intcrm_transferrullTableAdapter.Fill(this.mytestDataSet1.intcrm_transferrull);
            // TODO: 这行代码将数据加载到表“mytestDataSet.intcrm_wxuser”中。您可以根据需要移动或删除它。
            this.intcrm_wxuserTableAdapter.Fill(this.mytestDataSet.intcrm_wxuser);
            this.intcrm_transferlogTableAdapter.Fill(this.mytestDataSet1.intcrm_transferlog);


            int iInterval = 60;

            lbCurrentInterval.Text = DataAccess.IniConfig.GetValue("timer", "interval");
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

                m_accessMiddleDB = new AccessCRMForWechat.AccessMiddleDB();

                DataTable dtwx_crm = DataAccess.dbConnect.ConnectionPool_mysql.GetQuery("select * from intcrm_wxuser");

                foreach (DataRow dr in dtwx_crm.Rows)  //遍历公众号
                {
                    m_accessCrm = new AccessCRMForWechat.AccessCRM(dr["CRMUri"].ToString(), dr["crmuser"].ToString(), dr["crmpassword"].ToString());

                    DataTable dttransferfull = TransferRull( "1", dr["wxtoken"].ToString());   //获取启用状态的传输规则

                    foreach (DataRow drtransferfull in dttransferfull.Rows)   //遍历传输规则
                    {
                        switch (drtransferfull["entityname"].ToString())
                        {
                            case "cusgroup":
                                CusGroupToCRM(drtransferfull["token"].ToString());   //将关注者分组更新至CRM
                                break;
                            case "account":
                                FollowersToCRM(drtransferfull["token"].ToString());   //将关注者更新至CRM
                                break;
                            case "img":
                                IMGToCRM( drtransferfull["token"].ToString());   //将图文信息更新至CRM
                                break;
                            case "lottery":
                                LotteryToCRM( drtransferfull["token"].ToString());   //将微信互动更新至CRM
                                break;
                            case "company":
                                ShopToCRM(drtransferfull["token"].ToString());   //将门店关联记录更新至CRM
                                break;
                            case "wxuser":
                                WxuserToCRM( drtransferfull["token"].ToString());   //将公众号更新至CRM
                                break;
                            case "wxuseraccounts":
                                WxusersToCRM( drtransferfull["token"].ToString());   //将公众号管理员更新至CRM
                                break;
                            case "action":
                                ActionToCRM( drtransferfull["token"].ToString());   //将关注者行为记录更新至CRM
                                break;
                            case "order":
                                OrderToCRM( drtransferfull["token"].ToString());   //将订单记录更新至CRM
                                break;
                            case "orderlist":
                                OrderListToCRM( drtransferfull["token"].ToString());   //将订单明细记录更新至CRM
                                break;
                            case "enroll":
                                EnrollToCRM( drtransferfull["token"].ToString());   //将报名记录更新至CRM
                                break;
                            case "associator":
                                AssociatorToCRM( drtransferfull["token"].ToString());   //将会员卡记录更新至CRM
                                break;
                            case "scoredetail":
                                MemberScoreDetailToCRM( drtransferfull["token"].ToString());   //将积分明细更新至CRM
                                break;
                            case "score":
                                MemberScoreToCRM(drtransferfull["token"].ToString());   //将会员积分更新至CRM
                                break;
                            case "prd2prd":
                                prd2prdToCRM(drtransferfull["token"].ToString());   //将商品关联记录更新至CRM
                                break;
                            case "prd2img":
                                prd2imgToCRM(drtransferfull["token"].ToString());   //将商品图文关联记录更新至CRM
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

        #region 关注者分组信息至CRM
        private void CusGroupToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable cusgroup = null;
                DataTable dttransferlog = TransferLog("cusgroup", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                cusgroup = CusGroup(drtransferlog["wxrecordid"].ToString());
                                if (cusgroup.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建关注者分组信息方法
                                    strBackGUID = m_accessCrm.WriteCusGroup(cusgroup.Rows[0]["id"].ToString(), cusgroup.Rows[0]["token"].ToString(), cusgroup.Rows[0]["name"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新关注者分组信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除关注者分组信息方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopCusGroup(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除关注者分组信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者分组信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者分组信息更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable CusGroup(string p)
        {
            string strSql = "select * from wx_cusgroup where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

        #region 微信互动至CRM
        private void LotteryToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable lottery = null;
                DataTable dttransferlog = TransferLog("lottery", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                lottery = Lottery(drtransferlog["wxrecordid"].ToString());
                                if (lottery.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建微信互动信息方法
                                    strBackGUID = m_accessCrm.WriteLottery(lottery.Rows[0]["id"].ToString(),
                                        lottery.Rows[0]["token"].ToString(), lottery.Rows[0]["name"].ToString(),
                                        lottery.Rows[0]["type"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新微信互动至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除微信互动方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopLottery(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除微信互动更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将微信互动更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将微信互动更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Lottery(string p)
        {
            string strSql = "select * from wx_lottery where id=" + p;

            DataTable reader = ConnectionPool_mysql.GetQuery(strSql);

            return reader;
        }
        #endregion

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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                prd2img = Prd2img(drtransferlog["wxrecordid"].ToString());
                                if (prd2img.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建商品图文关联记录方法
                                    strBackGUID = m_accessCrm.WritePrd2img(prd2img.Rows[0]["token"].ToString(), prd2img.Rows[0]["productid"].ToString(), prd2img.Rows[0]["imgid"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopPrd2img(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除商品图文关联记录更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品图文关联记录更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                                prd2prd = Prd2prd(drtransferlog["wxrecordid"].ToString());
                                if (prd2prd.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建商品关联记录方法
                                    strBackGUID = m_accessCrm.WritePrd2prd(prd2prd.Rows[0]["token"].ToString(), prd2prd.Rows[0]["productid1"].ToString(), prd2prd.Rows[0]["productid2"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopPrd2prd(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除商品关联记录更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将商品关联记录更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新会员卡信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除会员卡信息方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopAssociator(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除会员卡信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将会员卡信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                enroll = Enroll(drtransferlog["wxrecordid"].ToString());
                                if (enroll.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建报名信息方法
                                    strBackGUID = m_accessMiddleDB.FindIDForUpdate(drtransferlog);
                                    strBackGUID = m_accessCrm.WriteEnroll(strBackGUID,
                                        enroll.Rows[0]["wecha_id"].ToString(), enroll.Rows[0]["token"].ToString(), enroll.Rows[0]["name"].ToString(),
                                        enroll.Rows[0]["intro"].ToString(), enroll.Rows[0]["enrolltime"].ToString(), enroll.Rows[0]["status"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新报名信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除报名信息方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopEnroll(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除报名信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将报名信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                wxuser = Wxuser( drtransferlog["wxrecordid"].ToString());
                                if (wxuser.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建公众号方法
                                    strBackGUID = m_accessCrm.WriteWechatAccount(wxuser.Rows[0]["token"].ToString(), wxuser.Rows[0]["wxname"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新公众号至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除公众号方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopWechatAccount(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除公众号更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新公众号管理员至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除公众号管理员方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopWechatAccountAdmin(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除公众号管理员更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将公众号管理员更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
        private void IMGToCRM( string token)
        {
            try
            {
                string strBackGUID;

                DataTable img = null;
                DataTable dttransferlog = TransferLog( "img", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                img = Img( drtransferlog["wxrecordid"].ToString());
                                if (img.Rows.Count>0)
                                {
                                    //调用CRM Webservice 创建图文信息方法
                                    strBackGUID = m_accessCrm.WriteIMG(img.Rows[0]["id"].ToString(), img.Rows[0]["token"].ToString(), img.Rows[0]["title"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新图文信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除图文信息方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopIMG(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除图文信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将图文信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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

        #region 门店信息至CRM
        private void ShopToCRM(string token)
        {
            try
            {
                string strBackGUID;

                DataTable shop = null;
                DataTable dttransferlog = TransferLog("company", token);
                foreach (DataRow drtransferlog in dttransferlog.Rows)
                {
                    try
                    {
                        //逐条执行数据同步
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

                        switch (drtransferlog["operatetype"].ToString())
                        {
                            case "0":   //增加
                            case "2":   //修改
                                shop = Shop(drtransferlog["wxrecordid"].ToString());
                                if (shop.Rows.Count > 0)
                                {
                                    //调用CRM Webservice 创建门店信息方法
                                    strBackGUID = m_accessCrm.WriteShop(shop.Rows[0]["id"].ToString(), shop.Rows[0]["token"].ToString(),
                                        shop.Rows[0]["name"].ToString(), shop.Rows[0]["address"].ToString(),Convert.ToBoolean( shop.Rows[0]["isbranch"]));
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建更新门店信息至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //To do :调用CRM Webservice 删除门店信息方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopShop(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除门店信息更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将门店信息更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }

            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将门店信息更新至CRM错误：" + ex.Message);
            }
            finally
            {
            }
        }
        private DataTable Shop(string p)
        {
            string strSql = "select * from wx_company where id=" + p;

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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                        "",
                                        dtaccount.Rows[0]["groupid"].ToString());
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "创建关注者至CRM错误：源数据id=" + drtransferlog["wxrecordid"].ToString() + "不存在！");
                                }
                                break;
                            case "1":   //删除
                                //调用CRM Webservice 删除客户方法
                                strBackGUID = m_accessMiddleDB.FindIDForDelete(drtransferlog);
                                if (strBackGUID.Length > 0)
                                {
                                    m_accessCrm.StopAccount(strBackGUID);
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "删除关注者更新至CRM错误：不存在CRM对应GUID，日志ID=" + drtransferlog["id"].ToString() + "。");
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "1");
                                }

                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "将关注者更新至CRM错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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
                        m_accessMiddleDB.WriteStartLog(drtransferlog["id"].ToString());

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
                                    m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), strBackGUID, "", "1");
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
                        m_accessMiddleDB.WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
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


        private void btnSave_Click(object sender, EventArgs e)
        {
            timer1.Stop();

            int iInterval = 60;
            int.TryParse(txtInterval.Text,out iInterval);
            DataAccess.IniConfig.WriteValue("timer", "interval", iInterval.ToString());

            lbCurrentInterval.Text = DataAccess.IniConfig.GetValue("timer", "interval");
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

        #region 规则表_数据操作事件

        private void dataGridView2_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                (this.mytestDataSet1.intcrm_transferrull.Rows.Count - 1) >= e.RowIndex &&
                m_iRullEditFlag==2)
            {
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.Open();

                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@UpdateMethod"].Value = dataGridView2.Rows[e.RowIndex].Cells["updateMethodDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@EntityName"].Value = dataGridView2.Rows[e.RowIndex].Cells["entityNameDataGridViewTextBoxColumn"].Value;
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@TransferFlag"].Value = dataGridView2.Rows[e.RowIndex].Cells["transferFlagDataGridViewCheckBoxColumn"].Value;
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@token"].Value = dataGridView2.Rows[e.RowIndex].Cells["tokenDataGridViewTextBoxColumn"].Value;
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@Original_ID"].Value = this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["ID"];
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@Original_EntityName"].Value = this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["EntityName"];
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@Original_TransferFlag"].Value = this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["TransferFlag"];
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@Original_UpdateMethod"].Value = this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"];
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Parameters["@Original_token"].Value = this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["token"];

                int i = intcrm_transferrullTableAdapter.Adapter.UpdateCommand.ExecuteNonQuery();
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.Close();
                dataGridView2.Update();

                this.intcrm_transferrullTableAdapter.Fill(this.mytestDataSet1.intcrm_transferrull);
            }

            if (intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                (this.mytestDataSet1.intcrm_transferrull.Rows.Count) >= e.RowIndex &&
                m_iRullEditFlag == 0)
            {
                intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.Open();

                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);

                intcrm_transferrullTableAdapter.Adapter.InsertCommand.Parameters["@UpdateMethod"].Value = dataGridView2.Rows[e.RowIndex].Cells["updateMethodDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_transferrullTableAdapter.Adapter.InsertCommand.Parameters["@EntityName"].Value = dataGridView2.Rows[e.RowIndex].Cells["entityNameDataGridViewTextBoxColumn"].Value;
                intcrm_transferrullTableAdapter.Adapter.InsertCommand.Parameters["@TransferFlag"].Value = dataGridView2.Rows[e.RowIndex].Cells["transferFlagDataGridViewCheckBoxColumn"].Value;
                intcrm_transferrullTableAdapter.Adapter.InsertCommand.Parameters["@token"].Value = dataGridView2.Rows[e.RowIndex].Cells["tokenDataGridViewTextBoxColumn"].Value;

                int i = intcrm_transferrullTableAdapter.Adapter.InsertCommand.ExecuteNonQuery();
                intcrm_transferrullTableAdapter.Adapter.InsertCommand.Connection.Close();
                dataGridView2.Refresh();
                mytestDataSet1.intcrm_transferlog.Reset();

                this.intcrm_transferrullTableAdapter.Fill(this.mytestDataSet1.intcrm_transferrull);
            }
        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            m_iRullEditFlag = 2;
            if (dataGridView2.Rows[e.RowIndex].Cells["ID"].Value != null)
                m_iCurrentID_Rull = (int)dataGridView2.Rows[e.RowIndex].Cells["ID"].Value;
            else
                m_iCurrentID_Rull = 0;
        }

        private void dataGridView2_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            m_iRullEditFlag = 0;

        }

        private void dataGridView2_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            m_iRullEditFlag = 1;

            if (intcrm_transferrullTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
              // (this.mytestDataSet1.intcrm_transferrull.Rows.Count) >= e.RowIndex &&
               m_iRullEditFlag == 1 && m_iCurrentID_Rull != 0)
            {
                intcrm_transferrullTableAdapter.Adapter.DeleteCommand.Connection.Open();

                dataGridView2.CommitEdit(DataGridViewDataErrorContexts.Commit);

                intcrm_transferrullTableAdapter.Adapter.DeleteCommand.Parameters["@Original_ID"].Value = e.Row.Cells["ID"].Value; //this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]);

                int i = intcrm_transferrullTableAdapter.Adapter.DeleteCommand.ExecuteNonQuery();
                intcrm_transferrullTableAdapter.Adapter.DeleteCommand.Connection.Close();

            }

        }

        #endregion

        #region 系统映射表数据_操作事件
        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                (this.mytestDataSet.intcrm_wxuser.Rows.Count - 1) >= e.RowIndex &&
                m_iCrmWxEditFlag == 2)
            {
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.Open();

                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@CrmUri"].Value = dataGridView1.Rows[e.RowIndex].Cells["cRMUriDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@wxtoken"].Value = dataGridView1.Rows[e.RowIndex].Cells["wxtokenDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@crmuser"].Value = dataGridView1.Rows[e.RowIndex].Cells["crmuserDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@crmpassword"].Value = dataGridView1.Rows[e.RowIndex].Cells["crmpasswordDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@state"].Value = dataGridView1.Rows[e.RowIndex].Cells["stateDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Parameters["@Original_wxtoken"].Value = m_strCurrentWxtoken_CrmWx; //this.mytestDataSet.intcrm_wxuser.Rows[e.RowIndex]["wxtoken"];

                int i = intcrm_wxuserTableAdapter.Adapter.UpdateCommand.ExecuteNonQuery();
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.Close();
                dataGridView1.Update();

                this.intcrm_wxuserTableAdapter.Fill(this.mytestDataSet.intcrm_wxuser);
            }

            if (intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                (this.mytestDataSet.intcrm_wxuser.Rows.Count) >= e.RowIndex &&
                m_iCrmWxEditFlag == 0)
            {
                intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.Open();

                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);

                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Parameters["@CrmUri"].Value = dataGridView1.Rows[e.RowIndex].Cells["cRMUriDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_transferrull.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Parameters["@wxtoken"].Value = dataGridView1.Rows[e.RowIndex].Cells["wxtokenDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Parameters["@crmuser"].Value = dataGridView1.Rows[e.RowIndex].Cells["crmuserDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Parameters["@crmpassword"].Value = dataGridView1.Rows[e.RowIndex].Cells["crmpasswordDataGridViewTextBoxColumn"].Value;
                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Parameters["@state"].Value = dataGridView1.Rows[e.RowIndex].Cells["stateDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]);

                int i = intcrm_wxuserTableAdapter.Adapter.InsertCommand.ExecuteNonQuery();
                intcrm_wxuserTableAdapter.Adapter.InsertCommand.Connection.Close();
                dataGridView1.Refresh();

                this.intcrm_wxuserTableAdapter.Fill(this.mytestDataSet.intcrm_wxuser);
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            m_iCrmWxEditFlag = 2;
            if (dataGridView1.Rows[e.RowIndex].Cells["wxtokenDataGridViewTextBoxColumn"].Value != null)
                m_strCurrentWxtoken_CrmWx = (string)dataGridView1.Rows[e.RowIndex].Cells["wxtokenDataGridViewTextBoxColumn"].Value;
            else
                m_strCurrentWxtoken_CrmWx = "";
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            m_iCrmWxEditFlag = 0;
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            m_iCrmWxEditFlag = 1;

            if (intcrm_wxuserTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                // (this.mytestDataSet1.intcrm_wxuser.Rows.Count) >= e.RowIndex &&
               m_iCrmWxEditFlag == 1 && m_strCurrentWxtoken_CrmWx != "")
            {
                intcrm_wxuserTableAdapter.Adapter.DeleteCommand.Connection.Open();

                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);

                intcrm_wxuserTableAdapter.Adapter.DeleteCommand.Parameters["@Original_wxtoken"].Value = e.Row.Cells["wxtokenDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]);

                int i = intcrm_wxuserTableAdapter.Adapter.DeleteCommand.ExecuteNonQuery();
                intcrm_wxuserTableAdapter.Adapter.DeleteCommand.Connection.Close();

            }
        }

        #endregion

        #region 同步日志表_操作事件
        private void dataGridView3_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open &&
                (this.mytestDataSet1.intcrm_transferlog.Rows.Count - 1) >= e.RowIndex )
            {
                intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Connection.Open();

                dataGridView3.CommitEdit(DataGridViewDataErrorContexts.Commit);
                intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Parameters["@isclosed"].Value = dataGridView3.Rows[e.RowIndex].Cells["isclosedDataGridViewTextBoxColumn"].Value; //this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]; //new MySqlParameter("UpdateMethod", this.mytestDataSet1.intcrm_wxuser.Rows[e.RowIndex]["UpdateMethod"]);
                intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Parameters["@Original_ID"].Value = dataGridView3.Rows[e.RowIndex].Cells["idDataGridViewTextBoxColumn"].Value; //this.mytestDataSet.intcrm_wxuser.Rows[e.RowIndex]["wxtoken"];

                int i = intcrm_transferlogTableAdapter.Adapter.UpdateCommand.ExecuteNonQuery();
                intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Connection.Close();
                dataGridView3.Update();
            }
        }

        private void dataGridView3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (intcrm_transferlogTableAdapter.Adapter.UpdateCommand.Connection.State != ConnectionState.Open)
                this.intcrm_transferlogTableAdapter.Fill(this.mytestDataSet1.intcrm_transferlog);

        }
        #endregion

        private void dataGridView3_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            return;
        }
    }
}
