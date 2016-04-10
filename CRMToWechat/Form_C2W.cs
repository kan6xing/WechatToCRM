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
    public partial class Form_C2W : Form
    {
        AccessCRMForWechat.AccessCRM m_accessCrm;

        public Form_C2W()
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
                    m_accessCrm.WriteSendLog(dr["wxtoken"].ToString());        //登记微信发送任务

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
                            case "lettergroup":
                                LetterGroupToWechat(drtransferfull["token"].ToString());   //向分组发送微信
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
                DataTable dttransferlog = TransferLog("letter", token,"1");
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
                                if (letter != null)
                                {
                                    //调用微信平台 Webservice 发送消息    //strSendResult = 
                                    DI_Wechat.WebService.WechatService.URL = DataAccess.IniConfig.GetValue("WechatService", "Url");
                                    string strReturn = DI_Wechat.WebService.WechatService.InvokeWebMethod("serviceMessageSendInterface",
                                        new object[] { letter["new_imgsourceid"].ToString(), letter["new_token"].ToString(), letter["new_openid"].ToString() });

                                    if (strReturn == "true" || strReturn == "True" || strReturn == "1")
                                    {
                                        WriteBackLog(drtransferlog["id"].ToString(), letter.Id.ToString(), "", "1");
                                        try
                                        {
                                            m_accessCrm.UpdateWASendState(letter, 100000002);
                                        }
                                        catch (Exception ex)
                                        {
                                            DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "发送微信消息回写letter错误错误：日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                                            //WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                                        }
                                    }
                                    else
                                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString(), "发送微信消息错误：源数据id=" + drtransferlog["crmrecordid"].ToString() + "发送失败！"+
                                            letter["new_imgsourceid"].ToString() + "--" + letter["new_token"].ToString() + "--" + letter["new_openid"].ToString()+
                                            "源错误描述:" + strReturn);
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

        private void LetterGroupToWechat(string token)
        {
            try
            {
                //string strSendResult;

                Entity campaignactivity = null;
                DataTable dttransferlog = TransferLog("lettergroup", token, "1");
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
                                campaignactivity = m_accessCrm.GetCampaignActivity(drtransferlog["crmrecordid"].ToString());
                                if (campaignactivity != null)
                                {
                                    //调用微信平台 Webservice 发送消息    //strSendResult = 
                                    DI_Wechat.WebService.WechatService.URL = DataAccess.IniConfig.GetValue("WechatService", "Url");
                                    string strReturn = DI_Wechat.WebService.WechatService.InvokeWebMethod("GroupMessageSendInterface",
                                        new object[] { campaignactivity["new_imgsourceid"].ToString(), campaignactivity["new_token"].ToString(), campaignactivity["new_cusgroupsourceid"].ToString() });

                                    if (strReturn == "true" || strReturn == "True" || strReturn == "1")
                                    {
                                        WriteBackLog(drtransferlog["id"].ToString(), campaignactivity.Id.ToString(), "", "1");
                                        try
                                        {
                                            m_accessCrm.UpdateWASendState(campaignactivity, 100000002);
                                        }
                                        catch (Exception ex)
                                        {
                                            DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString()+"按组发送微信消息回写campaignactivity错误错误：", "日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                                            //WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                                        }
                                    }
                                    else
                                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString()+"按组发送微信消息错误：", "源数据id=" + drtransferlog["crmrecordid"].ToString() + "发送失败！"+
                                            campaignactivity["new_imgsourceid"].ToString() + "--" + campaignactivity["new_token"].ToString() + "--" + campaignactivity["new_cusgroupsourceid"].ToString()+
                                            "源错误描述:"+strReturn);
                                }
                                else
                                {
                                    DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString()+"按组发送微信消息错误：", "源数据id=" + drtransferlog["crmrecordid"].ToString() + "不存在！");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString()+"按组发送微信消息错误：", "日志ID=" + drtransferlog["id"].ToString() + "。" + ex.Message);
                        WriteBackLog(drtransferlog["id"].ToString(), "", "", "0");
                    }
                    finally
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                DataAccess.ErrorLog.WriteValue(DateTime.Today.ToString(), DateTime.Now.ToLongTimeString() +"按组发送微信消息错误：",  ex.Message);
            }
            finally
            {
            }
        }


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
