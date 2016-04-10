using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data;

using DataAccess.dbConnect;

namespace AccessCRMForWechat
{
    public class AccessMiddleDB
    {
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
        public string FindIDForDelete(DataRow transferlog)
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
                    if (tmpReader.Rows.Count > 0)
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
                    if (tmpReader.Rows.Count > 0)
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

        public string FindIDForUpdate(DataRow transferlog)
        {
            string strSql, strReturn = "";
            DataTable tmpReader;
            if (transferlog["operatetype"].ToString() == "2")
            {
                if (transferlog["wxrecordid"].ToString().Length > 0)
                {
                    strSql = "select * from intcrm_transferlog where direct=0 and operatetype<>2 and wxrecordid='" + transferlog["wxrecordid"].ToString() + "'";
                    tmpReader = ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count > 0)
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
                    tmpReader = ConnectionPool_mysql.GetQuery(strSql);
                    if (tmpReader.Rows.Count > 0)
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

        public void CreateLog(string entityname, string operatetype, string crmrecordid, string wxrecordid, string direct, string token)
        {
            string strInsertSql,strValueSql;

            strInsertSql = "insert into intcrm_transferlog(createtime";
            strValueSql = " values('" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "'";

            if (entityname.Length > 0)
            {
                strInsertSql += ",entityname";
                strValueSql += ",'" + entityname + "'";
            }
            if (operatetype.Length > 0)
            {
                strInsertSql += ",operatetype";
                strValueSql += ",'" + operatetype + "'";
            }
            if (crmrecordid.Length > 0)
            {
                strInsertSql += ",crmrecordid";
                strValueSql += ",'" + crmrecordid + "'";
            }
            if (wxrecordid.Length > 0)
            {
                strInsertSql += ",wxrecordid";
                strValueSql += ",'" + wxrecordid + "'";
            }
            if (direct.Length > 0)
            {
                if (direct == "True" || direct == "False" || direct == "true" || direct == "false")
                {
                    strInsertSql += ",direct";
                    strValueSql += "," + direct + "";
                }
            }
            if (token.Length > 0)
            {
                strInsertSql += ",token";
                strValueSql += ",'" + token + "'";
            }

            strInsertSql += ")";
            strValueSql += ")";

            ConnectionPool_mysql.UpdateQuery(strInsertSql + strValueSql);
        }

        #endregion
    }
}
