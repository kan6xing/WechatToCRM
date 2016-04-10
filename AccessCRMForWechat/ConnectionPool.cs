using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Xml;
using DataAccess;
using MySql.Data;
using MySql.Data.MySqlClient;//using System.Web.Configuration;

namespace DataAccess.dbConnect
{
    public class ConnectionPool_mysql
    {
        private static MySql.Data.MySqlClient.MySqlConnection con= new MySql.Data.MySqlClient.MySqlConnection();
        private static MySql.Data.MySqlClient.MySqlTransaction trans;

        private static MySql.Data.MySqlClient.MySqlCommand cmd;
        private static MySql.Data.MySqlClient.MySqlDataAdapter myAdapter;

        //private static SqlTransaction trans;
        //private static SqlConnection con = new SqlConnection();
        private static bool m_isTransaction = false;
        private static string m_strConnection = "";//ConfigurationManager.AppSettings["con"];
        private static string m_strIsBS = "true";//�Ƿ�ΪBSϵͳӦ��


        public ConnectionPool_mysql()
        {
            m_strIsBS = IniConfig.GetValue("DBConnection", "IsBS");
            if (m_strIsBS == "") m_strIsBS = "true";
            if (m_strIsBS == "true")
                m_strConnection = ConfigurationManager.AppSettings["con_mysql"].ToString();
            else
                m_strConnection = IniConfig.GetValue("DBConnection", "CnnString");

            if (m_strConnection.Contains("Data Source"))
            {
                //��δ���������ִ����б�����ٱ���
                //if (m_strIsBS == "true")
                //    //ConfigurationManager.AppSettings["con"] = DI_JH.Functions.StringToBase64(m_strConnection);
                //else
                //    DI_JH.IniConfig.WriteValue("DBConnection", "CnnString", ServerConnection.StringToBase64(m_strConnection));
            }
            else//�����������ӱ����ִ�����
                //m_strConnection = DI_JH.Functions.StringFromBase64(m_strConnection);

            ConnectionPool_mysql.con = new MySql.Data.MySqlClient.MySqlConnection(m_strConnection);
        }

        public ConnectionPool_mysql(string Topic)
        {
            //m_strConnection = ConfigurationManager.ConnectionStrings[Topic].ToString();
            ConnectionPool_mysql.con = new MySql.Data.MySqlClient.MySqlConnection(m_strConnection);
        }

        #region "����"
        public MySql.Data.MySqlClient.MySqlTransaction Transaction
        {
            get
            { return trans; }
        }

        public bool IsTransaction
        {
            get { return m_isTransaction; }
        }

        public static MySql.Data.MySqlClient.MySqlConnection Connection
        {
            get
            {
                return ConnectionPool_mysql.con;
            }
        }

        public static string ConnectionString
        {
            get
            { return m_strConnection; }
            set
            { m_strConnection = value; }
        }
        #endregion

        #region "���Ӵ���"

        /// <summary>
        /// ��õ�ǰ����
        /// </summary>

        /// <summary>
        ///  ������
        /// </summary>
        public static void OpenConnecion()
        {
            if (con.State != System.Data.ConnectionState.Open || con == null)
            {
                con.ConnectionString = m_strConnection;
                con.Open();
            }
        }

        /// <summary>
        ///  �ر�����
        /// </summary>
        public static void CloseConnection()
        {
            if (ConnectionPool_mysql.con.State ==System.Data.ConnectionState.Open)
            {

                if (m_isTransaction == true)
                {
                    trans.Dispose();
                    m_isTransaction = false;
                }
                con.Close();
                con.Dispose();
            }
        }
        #endregion

        #region"������"
        /// <summary>
        /// ��ʼ������
        /// </summary>
        public static void BeginTransaction()
        {
            OpenConnecion();
            trans = con.BeginTransaction();
            m_isTransaction = true;

        }

        /// <summary>
        /// �ύ����
        /// </summary>
        public static void CommitTransaction()
        {
            if (m_isTransaction != false)
            {
                trans.Commit();
                m_isTransaction = false;
                con.Close();
                con.Dispose();
            }
        }
        /// <summary>
        /// �ع�����
        /// </summary>
        public static void RollbackTransaction()
        {
            if (m_isTransaction != false)
            {
                trans.Rollback();
                trans.Dispose();
                con.Close();
                con.Dispose();
                m_isTransaction = false;
            }
        }
        #endregion

        #region"��䴦��"
        /// <summary>
        ///  ���˵�����'����ֹsqlע�빥��
        /// </summary>
        /// <param name="strInput">�����ַ���</param>
        /// <returns></returns>
        private static string DealSqlStr(string strInput)
        {
            string strTemp;
            strTemp = strInput;
            if (strTemp != "" && strTemp.Length > 0)
                strTemp = strTemp.Replace("'", "''");
            else
                strTemp = "";
            return strTemp;
        }


        /// <summary>
        /// ִ��Sql���
        /// </summary>
        /// <param name="strSql">sql���(insert update delete)</param>
        public static int UpdateQuery(string strSql)
        {
            int iReturn;
            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            // ��
            OpenConnecion();
            cmd.Connection = ConnectionPool_mysql.con;
            if (m_isTransaction == true)
            {
                cmd.Transaction = ConnectionPool_mysql.trans;
            }
            cmd.CommandText = strSql;
            iReturn = cmd.ExecuteNonQuery();

            if (!m_isTransaction)
            {
                CloseConnection();
            }
            return iReturn;
        }

        /// <summary>
        /// ִ��Sql���
        /// </summary>
        /// <param name="strSql">sql���(insert update delete)</param>
        public static int UpdateQuery(string strSql, MySql.Data.MySqlClient.MySqlConnection con, MySql.Data.MySqlClient.MySqlTransaction trans)
        {
            int iReturn;
            // ��
            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            cmd.Connection = con;

            cmd.Transaction = trans;

            cmd.CommandText = strSql;
            iReturn = cmd.ExecuteNonQuery();
            if (!m_isTransaction)
            {
                CloseConnection();
            }
            return iReturn;
        }


        /// <summary>
        /// ִ��sql����identityֵ
        /// </summary>
        /// <param name="sql">sql���(insert)</param>
        /// <param name="a">����ֵ</param>
        /// <returns></returns>
        public static int ExeInsSql(string sql)
        {
            int identity = -1;

            // ��
            OpenConnecion();

            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            cmd.Connection = ConnectionPool_mysql.con;
            if (m_isTransaction == true)
            {
                cmd.Transaction = ConnectionPool_mysql.trans;
            }
            cmd.CommandText = sql + " select @@identity as 'identity'";

            // ��һ�е�һ�е�ֵΪ��ǰID
            MySqlDataReader dr = cmd.ExecuteReader();

            if (dr.Read())
            {
                identity = int.Parse(dr[0].ToString());
            }

            dr.Close();

            // �ͷ�
            CloseConnection();
            return identity;
        }

        public static long GetJustID()
        {
            long identity = -1;

            // ��
            OpenConnecion();

            MySql.Data.MySqlClient.MySqlCommand cmdID;
            //string  sID="";
            cmdID = new MySql.Data.MySqlClient.MySqlCommand("select @@identity", ConnectionPool_mysql.con);
            if (m_isTransaction) cmdID.Transaction = ConnectionPool_mysql.trans;

            identity = Convert.ToInt64(cmdID.ExecuteScalar());
            return identity;
        }


        /// <summary>
        ///  ��ȡDataSet
        /// </summary>
        /// <param name="sql">sql���(select)</param>
        /// <returns></returns>
        private DataSet GetDataSet(string sql)
        {

            // ��
            OpenConnecion();


            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            cmd.Connection = ConnectionPool_mysql.con;
            if (m_isTransaction == true)
            {
                cmd.Transaction = ConnectionPool_mysql.trans;
            }
            DataSet ds = new DataSet();
            MySql.Data.MySqlClient.MySqlDataAdapter da = new MySql.Data.MySqlClient.MySqlDataAdapter();
            cmd.CommandText = sql;
            da.SelectCommand = cmd;
            da.Fill(ds);
            CloseConnection();
            return ds;
        }

        /// <summary>
        /// ��ȡDataTable
        /// </summary>
        /// <param name="sql">sql���(select)</param>
        /// <returns>DataTable</returns>
        public static DataTable GetQuery(string sql)
        {

            // ��
            OpenConnecion();

            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            cmd.Connection = ConnectionPool_mysql.con;
            cmd.CommandTimeout = cmd.Connection.ConnectionTimeout;
            if (m_isTransaction == true)
            {
                cmd.Transaction = ConnectionPool_mysql.trans;
            }

            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            MySql.Data.MySqlClient.MySqlDataAdapter da = new MySql.Data.MySqlClient.MySqlDataAdapter();

            cmd.CommandText = sql;
            da.SelectCommand = cmd;

            da.Fill(ds, "tmp");
            dt = ds.Tables["tmp"];
            if (!m_isTransaction)
            {
                CloseConnection();
            }
            return dt;
        }

        /// <summary>
        /// ���������select����ѯ���
        /// </summary>
        /// <param name="strSelect">����ֶ��б�������select�ؼ��֣�</param>
        /// <param name="strFrom">����</param>
        /// <param name="strWhere">��������</param>
        /// <param name="strOrderBy">�����ֶα�</param>
        /// <param name="strGroupBy">�����ֶα�</param>
        /// <param name="strHaving">���������ֶα�</param>
        /// <returns>���ؽ�����ݱ�</returns>
        public static DataTable GetQuery(string strSelect, string strFrom, string strWhere, string strOrderBy, string strGroupBy, string strHaving)
        {

            // ��
            string strSql;

            if (strSelect.Length > 0)
                strSql = "select " + strSelect;
            else
                return null;

            if (strFrom.Length > 0)
                strSql += " from " + strFrom;
            else
                return null;

            if (strWhere.Length > 0)
                strSql += " where " + strWhere;

            if (strOrderBy.Length > 0)
                strSql += " order by " + strOrderBy;

            if (strGroupBy.Length > 0)
                strSql += " group by " + strGroupBy;

            if (strHaving.Length > 0)
                strSql += " having " + strHaving;

            return GetQuery(strSql);
        }

        /// <summary>
        /// ���ؼ�¼����
        /// </summary>
        /// <param name="sql">sql���(select)</param>
        /// <returns>��¼����</returns>
        private static int RowConut(string sql)
        {
            OpenConnecion();
            DataSet ds = new DataSet();

            MySql.Data.MySqlClient.MySqlDataAdapter command = new MySql.Data.MySqlClient.MySqlDataAdapter(sql, con);
            command.Fill(ds, "ds");
            if (!m_isTransaction)
                CloseConnection();
            con.Close();
            return ds.Tables.Count;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        private static string QueryReturnOneString(string sql)
        {
            string str = string.Empty;
            OpenConnecion();
            MySql.Data.MySqlClient.MySqlCommand cmd = new MySql.Data.MySqlClient.MySqlCommand();
            cmd.CommandText = sql;
            cmd.Connection = con;
            MySqlDataReader dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                str = dr[0].ToString();
            }
            return str;
        }
        #endregion

    }

}
