using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Configuration;

namespace MiddleDBService
{
    // 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码、svc 和配置文件中的类名“Service1”。
    // 注意: 为了启动 WCF 测试客户端以测试此服务，请在解决方案资源管理器中选择 Service1.svc 或 Service1.svc.cs，然后开始调试。
    public class Service1 : IService1
    {
        AccessCRMForWechat.AccessMiddleDB m_accessMiddleDB=new AccessCRMForWechat.AccessMiddleDB();

        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }
            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }
            return composite;
        }


        public string CreateLog(string entityname, string operatetype, string crmrecordid, string wxrecordid, string direct, string token)
        {
            try
            {
                DataAccess.dbConnect.ConnectionPool_mysql.ConnectionString = ConfigurationManager. AppSettings["con_mysql"].ToString();
                DataAccess.dbConnect.ConnectionPool_mysql.OpenConnecion();

                m_accessMiddleDB.CreateLog(entityname, operatetype, crmrecordid, wxrecordid, direct, token);

                return "";
            }
            catch (Exception ex)
            {
                return "友联中间服务创建日志错误：" + ex.Message;
            }
        }
    }
}
