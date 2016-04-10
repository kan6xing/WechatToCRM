using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestWCFClient
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //client;// = new Service1Client();

 /*           
            System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding();
            Service1Client client = new Service1Client(binding, new System.ServiceModel.EndpointAddress("http://crmserver/mds/service1.svc"));//http://115.28.81.79:7001/mds
            string strTest = client.CreateLog("","3","248239-039423439-2390234-9343942","","false","23902dsd");*/
            

            DI_Wechat.WebService.WechatService.URL = "http://test.ylxrm.com:80/serverSoap.php?WSDL";//serviceMessageSendInterface
            string strReturn = DI_Wechat.WebService.WechatService.InvokeWebMethod("GroupMessageSendInterface",
                new object[] { "253","pinvyp1416366518","35" }); //"274", "pinvyp1416366518", "oHrm1jq8e4Br20gzgPNTsXQvD9mw"


            /*            string strSql = "select * from intcrm_transferrull ";

                        DataTable reader = DataAccess.dbConnect.ConnectionPool_mysql.GetQuery(strSql);

                        DI_Wechat.WebService.WechatService.URL = "http://crmserver/mds/service1.svc?WSDL";
                        string strReturn = DI_Wechat.WebService.WechatService.InvokeWebMethod("CreateLog",
                            new object[] { "", "3", "248239-039423439-2390234-9343942", "", "false", "23902dsd" });*/
        }
    }
}
