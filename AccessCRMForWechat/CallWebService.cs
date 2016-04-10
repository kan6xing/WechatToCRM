using System;
using System.Collections.Generic;
using System.Text;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.IO;
using System.Net;
using System.Web.Services.Description;
using System.Reflection;
using System.Data;


namespace DI_Wechat.WebService
{
    public class _CallWebService
    {
        protected static string m_strURL, m_strOrgCode;
        protected static CompilerResults result;

        public static string URL
        {
            get { return m_strURL; }
            set
            {
                m_strURL = value;
                result = InvokeWebMethod();
            }
        }
        public _CallWebService() { }
        protected static CompilerResults InvokeWebMethod()
        {
            WebClient web = new WebClient();
            Stream stream = web.OpenRead(m_strURL);

            //创建和格式化WSDL文档。
            ServiceDescription description = ServiceDescription.Read(stream);

            //创建客户端代理类
            ServiceDescriptionImporter importer = new ServiceDescriptionImporter();

            importer.ProtocolName = "Soap";//指定访问协议
            importer.Style = ServiceDescriptionImportStyle.Client;//生成客户端代理
            importer.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties | System.Xml.Serialization.CodeGenerationOptions.GenerateNewAsync;

            importer.AddServiceDescription(description, null, null);//添加WSDL文档

            //使用CodeDom编译客户端代理类
            CodeNamespace nmspace = new CodeNamespace();//为代理类添加命名空间，缺省为全局空间
            CodeCompileUnit unit = new CodeCompileUnit();
            unit.Namespaces.Add(nmspace);

            ServiceDescriptionImportWarnings warning = importer.Import(nmspace, unit);
            CodeDomProvider provider = CodeDomProvider.CreateProvider("CSharp");

            CompilerParameters parameter = new CompilerParameters();
            parameter.GenerateExecutable = false;
            parameter.GenerateInMemory = true;

            parameter.ReferencedAssemblies.Add("System.dll");
            parameter.ReferencedAssemblies.Add("System.XML.dll");
            parameter.ReferencedAssemblies.Add("System.Web.Services.dll");
            parameter.ReferencedAssemblies.Add("System.Data.dll");

            CompilerResults result = provider.CompileAssemblyFromDom(parameter, unit);
            return result;

        }
    }

    public class WechatService : _CallWebService
    {
        public WechatService() 
        {
        }

        public static string InvokeWebMethod(string MethodName, object[] Parameter)
        {

            //CompilerResults result = InvokeWebMethod();

            //使用Refeciton调用Webservice.
            if (!result.Errors.HasErrors)
            {
                Assembly sam = result.CompiledAssembly;
                Type t = sam.GetType("Customer");//如果在前面为代理类添加了命名空间，此处需要将命名空间添加到类型前面。

                object o = Activator.CreateInstance(t);
                MethodInfo method = t.GetMethod(MethodName);
                MethodInfo[] methods = t.GetMethods();
                try
                {
                    object Result = method.Invoke(o, Parameter);

                    if (Result == null)
                        return "true";
                    else 
                    return Result.ToString();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.ToString());
                }
            }
            else
            {
                throw new Exception(result.Errors.ToString());
            }
        }

        public static string InvokeWeb_UserLogin(string UserID, string Password)
        {

            //CompilerResults result = InvokeWebMethod();

            //使用Refeciton调用WebCustomer.
            if (!result.Errors.HasErrors)
            {
                Assembly sam = result.CompiledAssembly;
                Type t = sam.GetType("Customer");//如果在前面为代理类添加了命名空间，此处需要将命名空间添加到类型前面。

                object o = Activator.CreateInstance(t);
                MethodInfo method = t.GetMethod("UserLogin");
                MethodInfo[] methods = t.GetMethods();
                try
                {
                    object Result = method.Invoke(o, new object[] {UserID,Password }).ToString();
                    return Result.ToString();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.ToString());
                }
            }
            else
            {
                throw new Exception(result.Errors.ToString());
            }
        }


    }

}
