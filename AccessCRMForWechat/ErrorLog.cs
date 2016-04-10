using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace DataAccess
{
    public class ErrorLog
    {
        const string INIFileName = "ErrorLog.ini";
        [DllImport("kernel32")]
        static extern int WritePrivateProfileString(string Section, string Key, string Value, string iniFile);
        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(string Section, string Key, string defaultValue, StringBuilder returnValue, int Size, string iniFile);

        public ErrorLog() { }
        public static string GetValue(string Section, string Key)
        {
            StringBuilder SBReturn = new StringBuilder(2048);
            GetPrivateProfileString(Section, Key, "", SBReturn, 1024, INIFileName);
            return SBReturn.ToString();
        }

        public static int WriteValue(string Section, string Key, string Value)
        {
            return WritePrivateProfileString(Section, Key, Value, INIFileName);
        }
    }
    public class IniConfig
    {
        const string INIFileName = "Wechat_Config.ini";
        [DllImport("kernel32")]
        static extern int WritePrivateProfileString(string Section, string Key, string Value, string iniFile);
        [DllImport("kernel32")]
        static extern int GetPrivateProfileString(string Section, string Key, string defaultValue, StringBuilder returnValue, int Size, string iniFile);

        public IniConfig() { }
        public static string GetValue(string Section, string Key)
        {
            StringBuilder SBReturn = new StringBuilder(2048);
            GetPrivateProfileString(Section, Key, "", SBReturn, 1024, INIFileName);
            return SBReturn.ToString();
        }

        public static int WriteValue(string Section, string Key, string Value)
        {
            return WritePrivateProfileString(Section, Key, Value, INIFileName);
        }
    }
}

