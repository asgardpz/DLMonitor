using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace DLMonitor
{
    class SetupIniIP
    {

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section,
        string key, string val, string filePath);
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(string section,
        string key, string def, StringBuilder retVal,
        int size, string filePath);
        public void IniWriteValue(string Section, string Key, string Value, string inipath)
        {
            WritePrivateProfileString(Section, Key, Value, AppDomain.CurrentDomain.BaseDirectory + "\\" + inipath);
        }
        public string IniReadValue(string Section, string Key, string inipath)
        {
            StringBuilder temp = new StringBuilder(255);
            GetPrivateProfileString(Section, Key, "", temp, 255, AppDomain.CurrentDomain.BaseDirectory + "\\" + inipath);
            return temp.ToString();
        }
    }
}
