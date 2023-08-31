using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_UTIL
{
    public class Global
    {
        public static SAPbobsCOM.Company sboCompany = new SAPbobsCOM.Company();
        public static SAPbobsCOM.UserTable userTable;
        public static SAPbouiCOM.Application sboApplication;
        public static SAPbobsCOM.SBObob sboBob;

        public static SAPbobsCOM.Recordset oSq;


        public static int QueryPosition;
        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\Service_Creation_Log_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " - " + Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " - " + Message);
                }
            }
        }
    }
}
