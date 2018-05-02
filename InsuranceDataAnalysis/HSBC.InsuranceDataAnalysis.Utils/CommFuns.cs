using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace HSBC.InsuranceDataAnalysis.Utils
{
    public static class CommFuns
    {
        public static string OriganizationCode
        {
            get
            {
                return "000131";
            }
        }

        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        public static bool Equal<T>(this T obj1, T obj2)
        {
            if (obj1 == null && obj2 == null)
            {
                return true;
            }
            else if (obj1 == null || obj2 == null)
            {
                return false;
            }
            else if (obj1.Equals(obj2))
            {
                return true;
            }

            return false;
        }

        public static bool IsChinese(this string str)
        {
            return Regex.IsMatch(str, @"[\u4e00-\u9fa5]+"); // chinese
        }

        public static bool IsEnglish(this string str)
        {
            return Regex.IsMatch(str, @"^[A-Za-z0-9\s?]+$");
        }

        public static bool IsHasEnglishChart(this string str)
        {
            return Regex.IsMatch(str, @"\.[A-Za-z\s?]+\.");
        }

        public static string ConvertNull2String<T>(this T str)
        {
            if (str == null)
                return "";
            else
                return str.ToString();
        }

        /// <summary>
        /// use for convert excel Scientific Counting string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ConvertScientific2String(this string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return "";
            else
                return "'" + str;
        }

        public static void KillThisProcess(this string processName)
        {
            if (string.IsNullOrWhiteSpace(processName)) return;

            Process[] ps = Process.GetProcesses();
            foreach (var item in ps)
            {
                if (item.ProcessName.ToUpper().Contains(processName))
                {
                    item.Kill();
                }
            }
        }

        private const int OF_READWRITE = 2;
        private const int OF_SHARE_DENY_NONE = 0x40;
        private static readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        public static bool IsFileOpen(this string path)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {
                fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);

                inUse = false;
            }
            catch
            {
            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true in used,false no used  
        }

        public static bool IsHasSpecialChart(this string str)
        {
            // TODO： checck ???
            return false;
        }

        public static string GetTransactionNo(int serialNumber, string yearMonthDay)
        {
            string origanizationCode = string.Empty;
            origanizationCode = OriganizationCode + yearMonthDay + serialNumber.ToString().PadLeft(10, '0');
            return origanizationCode;
        }

        public static string GetTransactionNo2(int serialNumber, string yearMonthDay)
        {
            string origanizationCode = string.Empty;
            origanizationCode = OriganizationCode + yearMonthDay + "RE0001";
            return origanizationCode;
        }

        public static string GetTransactionNo4(int serialNumber, string yearMonthDay)
        {
            string origanizationCode = string.Empty;
            origanizationCode = OriganizationCode + yearMonthDay + "RE4" + serialNumber.ToString().PadLeft(7, '0');
            return origanizationCode;
        }

        public static string GetTransactionNo5(int serialNumber, string yearMonthDay)
        {
            string origanizationCode = string.Empty;
            origanizationCode = OriganizationCode + yearMonthDay + "RE5" + serialNumber.ToString().PadLeft(7, '0');
            return origanizationCode;
        }

        public static string GetTransactionNo6(int serialNumber, string yearMonthDay)
        {
            string origanizationCode = string.Empty;
            origanizationCode = OriganizationCode + yearMonthDay + "RE6" + serialNumber.ToString().PadLeft(7, '0');
            return origanizationCode;
        }


        public static void KillExcelProcess()
        {
            try
            {
                var listProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL").ToList();
                foreach (var item in listProcesses)
                {
                    item.Kill();
                }

            }
            catch (Exception)
            {
            }
        }


        public static string GetMainProductFlag(string productCode)
        {
            string mainProductFlag = string.Empty;
            productCode = string.IsNullOrWhiteSpace(productCode) ? string.Empty : productCode.Trim().ToUpper();
            switch (productCode)
            {
                case "HC2":
                    mainProductFlag = "2";
                    break;
                case "MI1":
                    mainProductFlag = "2";
                    break;
                case "MI2":
                    mainProductFlag = "2";
                    break;
                case "MI3":
                    mainProductFlag = "2";
                    break;
                case "MM1":
                    mainProductFlag = "2";
                    break;
                default:
                    mainProductFlag = "1";
                    break;
            }
            return mainProductFlag;
        }
    }
}
