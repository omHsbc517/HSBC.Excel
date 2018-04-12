using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public static class ExcelProcess
    {
        private static HashSet<uint> _excelProcessIds = new HashSet<uint>();
        public static HashSet<uint> ExcelProcessIds
        {
            get
            {
                return _excelProcessIds;
            }
        }

        public static void AddProcessId(uint processId)
        {
            _excelProcessIds.Add(processId);
        }

        public static void ClearProcessIds()
        {
            _excelProcessIds.Clear();
        }

        public static void KillAllExcel()
        {
            foreach (var item in _excelProcessIds)
            {
                if (item > 0)
                {
                    try
                    {
                        var process = Process.GetProcessById((int)item);
                        if (process != null)
                        {
                            process.Kill();
                        }
                    }
                    catch (Exception)
                    {


                    }
                }
            }
        }
    }
}
