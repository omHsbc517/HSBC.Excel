using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class WriteSheetData
    {
        public object[,] Data { set; get; }
        public string SheetName { set; get; }
        public string StartCell { set; get; }
        public string EndCell { set; get; }
    }
}
