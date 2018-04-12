using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class CreateSheetDate
    {
        public int ColumnCount { set; get; }
        public string SheetName { set; get; }
        public object[,] Data { set; get; }
    }
}
