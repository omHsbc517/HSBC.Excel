using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class SpecialValue
    {
        public string SheetName { set; get; }
        public int RowIndex { set; get; }
        public int ColIndex { set; get; }
        public string Value { set; get; }
    }
}
