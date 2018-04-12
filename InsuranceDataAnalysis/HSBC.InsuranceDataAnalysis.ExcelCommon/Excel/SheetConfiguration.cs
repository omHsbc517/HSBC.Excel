using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class SheetConfiguration
    {
        public string SheetName { set; get; }
        public string Range { set; get; }
        public bool UseSpecialRange { set; get; }

    }
}
