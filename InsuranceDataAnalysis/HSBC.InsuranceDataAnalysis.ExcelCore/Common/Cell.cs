using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCore
{
    public class Cell
    {
        public int RowIndex { get; set; }
        public string ColumnName { get; set; }
        public string Value { get; set; }
    }
}
