using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCore
{
    public class Column
    {
        public string ColumnName { get; set; }

        public List<Cell> Cells { get; set; }

        public Column()
        {
            Cells = new List<Cell>();
        }
    }
}
