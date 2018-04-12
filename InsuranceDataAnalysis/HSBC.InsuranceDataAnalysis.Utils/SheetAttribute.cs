using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
namespace HSBC.InsuranceDataAnalysis.Utils
{
    [AttributeUsage(AttributeTargets.Class)]
    public class SheetAttribute : Attribute
    {

        public SheetAttribute(string sheetName, string endColoumn)
        {
            this.SheetName = sheetName;
            this.EndColoumn = endColoumn;
        }

        public string EndColoumn { private set; get; }

        public string SheetName { private set; get; }
      
    }
}
