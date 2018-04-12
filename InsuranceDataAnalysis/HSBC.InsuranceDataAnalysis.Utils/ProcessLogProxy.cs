using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HSBC.InsuranceDataAnalysis.Utils
{
    public static class ProcessLogProxy
    {
        public static Action<int> ProgressValue;
        public static Action<string> Normal;
        public static Action<string> Error;
        public static Action<ProcessMsg> MessageAction;
        public static Action<string> SuccessMessage;
       
    }
}
