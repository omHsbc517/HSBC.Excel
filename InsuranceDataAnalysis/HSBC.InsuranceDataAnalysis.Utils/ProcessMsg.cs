using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HSBC.InsuranceDataAnalysis.Utils
{
    public class ProcessMsg
    {
        public string Color { get; set; }
        public string Msg { get; set; }
        public int FontSize { get; set; }
        
        public ProcessMsg()
        {
            this.Color = "Black";
            this.FontSize = 12;
        }

        public ProcessMsg(string msg)
        {
            this.Color = "Black";
            this.Msg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ")+ msg;
            this.FontSize = 12;

        }

        public ProcessMsg(string msg, string clr)
        {
            this.Color = clr;
            this.Msg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg;
            this.FontSize = 12;
        }

        public ProcessMsg(string msg, string clr, int fontSize)
        {
            this.Color = clr;
            this.Msg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg;
            this.FontSize = fontSize;
        }
    }
}
