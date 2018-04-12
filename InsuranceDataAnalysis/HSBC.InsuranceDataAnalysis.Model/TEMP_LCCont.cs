using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "BI")]//aaaaaaaaaaaaa
    public class TEMP_LCCont
    {
        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { get; set; }

        [Description("PolicyNo")]
        public string PolicyNo
        {
            get
            {
                return string.IsNullOrWhiteSpace(this.policyNo) ? string.Empty : this.policyNo.PadLeft(8, '0');
            }
            set
            {
                this.policyNo = value;
            }
        }


        private string policyNo;
        [Description("RenewalTimes")]
        public string RenewalTimes { get; set; }
        [Description("ManageCom")]
        public string ManageCom { get; set; }
        [Description("SignDate")]
        public string SignDate { get; set; }
        /// <summary>
        /// 保费
        /// </summary>
        [Description("Premium")]
        public string Premium { get; set; }
    }
}
