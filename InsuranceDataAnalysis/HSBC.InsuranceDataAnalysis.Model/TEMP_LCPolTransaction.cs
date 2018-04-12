using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "R")]
    public class TEMP_LCPolTransaction
    {
        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { set; get; }//C

        [Description("PolicyNo")]
        public string PolicyNo//D
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

        [Description("EndorAcceptNo")]
        public string EndorAcceptNo { set; get; }//Q

        [Description("EndorsementNo")]
        public string EndorsementNo { set; get; }//R
    }
}
