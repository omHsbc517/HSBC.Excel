using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class TEMP_LLClaimPolicy
    {
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
        public string ClaimNo { set; get; }//C
        public string PayStatusCode { set; get; }//Y
    }
}
