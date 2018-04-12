using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class TEMP_LLClaimDetail
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
        public string GetLiabilityCode { set; get; }//M
        public string GetLiabilityName { set; get; }//O
        public string BenefitType { set; get; }//J
        public string DeductibleType { set; get; }//W
        public string Deductible { set; get; }//X
        public string ClaimRatio { set; get; }//Y


    }
}
