using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "AB")]
    public class TEMP_LLClaimDetail
    {
        [Description("PolicyNo")]
        public string PolicyNo { set; get; }//G

        private string policyNo;



        [Description("GetLiabilityCode")]
        public string GetLiabilityCode { set; get; }//M


        [Description("GetLiabilityName")]
        public string GetLiabilityName { set; get; }//O

        [Description("BenefitType")]
        public string BenefitType { set; get; }//J

        [Description("DeductibleType")]
        public string DeductibleType { set; get; }//W

        [Description("Deductible")]
        public string Deductible { set; get; }//X

        [Description("ClaimRatio")]
        public string ClaimRatio { set; get; }//Y

        [Description("ClmCaseNo")]
        public string ClmCaseNo { set; get; }//C

        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { get; set; }//E

        [Description("ClmSettDate")]
        public string ClmSettDate { get; set; }//AA

        [Description("PayStatusCode")]
        public string PayStatusCode { get; set; }//AB

        [Description("ProductNo")]
        public string ProductNo { get; set; }//H
    }
}
