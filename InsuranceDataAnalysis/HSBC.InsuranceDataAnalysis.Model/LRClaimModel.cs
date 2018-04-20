using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class LRClaimModel
    {
        public string TransactionNo { get; set; }
        public string CompanyCode { get; set; }
        public string GrpPolicyNo { get; set; }
        public string GrpProductNo { get; set; }
        public string PolicyNo { get; set; }
        public string ProductNo { get; set; }
        public string GPFlag { get; set; }
        public string MainProductNo { get; set; }
        public string MainProductFlag { get; set; }
        public string ProductCode { get; set; }
        public string LiabilityCode { get; set; }
        public string LiabilityName { get; set; }
        public string GetLiabilityCode { get; set; }
        public string GetLiabilityName { get; set; }
        public string BenefitType { get; set; }
        public string TermType { get; set; }
        public string ManageCom { get; set; }
        public string SignDate { get; set; }
        public string EffDate { get; set; }
        public string PolYear { get; set; }
        public string InvalidDate { get; set; }
        public string UWConclusion { get; set; }
        public string PolStatus { get; set; }
        public string Status { get; set; }
        public string BasicSumInsured { get; set; }
        public string RiskAmnt { get; set; }
        public string Premium { get; set; }
        public string DeductibleType { get; set; }
        public string Deductible { get; set; }
        public string ClaimRatio { get; set; }
        public string AccountValue { get; set; }
        public string FacultativeFlag { get; set; }
        public string AnonymousFlag { get; set; }
        public string WaiverFlag { get; set; }
        public string WaiverPrem { get; set; }
        public string FinalCashValue { get; set; }
        public string InsuredNo { get; set; }
        public string InsuredName { get; set; }
        public string InsuredSex { get; set; }
        public string InsuredCertType { get; set; }
        public string InsuredCertNo { get; set; }
        public string OccupationType { get; set; }
        public string AppntAge { get; set; }
        public string PreAge { get; set; }
        public string FinalLiabilityReserve { get; set; }
        public string ProfessionalFee { get; set; }
        public string SubStandardFee { get; set; }
        public string EMRate { get; set; }
        public string ProjectFlag { get; set; }
        public string InsurePeoples { get; set; }
        public string SaparateFlag { get; set; }
        public string ReInsuranceContNo { get; set; }
        public string ReinsurerCode { get; set; }
        public string ReinsurerName { get; set; }
        public string ReinsurMode { get; set; }
        public string ReinsuranceAmnt { get; set; }
        public string RetentionAmount { get; set; }
        public string QuotaSharePercentage { get; set; }
        public string ClaimNo { get; set; }
        public string AccidentDate { get; set; }
        public string ClmSettDate { get; set; }
        public string PayStatusCode { get; set; }
        public string ClaimMoney { get; set; }
        public string BackClaimMoney { get; set; }
        public string BackDate { get; set; }
        public string Currency { get; set; }
        public string ReComputationsDate { get; set; }
        public string AccountGetDate { get; set; }
    }
}
