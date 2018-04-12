using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class InforceBusinessListing
    {
        public string CompanyName { get; set; }

        private string policyNo;
        public string PolicyNo
        {
            get
            {
                return string.IsNullOrWhiteSpace(this.policyNo) ? string.Empty : string.IsNullOrWhiteSpace(this.policyNo) ? string.Empty : this.policyNo.PadLeft(8, '0');
            }
            set
            {
                this.policyNo = value;
            }
        }
        public string MemberCertificateNo { get; set; }
        public string Sex { get; set; }
        public string DateofBirth { get; set; }
        public string OccupationClass { get; set; }
        public string AgeofMemberWhenJoiningtheScheme { get; set; }
        public string ProductCode { get; set; }
        public string Coverage1 { get; set; }
        public string Attainedage { get; set; }
        public string ExtraMortality { get; set; }
        public string SumInsured { get; set; }
        public string InitialSumatRisk { get; set; }
        public string SumReinsured { get; set; }
        public string Retention { get; set; }
        public string MonthlyReinsurancePremium { get; set; }
        public string MonthlyReinsuranceCommission { get; set; }
        public string Coverage2 { get; set; }
        public string ExtraMorbidity { get; set; }
        public string SumInsured2 { get; set; }
        public string InitialSumatRisk2 { get; set; }
        public string SumReinsured2 { get; set; }
        public string Retention2 { get; set; }
        public string MonthlyReinsurancePremium2 { get; set; }
        public string MonthlyReinsuranceCommission2 { get; set; }
        public string EffectiveDate { get; set; }
        public string AutomaticorFacultative { get; set; }
        public string TaxInd { get; set; }
        public string RI_RATIO_1 { get; set; }
        public string RI_RATIO_2 { get; set; }

        public bool IsMrHealth { get; set; }

    }
}
