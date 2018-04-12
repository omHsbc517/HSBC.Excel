using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class ClaimSheetModel
    {
        public string Product { get; set; }
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
        public string GroupName { get; set; }
        public string MembersCertificateNo { get; set; }
        public string Membereffectivedate { get; set; }
        public string Memberexpire { get; set; }
        public string CauseOfClaim { get; set; }
        public string AdmissionServiceDate { get; set; }
        public string Discharge { get; set; }
        public string PaymentDate { get; set; }
        public string PaidAmount { get; set; }
        public string PaidAmountCurrency { get; set; }
        public string RecoveryAmount { get; set; }
        public string CompanyName { get; set; }
    }
}
