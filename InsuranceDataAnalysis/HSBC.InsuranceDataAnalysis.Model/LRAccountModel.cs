using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
   public class LRAccountModel
    {
        public string TransactionNo { get; set; }
        public string CompanyCode { get; set; }
        public string AccountID { get; set; }
        public string AccountingPeriodfrom { get; set; }
        public string AccountingPeriodto { get; set; }
        public string ReinsurerCode { get; set; }
        public string ReinsurerName { get; set; }
        public string ReInsuranceContNo { get; set; }
        public string ReInsuranceContName { get; set; }
        public string Currency { get; set; }
        public string ReinsurancePremium { get; set; }
        public string ReinsuranceCommssionRate { get; set; }
        public string ReinsuranceCommssion { get; set; }
        public string ReturnReinsurancePremium { get; set; }
        public string ReturnReinsuranceCommssion { get; set; }
        public string ReturnSurrenderPay { get; set; }
        public string ReturnClaimPay { get; set; }
        public string ReturnMaturity { get; set; }
        public string ReturnAnnuity { get; set; }
        public string ReturnLivBene { get; set; }
        public string AccountStatus { get; set; }
        public string PairingStatus { get; set; }
        public string PairingDate { get; set; }
        public string CurrentRate { get; set; }

    }
}
