using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class LRInsureContModel
    {
        public string TransactionNo { get; set; }
        public string CompanyCode { get; set; }
        public string ReInsuranceContNo { get; set; }
        public string ReInsuranceContName { get; set; }
        public string ReInsuranceContTitle { get; set; }
        public string MainReInsuranceContNo { get; set; }
        public string ContOrAmendmentType { get; set; }
        public string ContAttribute { get; set; }
        public string ContStatus { get; set; }
        public string TreatyOrFacultativeFlag { get; set; }
        public string ContSigndate { get; set; }
        public string PeriodFrom { get; set; }
        public string PeriodTo { get; set; }
        public string ContType { get; set; }
        public string ReinsurerCode { get; set; }
        public string ReinsurerName { get; set; }
        public string ChargeType { get; set; }

    }
}
