using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class HugeDisasterModel
    {
        public string ProductCode { set; get; }
        public string TypeI { set; get; }
        public string TypeII { set; get; }
        public string ReinsurerName { set; get; }
        public string BenefitReinsured { set; get; }
        public string RImethodI { set; get; }
        public string RImethodII { set; get; }
        public string Percentage { set; get; }
        public string Retention { set; get; }
        public string Remark { set; get; }

        public string TreatyName { set; get; }
        public string ContOrAmendmentType { set; get; }
        public DateTime EffectiveDate { set; get; }
        public string Reinsurer { set; get; }
        public string RIratio { set; get; }
        public string SignDate_Rein { set; get; }
        public string SignDate_INSH { set; get; }
        public string RIcomm { set; get; }

        public string MinNoofDeath { get; set; }
        public string LimitPerEvent { get; set; }

        public string LimitPerYear { get; set; }

        public string MinPrem { get; set; }

        public string Reinstatement { get; set; }
    }
}
