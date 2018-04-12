using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class RIContractInfo
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
        public string EffectiveDate { set; get; }
        public string Reinsurer { set; get; }
        public string RIratio { set; get; }
        public string SignDate_Rein { set; get; }
        public string SignDate_INSH { set; get; }
        public string RIcomm { set; get; }

        public List<RIContractInfo> lstChildRIContractInfo { get; set; }
        public string ContractTypeSign { get; set; }


        //public string MRTreatyName { set; get; }
        //public string MRContOrAmendmentType { set; get; }
        //public string MREffectiveDate { set; get; }
        //public string MRReinsurer { set; get; }
        //public string MRRIratio { set; get; }
        //public string MRSignDate_Rein { set; get; }
        //public string MRSignDate_INSH { set; get; }
        //public string MRRIcomm { set; get; }
        //public string HRTreatyName { set; get; }
        //public string HRContOrAmendmentType { set; get; }
        //public string HREffectiveDate { set; get; }
        //public string HRReinsurer { set; get; }
        //public string HRRIratio { set; get; }
        //public string HRSignDate { set; get; }
        //public string HRSignDate_INSH { set; get; }
        //public string HRRIcomm { set; get; }
        //public string Remark { set; get; }
        //public string RGATreatyName { set; get; }
        //public string RGAContOrAmendmentType { set; get; }
        //public string RGAEffectiveDate { set; get; }
        //public string RGAReinsurer { set; get; }
        //public string RGARIratio { set; get; }
        //public string RGASignDate_Rein { set; get; }
        //public string RGASignDate_INSH { set; get; }
        //public string RGARIcomm { set; get; }
        //public string SRTreatyName { set; get; }
        //public string SRContOrAmendmentType { set; get; }
        //public string SREffectiveDate { set; get; }
        //public string SRReinsurer { set; get; }
        //public string SRRIratio { set; get; }
        //public string SRSignDate { set; get; }
        //public string SRSignDate_INSH { set; get; }
        //public string SRRIcomm { set; get; }

    }
}
