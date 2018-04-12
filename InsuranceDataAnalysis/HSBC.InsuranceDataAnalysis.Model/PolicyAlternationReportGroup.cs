using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class PolicyAlternationReportGroup
    {

        public string Day { set; get; }
        public string Chdrcoy { set; get; }
        public string ChdrNum{ set; get; }
        public string  ProdTyp { set; get; }
        public string LiabilityCode { set; get; }
        public string SumSi { get; set; }//BS
        public string Pprem { get; set; }//AY
        public string Clntnum { get; set; }//I

        public string ProductCode { get; set; }//I
    }
}
