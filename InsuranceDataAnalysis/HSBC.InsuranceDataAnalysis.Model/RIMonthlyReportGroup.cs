using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class RIMonthlyReportGroup
    {
        /// <summary>
        /// column c
        /// </summary>
        public string ChdrNumber { get; set; }

        /// <summary>
        /// column F   保险产品编码
        /// </summary>
        public string Prodtyp { get; set; }

        /// <summary>
        /// un decide property
        /// </summary>
        public string ProductCode { get; set; }

        /// <summary>
        /// column Bf
        /// </summary>
        public string SumSi { get; set; }

        /// <summary>
        /// column AY
        /// </summary>
        public string Pprem { get; set; }

        /// <summary>
        /// column I
        /// </summary>
        public string Clntnum { get; set; }

        /// <summary>
        /// column Bk
        /// </summary>
        public string RIAnnualizedPremiumTot { get; set; }

        /// <summary>
        /// column bq
        /// </summary>
        public string RICommissionTot { get; set; }

        /// <summary>
        /// column bp
        /// </summary>
        public string ReinsuranceCommssion { get; set; }

    }
}
