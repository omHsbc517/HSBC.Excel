using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "AJ")]
    public class LCGrpContGroup
    {
        /// <summary>
        /// column C
        /// </summary>
        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { get; set; }

        /// <summary>
        /// column I
        /// </summary>
        [Description("ManageCom")]
        public string ManageCom { get; set; }

        /// <summary>
        /// column T
        /// </summary>
        [Description("SignDate")]
        public string SignDate { get; set; }

        /// <summary>
        /// column AJ
        /// </summary>
        [Description("EffDate")]
        public string EffDate { get; set; }
    }
}
