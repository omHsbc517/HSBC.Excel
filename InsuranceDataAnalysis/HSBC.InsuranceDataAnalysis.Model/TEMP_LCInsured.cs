
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "AD")]
    public class TEMP_LCInsured
    {
        /// <summary>
        /// 被保人客户编号
        /// </summary>
        /// 
        [Description("InsuredNo")]
        public string InsuredNo { get; set; }

        /// <summary>
        /// 个人保单号
        /// </summary>
        [Description("PolicyNo")]
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

        /// <summary>
        /// column c
        /// </summary>
        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { get; set; }


        /// <summary>
        /// 被保人姓名
        /// </summary>
        [Description("InsuredName")]
        public string InsuredName { get; set; }

        /// <summary>
        /// 被保人性别
        /// </summary>
        [Description("InsuredSex")]
        public string InsuredSex { get; set; }

        /// <summary>
        /// 被保人证件类型
        /// </summary>
        [Description("InsuredCertType")]
        public string InsuredCertType { get; set; }

        /// <summary>
        /// 被保人证件编码
        /// </summary>
        [Description("InsuredCertNo")]
        public string InsuredCertNo { get; set; }

        /// <summary>
        /// 职业代码
        /// </summary>
        [Description("OccupationType")]
        public string OccupationType { get; set; }

        /// <summary>
        /// 投保年龄
        /// </summary>
        [Description("AppAge")]
        public string AppAge { get; set; }
    }
}
