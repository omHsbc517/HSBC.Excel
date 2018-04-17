using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "AU")]
    public class TEMP_LCProduct
    {
        /// <summary>
        /// 团体保单号
        /// </summary>
        [Description("GrpPolicyNo")]
        public string GrpPolicyNo { get; set; }

        private string policyNo;
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

        /// <summary>
        /// 个单保险险种号码
        /// </summary>
        [Description("ProductNo")]
        public string ProductNo { get; set; }

        /// <summary>
        /// 产品编码
        /// </summary>
        [Description("ProductCode")]
        public string ProductCode { get; set; }

        /// <summary>
        /// 主险保险险种号码
        /// </summary>
        [Description("MainProductNo")]
        public string MainProductNo { get; set; }

        /// <summary>
        /// 主附险性质代码
        /// </summary>
        [Description("MainProductFlag")]
        public string MainProductFlag { get; set; }

        /// <summary>
        /// 保险责任生效日期 
        /// </summary>
        [Description("EffDate")]
        public string EffDate { get; set; }

        /// <summary>
        /// 保险责任终止日期
        /// </summary>
        [Description("InvalidDate")]
        public string InvalidDate { get; set; }

        /// <summary>
        /// 核保结论代码
        /// </summary>
        [Description("UWConclusion")]
        public string UWConclusion { get; set; }

        /// <summary>
        /// 职业加费金额
        /// </summary>
        [Description("ProfessionalFee")]
        public string ProfessionalFee { get; set; }

        /// <summary>
        /// 次标准体加费金额
        /// </summary>
        [Description("SubStandardFee")]
        public string SubStandardFee { get; set; }

        /// <summary>
        /// EM加点
        /// </summary>
        [Description("EMRate")]
        public string EMRate { get; set; }

        /// <summary>
        /// 基本保额
        /// </summary>
        [Description("BasicSumInsured")]
        public string BasicSumInsured { get; set; }

        /// <summary>
        /// 风险保额
        /// </summary>
        [Description("RiskAmnt")]
        public string RiskAmnt { get; set; }

        /// <summary>
        /// 保费
        /// </summary>
        [Description("Premium")]
        public string Premium { get; set; }
    }
}
