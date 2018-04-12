using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    [Sheet("Sheet1", "Q")]
    public class TEMP_LCInsureAcc
    {
        /// <summary>
        /// 保险账户价值
        /// </summary>
        [Description("AccountValue")]
        public string AccountValue { get; set; }

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
        /// 个单保险险种号码
        /// </summary>
        [Description("ProductNo")]
        public string ProductNo { get; set; }
    }
}