using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class ZaiBaoProductInfo
    {
        /// <summary>
        /// 交易编码
        /// </summary>
        public string TransactionNo { get; set; }
        /// <summary>
        /// 保险机构代码
        /// </summary>
        public string CompanyCode { get; set; }
        public string ReInsuranceContNo { get; set; }
        public string ReInsuranceContName { get; set; }
        public string ReInsuranceContTitle { get; set; }
        public string MainReInsuranceContNo { get; set; }
        public string ContOrAmendmentType { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public string GPFlag { get; set; }
        public string ProductType { get; set; }
        public string LiabilityCode { get; set; }
        public string LiabilityName { get; set; }
        public string ReinsurerCode { get; set; }
        public string ReinsurerName { get; set; }
        public string ReinsuranceShare { get; set; }
        public string ReinsurMode { get; set; }
        public string ReInsuranceType { get; set; }
        public string TermType { get; set; }
        public string RetentionAmount { get; set; }
        public string RetentionPercentage { get; set; }
        public string QuotaSharePercentage { get; set; }
    }
}
