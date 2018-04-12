using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class InsuranceReinsuranceStatement
    {
        public string FilePath{set;get;}
        public string ToCompanyName { set; get; }// 慕尼黑再保险公司北京分公司
        public string Period { set; get; }
        public string Currency { set; get; }
        public string ContractType { set; get; }
        public ReinsuranceParticulars Debit { set; get; }
        public ReinsuranceParticulars Credit { set; get; }
    }
}
