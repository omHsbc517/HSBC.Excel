using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.Model
{
    public class ReinsuranceParticulars
    {
        public string ReinsurancePremiums { set; get; } // (再保费)

        public string ReturnPremiumForLapses { set; get; }//(再保费退费)

        public string ReinsuranceCommissions { set; get; }//(再保手续费)

        public string ReturnCommissionForLapses { set; get; }//(再保手续费退费)

        public string ReinsuranceClaimAmounts { set; get; }//(再保索赔)

        public string ProfitCommission { set; get; }//(盈余佣金）

        public string Total { set; get; }
        public string BalanceDueToByTheReinsurer { set; get; }//

    }
}
