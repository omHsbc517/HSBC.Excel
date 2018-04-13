using HSBC.InsuranceDataAnalysis.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class Reinsurer
    {
        private List<ReinsurerInfo> listReinsurerInfor;
        public Reinsurer()
        {
            listReinsurerInfor = new List<ReinsurerInfo>();
            listReinsurerInfor.Add(new ReinsurerInfo { ReinsurerEnglishName = "ChinaRe", ReinsurerCode = "000007", ReinsurerChineseName = "中国再保险（集团）股份有限公司" });
            listReinsurerInfor.Add(new ReinsurerInfo { ReinsurerEnglishName = "MuRe", ReinsurerCode = "000059", ReinsurerChineseName = "慕尼黑再保险公司北京分公司" });
            listReinsurerInfor.Add(new ReinsurerInfo { ReinsurerEnglishName = "HanRe", ReinsurerCode = "000128", ReinsurerChineseName = "汉诺威再保险股份公司上海分公司" });
            listReinsurerInfor.Add(new ReinsurerInfo { ReinsurerEnglishName = "RGA", ReinsurerCode = "000182", ReinsurerChineseName = "RGA美国再保险公司上海分公司" });
            listReinsurerInfor.Add(new ReinsurerInfo { ReinsurerEnglishName = "Swiss Re", ReinsurerCode = "000058", ReinsurerChineseName = "瑞士再保险股份有限公司北京分公司" });
        }

        public ReinsurerInfo GetReinsurerInforByName(string reinsurerName)
        {
            return listReinsurerInfor.Where(a => a.ReinsurerEnglishName == reinsurerName || a.ReinsurerChineseName == reinsurerName).Count()==0?null: listReinsurerInfor.Where(a => a.ReinsurerEnglishName == reinsurerName || a.ReinsurerChineseName == reinsurerName).ToList().FirstOrDefault();
        }

        public ReinsurerInfo GetReinsurerInforByCode(string reinsurerCode)
        {
            return listReinsurerInfor.Where(a => a.ReinsurerCode == reinsurerCode).ToList().Count == 0 ? null : listReinsurerInfor.Where(a => a.ReinsurerCode == reinsurerCode).ToList().FirstOrDefault();
        }


    }
}
