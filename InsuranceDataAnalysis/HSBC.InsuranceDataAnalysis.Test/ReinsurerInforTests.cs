using Microsoft.VisualStudio.TestTools.UnitTesting;
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.BLL;

namespace HSBC.InsuranceDataAnalysis.Utils.Tests
{
    [TestClass()]
    public class ReinsurerInforTests
    {
        [TestMethod()]
        public void GetReinsurerInforByNameTest()
        {

            Reinsurer reinsurer = new Reinsurer();
            var reinsurerInfor = reinsurer.GetReinsurerInforByName("慕尼黑再保险公司北京分公司");
            reinsurerInfor = reinsurer.GetReinsurerInforByCode("000007");
            reinsurerInfor = reinsurer.GetReinsurerInforByName("Swiss Re");
            reinsurerInfor = reinsurer.GetReinsurerInforByName("Swiss Re0");
            Assert.Fail();
        }
    }
}