using Microsoft.VisualStudio.TestTools.UnitTesting;
using HSBC.InsuranceDataAnalysis.BLL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL.Tests
{
    [TestClass()]
    public class ContractInfoBusinessTests
    {

        [TestMethod()]
        public void GetInformationDataFromExcelTest()
        {
            try
            {
                string path = @"C:\Users\Administrator\Desktop\template\template\yyyymm\Contract Info.xlsx";
                string path1 = @"C:\Users\Administrator\Desktop\template\template\inputPath2\TEMP_LMLiability.xlsx";
                string path2 = @"C:\Users\Administrator\Desktop\template\template\inputPath2\TEMP_LMProduct.xlsx";
                ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
                contractInfoBusiness.GetInformationDataFromExcel(path1, @"C:\Users\Administrator\Desktop\template\template\yyyymm");
                Assert.Fail();
            }
            catch (Exception EX)
            {

                throw;
            }
        }

        [TestMethod()]
        public void GetInforceBusinessListingDataTest()
        {
            ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
           // contractInfoBusiness.GetInforceBusinessListingData(@"C:\Users\Administrator\Desktop\template00\template\yyyymm");

            Assert.Fail();
        }

        [TestMethod()]
        public void GetPolicyAlternationReportGroupDataTest()
        {
            ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
           // contractInfoBusiness.GetPolicyAlternationReportGroupData(@"C:\Users\Administrator\Desktop\template00\template\yyyymm\group");

            Assert.Fail();
          
        }
    }
}