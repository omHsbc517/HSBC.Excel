using Microsoft.VisualStudio.TestTools.UnitTesting;
using HSBC.InsuranceDataAnalysis.BLL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using HSBC.InsuranceDataAnalysis.Utils;
using HSBC.InsuranceDataAnalysis.ExcelCore;
using System.Configuration;
namespace HSBC.InsuranceDataAnalysis.BLL.Tests
{
    [TestClass()]
    public class LRInsureContTests
    {
        [TestMethod()]
        public void WriteLRInsureContSheetTest()
        {
            try
            {
                LRInsureCont lRInsureCont = new LRInsureCont();
                ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
                string outPutFilePath = @"C:\Users\Administrator\Desktop\v20180309.xlsx";
                string TEMP_LMLiabilityInfoExcelPath = @"C:\Users\Administrator\Desktop\template\template\inputPath2";
                contractInfoBusiness.GetInformationDataFromExcel( TEMP_LMLiabilityInfoExcelPath, @"C:\Users\Administrator\Desktop\template\template\yyyymm");
                ExcelTemplate excelTemplate = new ExcelTemplate();
                //  excelTemplate.CreateTemplate(outPutFilePath);
                lRInsureCont.WriteLRInsureContSheet(contractInfoBusiness, outPutFilePath, "20170131");
                LRAccount lRAccount = new LRAccount();
                lRAccount.WriteLRAccountSheet(contractInfoBusiness, outPutFilePath, "201701");
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        [TestMethod()]
        public void GetFiles()
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(@"C:\Users\Administrator\Desktop\template\template\yyyymm");
                foreach (FileInfo file in dir.GetFiles("RI Statement & Statistics*", SearchOption.TopDirectoryOnly))//第二个参数表示搜索包含子目录中的文件；
                {
                    var fileName = file.Name;
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        [TestMethod()]
        public void WriteLRInsureContSheet()
        {
            try
            {
                string path = @"C:\Users\Administrator\Desktop\template\template\yyyymm\output";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                IExcel excelApp = new ExcelCore.ExcelCore();
                excelTemplate.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LRProduct + ".xlsx", ExcelTemplateName.LRProduct);
                excelTemplate.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LRInsureCont + ".xlsx", ExcelTemplateName.LRInsureCont);
                excelTemplate.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LRAccount + ".xlsx", ExcelTemplateName.LRAccount);
                excelTemplate.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LJInvoiceRelation + ".xlsx", ExcelTemplateName.LJInvoiceRelation);
                excelTemplate.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LJInvoice + ".xlsx", ExcelTemplateName.LJInvoice);
                LRInsureCont lRInsureCont = new LRInsureCont();
                ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
                string TEMP_LMLiabilityInfoExcelPath = @"C:\Users\Administrator\Desktop\template\template\inputPath2";
                contractInfoBusiness.GetInformationDataFromExcel(TEMP_LMLiabilityInfoExcelPath, @"C:\Users\Administrator\Desktop\template\template\yyyymm");
                lRInsureCont.WriteLRInsureContSheet(contractInfoBusiness, path + @"\TEMP_" + ExcelTemplateName.LRInsureCont + ".xlsx", "20170131");
                LRAccount lRAccount = new LRAccount();
                lRAccount.WriteLRAccountSheet(contractInfoBusiness, path + @"\TEMP_" + ExcelTemplateName.LRAccount + ".xlsx", "201701");
            }
            catch (Exception ex)
            {

                throw;
            }
        }



      

    }
}