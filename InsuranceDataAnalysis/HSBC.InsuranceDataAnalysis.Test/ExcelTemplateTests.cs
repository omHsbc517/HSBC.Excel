using Microsoft.VisualStudio.TestTools.UnitTesting;
using HSBC.InsuranceDataAnalysis.BLL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.Utils;
using HSBC.InsuranceDataAnalysis.ExcelCore;

namespace HSBC.InsuranceDataAnalysis.BLL.Tests
{
    [TestClass()]
    public class ExcelTemplateTests
    {
        [TestMethod()]
        public void CreateTemplateTest()
        {
            try
            {
                string path = @"C:\Users\Administrator\Desktop\aaaa";
                ExcelTemplate lrProduct = new ExcelTemplate();
                IExcel excelApp = new ExcelCore.ExcelCore();
                lrProduct.CreateTemplate(excelApp,path + @"\TEMP_" + ExcelTemplateName.LRProduct + ".xlsx", ExcelTemplateName.LRProduct);
                lrProduct.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LRInsureCont + ".xlsx", ExcelTemplateName.LRInsureCont);
                lrProduct.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LRAccount + ".xlsx", ExcelTemplateName.LRAccount);
                lrProduct.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LJInvoiceRelation + ".xlsx", ExcelTemplateName.LJInvoiceRelation);
                lrProduct.CreateTemplate(excelApp, path + @"\TEMP_" + ExcelTemplateName.LJInvoice + ".xlsx", ExcelTemplateName.LJInvoice);
               
            }
            catch (Exception ex )
            {
                throw ex;
            }
        }
    }
}