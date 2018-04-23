using HSBC.InsuranceDataAnalysis.ExcelCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.Utils;
namespace HSBC.InsuranceDataAnalysis.BLL
{

    public class ExcelTemplate
    {
        public void CreateTemplate(IExcel excelApp, string filePath, ExcelTemplateName excelTemplateName)
        {
            try
            {
                CreateExcel(excelApp, filePath);
                excelApp.OpenExcel(filePath, false);
                if (excelTemplateName == ExcelTemplateName.LRProduct) CreateLRProductSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LRInsureCont) CreateLRInsureContSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LRAccount) CreateLRAccountSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LRCont) CreateLRContSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LRClaim) CreateLRClaimSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LREdor) CreateLREdorSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LJInvoiceRelation) CreateLJInvoiceRelationSheet(excelApp);
                if (excelTemplateName == ExcelTemplateName.LJInvoice) CreateLJInvoiceSheet(excelApp);
                excelApp.Save();
                excelApp.CloseExcel();
            }
            catch (Exception ex)
            {
                throw new Exception("Create excel Template error " + ex.Message);
            }
        }

        #region private method
        private void CreateExcel(IExcel excelApp, string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            excelApp.CreateExcel(filePath);
            // AddNewSheet(filePath);
        }

        private void CreateLRProductSheet(IExcel excelApp)
        {

            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "D", "ReInsuranceContName");
            excelApp.SetCellValue(1, "E", "ReInsuranceContTitle");
            excelApp.SetCellValue(1, "F", "MainReInsuranceContNo");
            excelApp.SetCellValue(1, "G", "ContOrAmendmentType");
            excelApp.SetCellValue(1, "H", "ProductCode");
            excelApp.SetCellValue(1, "I", "ProductName");
            excelApp.SetCellValue(1, "J", "GPFlag");
            excelApp.SetCellValue(1, "K", "ProductType");
            excelApp.SetCellValue(1, "L", "LiabilityCode");
            excelApp.SetCellValue(1, "M", "LiabilityName");
            excelApp.SetCellValue(1, "N", "ReinsurerCode");
            excelApp.SetCellValue(1, "O", "ReinsurerName");
            excelApp.SetCellValue(1, "P", "ReinsuranceShare");
            excelApp.SetCellValue(1, "Q", "ReinsurMode");
            excelApp.SetCellValue(1, "R", "ReInsuranceType");
            excelApp.SetCellValue(1, "S", "TermType");
            excelApp.SetCellValue(1, "T", "RetentionAmount");
            excelApp.SetCellValue(1, "U", "RetentionPercentage");
            excelApp.SetCellValue(1, "V", "QuotaSharePercentage");

            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 10);
            excelApp.SetColumnTextType("Sheet1", 14);
            excelApp.Save();
        }

        private void CreateLRInsureContSheet(IExcel excelApp)
        {
            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "D", "ReInsuranceContName");
            excelApp.SetCellValue(1, "E", "ReInsuranceContTitle");
            excelApp.SetCellValue(1, "F", "MainReInsuranceContNo");
            excelApp.SetCellValue(1, "G", "ContOrAmendmentType");
            excelApp.SetCellValue(1, "H", "ContAttribute");
            excelApp.SetCellValue(1, "I", "ContStatus");
            excelApp.SetCellValue(1, "J", "TreatyOrFacultativeFlag");
            excelApp.SetCellValue(1, "K", "ContSigndate");
            excelApp.SetCellValue(1, "L", "PeriodFrom");
            excelApp.SetCellValue(1, "M", "PeriodTo");
            excelApp.SetCellValue(1, "N", "ContType");
            excelApp.SetCellValue(1, "O", "ReinsurerCode");
            excelApp.SetCellValue(1, "P", "ReinsurerName");
            excelApp.SetCellValue(1, "Q", "ChargeType");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 15);
            excelApp.SetColumnDateType("Sheet1", 11);
            excelApp.SetColumnDateType("Sheet1", 12);
            excelApp.SetColumnDateType("Sheet1", 13);
            excelApp.Save();
        }

        private void CreateLRAccountSheet(IExcel excelApp)
        {
            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "AccountID");
            excelApp.SetCellValue(1, "D", "AccountingPeriodfrom");
            excelApp.SetCellValue(1, "E", "AccountingPeriodto");
            excelApp.SetCellValue(1, "F", "ReinsurerCode");
            excelApp.SetCellValue(1, "G", "ReinsurerName");
            excelApp.SetCellValue(1, "H", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "I", "ReInsuranceContName");
            excelApp.SetCellValue(1, "J", "Currency");
            excelApp.SetCellValue(1, "K", "ReinsurancePremium");
            excelApp.SetCellValue(1, "L", "ReinsuranceCommssionRate");
            excelApp.SetCellValue(1, "M", "ReinsuranceCommssion");
            excelApp.SetCellValue(1, "N", "ReturnReinsurancePremium");
            excelApp.SetCellValue(1, "O", "ReturnReinsuranceCommssion");
            excelApp.SetCellValue(1, "P", "ReturnSurrenderPay");
            excelApp.SetCellValue(1, "Q", "ReturnClaimPay");
            excelApp.SetCellValue(1, "R", "ReturnMaturity");
            excelApp.SetCellValue(1, "S", "ReturnAnnuity");
            excelApp.SetCellValue(1, "T", "ReturnLivBene");
            excelApp.SetCellValue(1, "U", "AccountStatus");
            excelApp.SetCellValue(1, "V", "PairingStatus");
            excelApp.SetCellValue(1, "W", "PairingDate");
            excelApp.SetCellValue(1, "X", "CurrentRate");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 6);

            excelApp.SetColumnDateType("Sheet1", 4);
            excelApp.SetColumnDateType("Sheet1", 5);
            excelApp.SetColumnDateType("Sheet1", 23);

            excelApp.Save();
        }

        private void CreateLJInvoiceRelationSheet(IExcel excelApp)
        {
            excelApp.SetCellValue("Sheet1", 1, "A", "BusiNo");
            excelApp.SetCellValue("Sheet1", 1, "B", "CompanyCode");
            excelApp.SetCellValue("Sheet1", 1, "C", "GrpPolicyNo");
            excelApp.SetCellValue("Sheet1", 1, "D", "PolicyNo");
            excelApp.SetCellValue("Sheet1", 1, "E", "GPFlag");
            excelApp.SetCellValue("Sheet1", 1, "F", "BussinessType");
            excelApp.SetCellValue("Sheet1", 1, "G", "BussinessCode");
            excelApp.SetCellValue("Sheet1", 1, "H", "InvoiceNo");
            excelApp.SetCellValue("Sheet1", 1, "I", "InvoiceCode");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.Save();

        }

        private void CreateLJInvoiceSheet(IExcel excelApp)
        {
            excelApp.SetCellValue("Sheet1", 1, "A", "BusiNo");
            excelApp.SetCellValue("Sheet1", 1, "B", "CompanyCode");
            excelApp.SetCellValue("Sheet1", 1, "C", "InvoiceNo");
            excelApp.SetCellValue("Sheet1", 1, "D", "InvoiceCode");
            excelApp.SetCellValue("Sheet1", 1, "E", "TaxCompanyCode");
            excelApp.SetCellValue("Sheet1", 1, "F", "Drawer");
            excelApp.SetCellValue("Sheet1", 1, "G", "InvoiceClass");
            excelApp.SetCellValue("Sheet1", 1, "H", "InvoiceType");
            excelApp.SetCellValue("Sheet1", 1, "I", "InvoiceAmount");
            excelApp.SetCellValue("Sheet1", 1, "J", "TaxAmount");
            excelApp.SetCellValue("Sheet1", 1, "K", "SubFeeCode");
            excelApp.SetCellValue("Sheet1", 1, "L", "SubFeeName");
            excelApp.SetCellValue("Sheet1", 1, "M", "SubFee");
            excelApp.SetCellValue("Sheet1", 1, "N", "SubTaxRate");
            excelApp.SetCellValue("Sheet1", 1, "O", "SubTaxAmount");
            excelApp.SetCellValue("Sheet1", 1, "P", "ProductCode");
            excelApp.SetCellValue("Sheet1", 1, "Q", "ProductName");
            excelApp.SetCellValue("Sheet1", 1, "R", "TaxpayerName");
            excelApp.SetCellValue("Sheet1", 1, "S", "TaxpayerId");
            excelApp.SetCellValue("Sheet1", 1, "T", "TaxpayerAddress");
            excelApp.SetCellValue("Sheet1", 1, "U", "TaxpayerType");
            excelApp.SetCellValue("Sheet1", 1, "V", "TaxpayerPhone");
            excelApp.SetCellValue("Sheet1", 1, "W", "TaxpayerBankCode");
            excelApp.SetCellValue("Sheet1", 1, "X", "TaxpayerBankAccount");
            excelApp.SetCellValue("Sheet1", 1, "Y", "Currency");
            excelApp.SetCellValue("Sheet1", 1, "Z", "InvioceState");
            excelApp.SetCellValue("Sheet1", 1, "AA", "InvoicePrintDate");
            excelApp.Save();
        }

        private void AddNewSheet(string excelFilePath, IExcel excelApp)
        {
            excelApp.AddNewSheet(excelFilePath, "LRProduct");
            excelApp.AddNewSheet(excelFilePath, "LRInsureCont");
            excelApp.AddNewSheet(excelFilePath, "LRAccount");
            excelApp.AddNewSheet(excelFilePath, "LJInvoiceRelation");
            excelApp.AddNewSheet(excelFilePath, "LJInvoice");
        }


        private void CreateLRContSheet(IExcel excelApp)
        {
            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "GrpPolicyNo");
            excelApp.SetCellValue(1, "D", "GrpProductNo");
            excelApp.SetCellValue(1, "E", "PolicyNo");
            excelApp.SetCellValue(1, "F", "ProductNo");
            excelApp.SetCellValue(1, "G", "GPFlag");
            excelApp.SetCellValue(1, "H", "MainProductNo");
            excelApp.SetCellValue(1, "I", "MainProductFlag");
            excelApp.SetCellValue(1, "J", "ProductCode");
            excelApp.SetCellValue(1, "K", "LiabilityCode");
            excelApp.SetCellValue(1, "L", "LiabilityName");
            excelApp.SetCellValue(1, "M", "Classification");
            excelApp.SetCellValue(1, "N", "EventType");
            excelApp.SetCellValue(1, "O", "RenewalTimes");
            excelApp.SetCellValue(1, "P", "TermType");
            excelApp.SetCellValue(1, "Q", "ManageCom");
            excelApp.SetCellValue(1, "R", "SignDate");
            excelApp.SetCellValue(1, "S", "EffDate");
            excelApp.SetCellValue(1, "T", "PolYear");
            excelApp.SetCellValue(1, "U", "InvalidDate");
            excelApp.SetCellValue(1, "V", "UWConclusion");
            excelApp.SetCellValue(1, "W", "PolStatus");
            excelApp.SetCellValue(1, "X", "Status");
            excelApp.SetCellValue(1, "Y", "BasicSumInsured");
            excelApp.SetCellValue(1, "Z", "RiskAmnt");
            excelApp.SetCellValue(1, "AA", "Premium");
            excelApp.SetCellValue(1, "AB", "AccountValue");
            excelApp.SetCellValue(1, "AC", "FacultativeFlag");
            excelApp.SetCellValue(1, "AD", "AnonymousFlag");
            excelApp.SetCellValue(1, "AE", "WaiverFlag");
            excelApp.SetCellValue(1, "AF", "WaiverPrem");
            excelApp.SetCellValue(1, "AG", "FinalCashValue");
            excelApp.SetCellValue(1, "AH", "FinalLiabilityReserve");
            excelApp.SetCellValue(1, "AI", "InsuredNo");
            excelApp.SetCellValue(1, "AJ", "InsuredName");
            excelApp.SetCellValue(1, "AK", "InsuredSex");
            excelApp.SetCellValue(1, "AL", "InsuredCertType");
            excelApp.SetCellValue(1, "AM", "InsuredCertNo");
            excelApp.SetCellValue(1, "AN", "OccupationType");
            excelApp.SetCellValue(1, "AO", "AppntAge");
            excelApp.SetCellValue(1, "AP", "PreAge");
            excelApp.SetCellValue(1, "AQ", "ProfessionalFee");
            excelApp.SetCellValue(1, "AR", "SubStandardFee");
            excelApp.SetCellValue(1, "AS", "EMRate");
            excelApp.SetCellValue(1, "AT", "ProjectFlag");
            excelApp.SetCellValue(1, "AU", "InsurePeoples");
            excelApp.SetCellValue(1, "AV", "SaparateFlag");
            excelApp.SetCellValue(1, "AW", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "AX", "ReinsurerCode");
            excelApp.SetCellValue(1, "AY", "ReinsurerName");
            excelApp.SetCellValue(1, "AZ", "ReinsurMode");
            excelApp.SetCellValue(1, "BA", "ReinsuranceAmnt");
            excelApp.SetCellValue(1, "BB", "RetentionAmount");
            excelApp.SetCellValue(1, "BC", "Currency");
            excelApp.SetCellValue(1, "BD", "QuotaSharePercentage");
            excelApp.SetCellValue(1, "BE", "ReinsurancePremium");
            excelApp.SetCellValue(1, "BF", "ReinsuranceCommssion");
            excelApp.SetCellValue(1, "BG", "ReComputationsDate");
            excelApp.SetCellValue(1, "BH", "AccountGetDate");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 3);
            excelApp.SetColumnTextType("Sheet1", 4);
            excelApp.SetColumnTextType("Sheet1", 5);
            excelApp.SetColumnTextType("Sheet1", 6);
            excelApp.SetColumnTextType("Sheet1", 7);
            excelApp.SetColumnTextType("Sheet1", 8);
            excelApp.SetColumnTextType("Sheet1", 14);
            excelApp.SetColumnTextType("Sheet1", 17);
            excelApp.SetColumnTextType("Sheet1", 23);
            excelApp.SetColumnTextType("Sheet1", 24);
            excelApp.SetColumnTextType("Sheet1", 35);
            excelApp.SetColumnTextType("Sheet1", 49);
            excelApp.SetColumnTextType("Sheet1", 50);

            excelApp.SetColumnDateType("Sheet1", 18);
            excelApp.SetColumnDateType("Sheet1", 19);
            excelApp.SetColumnDateType("Sheet1", 21);
            excelApp.SetColumnDateType("Sheet1", 59);
            excelApp.SetColumnDateType("Sheet1", 60);


            excelApp.Save();
        }




        private void CreateLREdorSheet(IExcel excelApp)
        {
            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "GrpPolicyNo");
            excelApp.SetCellValue(1, "D", "GrpProductNo");
            excelApp.SetCellValue(1, "E", "PolicyNo");
            excelApp.SetCellValue(1, "F", "ProductNo");
            excelApp.SetCellValue(1, "G", "GPFlag");
            excelApp.SetCellValue(1, "H", "MainProductNo");
            excelApp.SetCellValue(1, "I", "MainProductFlag");
            excelApp.SetCellValue(1, "J", "ProductCode");
            excelApp.SetCellValue(1, "K", "LiabilityCode");
            excelApp.SetCellValue(1, "L", "LiabilityName");
            excelApp.SetCellValue(1, "M", "Classification");
            excelApp.SetCellValue(1, "N", "TermType");
            excelApp.SetCellValue(1, "O", "ManageCom");
            excelApp.SetCellValue(1, "P", "SignDate");
            excelApp.SetCellValue(1, "Q", "EffDate");
            excelApp.SetCellValue(1, "R", "PolYear");
            excelApp.SetCellValue(1, "S", "InvalidDate");
            excelApp.SetCellValue(1, "T", "UWConclusion");
            excelApp.SetCellValue(1, "U", "PolStatus");
            excelApp.SetCellValue(1, "V", "Status");
            excelApp.SetCellValue(1, "W", "BasicSumInsured");
            excelApp.SetCellValue(1, "X", "RiskAmnt");
            excelApp.SetCellValue(1, "Y", "Premium");
            excelApp.SetCellValue(1, "Z", "AccountValue");
            excelApp.SetCellValue(1, "AA", "FacultativeFlag");
            excelApp.SetCellValue(1, "AB", "AnonymousFlag");
            excelApp.SetCellValue(1, "AC", "WaiverFlag");
            excelApp.SetCellValue(1, "AD", "WaiverPrem");
            excelApp.SetCellValue(1, "AE", "FinalCashValue");
            excelApp.SetCellValue(1, "AF", "FinalLiabilityReserve");
            excelApp.SetCellValue(1, "AG", "InsuredNo");
            excelApp.SetCellValue(1, "AH", "InsuredName");
            excelApp.SetCellValue(1, "AI", "InsuredSex");
            excelApp.SetCellValue(1, "AJ", "InsuredCertType");
            excelApp.SetCellValue(1, "AK", "InsuredCertNo");
            excelApp.SetCellValue(1, "AL", "OccupationType");
            excelApp.SetCellValue(1, "AM", "AppntAge");
            excelApp.SetCellValue(1, "AN", "PreAge");
            excelApp.SetCellValue(1, "AO", "ProfessionalFee");
            excelApp.SetCellValue(1, "AP", "SubStandardFee");
            excelApp.SetCellValue(1, "AQ", "EMRate");
            excelApp.SetCellValue(1, "AR", "ProjectFlag");
            excelApp.SetCellValue(1, "AS", "InsurePeoples");
            excelApp.SetCellValue(1, "AT", "EndorAcceptNo");
            excelApp.SetCellValue(1, "AU", "EndorsementNo");
            excelApp.SetCellValue(1, "AV", "EdorType");
            excelApp.SetCellValue(1, "AW", "EdorValiDate");
            excelApp.SetCellValue(1, "AX", "EdorConfDate");
            excelApp.SetCellValue(1, "AY", "EdorMoney");
            excelApp.SetCellValue(1, "AZ", "SaparateFlag");
            excelApp.SetCellValue(1, "BA", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "BB", "ReinsurerCode");
            excelApp.SetCellValue(1, "BC", "ReinsurerName");
            excelApp.SetCellValue(1, "BD", "ReinsurMode");
            excelApp.SetCellValue(1, "BE", "QuotaSharePercentage");
            excelApp.SetCellValue(1, "BF", "PreInsuredAge");
            excelApp.SetCellValue(1, "BG", "PreBasicSumInsured");
            excelApp.SetCellValue(1, "BH", "PreRiskAmnt");
            excelApp.SetCellValue(1, "BI", "PreReinsuranceAmnt");
            excelApp.SetCellValue(1, "BJ", "PreRetentionAmount");
            excelApp.SetCellValue(1, "BK", "PrePremium");
            excelApp.SetCellValue(1, "BL", "PreAccountValue");
            excelApp.SetCellValue(1, "BM", "PreWaiverPrem");
            excelApp.SetCellValue(1, "BN", "ProjectAcreageChange");
            excelApp.SetCellValue(1, "BO", "ProjectCostChange");
            excelApp.SetCellValue(1, "BP", "ReinsuranceAmntChange");
            excelApp.SetCellValue(1, "BQ", "RetentionAmount");
            excelApp.SetCellValue(1, "BR", "ReinsurancePremiumChange");
            excelApp.SetCellValue(1, "BS", "ReinsuranceCommssionChange");
            excelApp.SetCellValue(1, "BT", "Currency");
            excelApp.SetCellValue(1, "BU", "ReComputationsDate");
            excelApp.SetCellValue(1, "BV", "AccountGetDate");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 3);
            excelApp.SetColumnTextType("Sheet1", 4);
            excelApp.SetColumnTextType("Sheet1", 5);
            excelApp.SetColumnTextType("Sheet1", 6);
            excelApp.SetColumnTextType("Sheet1", 7);
            excelApp.SetColumnTextType("Sheet1", 8);
            excelApp.SetColumnTextType("Sheet1", 9);
            excelApp.SetColumnTextType("Sheet1", 15);
            excelApp.SetColumnTextType("Sheet1", 21);
            excelApp.SetColumnTextType("Sheet1", 22);
            excelApp.SetColumnTextType("Sheet1", 53);
            excelApp.SetColumnTextType("Sheet1", 54);
            excelApp.SetColumnTextType("Sheet1", 33);

            excelApp.SetColumnDateType("Sheet1", 16);
            excelApp.SetColumnDateType("Sheet1", 17);
            excelApp.SetColumnDateType("Sheet1", 19);
            excelApp.SetColumnDateType("Sheet1", 49);
            excelApp.SetColumnDateType("Sheet1", 50);
            excelApp.SetColumnDateType("Sheet1", 73);
            excelApp.SetColumnDateType("Sheet1", 74);

            excelApp.Save();
        }



        private void CreateLRClaimSheet(IExcel excelApp)
        {
            excelApp.SetCellValue(1, "A", "BusiNo");
            excelApp.SetCellValue(1, "B", "CompanyCode");
            excelApp.SetCellValue(1, "C", "GrpPolicyNo");
            excelApp.SetCellValue(1, "D", "GrpProductNo");
            excelApp.SetCellValue(1, "E", "PolicyNo");
            excelApp.SetCellValue(1, "F", "ProductNo");
            excelApp.SetCellValue(1, "G", "GPFlag");
            excelApp.SetCellValue(1, "H", "MainProductNo");
            excelApp.SetCellValue(1, "I", "MainProductFlag");
            excelApp.SetCellValue(1, "J", "ProductCode");
            excelApp.SetCellValue(1, "K", "LiabilityCode");
            excelApp.SetCellValue(1, "L", "LiabilityName");
            excelApp.SetCellValue(1, "M", "GetLiabilityCode");
            excelApp.SetCellValue(1, "N", "GetLiabilityName");
            excelApp.SetCellValue(1, "O", "BenefitType");
            excelApp.SetCellValue(1, "P", "TermType");
            excelApp.SetCellValue(1, "Q", "ManageCom");
            excelApp.SetCellValue(1, "R", "SignDate");
            excelApp.SetCellValue(1, "S", "EffDate");
            excelApp.SetCellValue(1, "T", "PolYear");
            excelApp.SetCellValue(1, "U", "InvalidDate");
            excelApp.SetCellValue(1, "V", "UWConclusion");
            excelApp.SetCellValue(1, "W", "PolStatus");
            excelApp.SetCellValue(1, "X", "Status");
            excelApp.SetCellValue(1, "Y", "BasicSumInsured");
            excelApp.SetCellValue(1, "Z", "RiskAmnt");
            excelApp.SetCellValue(1, "AA", "Premium");
            excelApp.SetCellValue(1, "AB", "DeductibleType");
            excelApp.SetCellValue(1, "AC", "Deductible");
            excelApp.SetCellValue(1, "AD", "ClaimRatio");
            excelApp.SetCellValue(1, "AE", "AccountValue");
            excelApp.SetCellValue(1, "AF", "FacultativeFlag");
            excelApp.SetCellValue(1, "AG", "AnonymousFlag");
            excelApp.SetCellValue(1, "AH", "WaiverFlag"); 
            excelApp.SetCellValue(1, "AI", "WaiverPrem");
            excelApp.SetCellValue(1, "AJ", "FinalCashValue");
            excelApp.SetCellValue(1, "AK", "InsuredNo");
            excelApp.SetCellValue(1, "AL", "InsuredName");
            excelApp.SetCellValue(1, "AM", "InsuredSex");
            excelApp.SetCellValue(1, "AN", "InsuredCertType");
            excelApp.SetCellValue(1, "AO", "InsuredCertNo");
            excelApp.SetCellValue(1, "AP", "OccupationType");
            excelApp.SetCellValue(1, "AQ", "AppntAge");
            excelApp.SetCellValue(1, "AR", "PreAge");
            excelApp.SetCellValue(1, "AS", "FinalLiabilityReserve");
            excelApp.SetCellValue(1, "AT", "ProfessionalFee");
            excelApp.SetCellValue(1, "AU", "SubStandardFee");
            excelApp.SetCellValue(1, "AV", "EMRate");
            excelApp.SetCellValue(1, "AW", "ProjectFlag");
            excelApp.SetCellValue(1, "AX", "InsurePeoples");
            excelApp.SetCellValue(1, "AY", "SaparateFlag");
            excelApp.SetCellValue(1, "AZ", "ReInsuranceContNo");
            excelApp.SetCellValue(1, "BA", "ReinsurerCode");
            excelApp.SetCellValue(1, "BB", "ReinsurerName");
            excelApp.SetCellValue(1, "BC", "ReinsurMode");
            excelApp.SetCellValue(1, "BD", "ReinsuranceAmnt");
            excelApp.SetCellValue(1, "BE", "RetentionAmount");
            excelApp.SetCellValue(1, "BF", "QuotaSharePercentage");
            excelApp.SetCellValue(1, "BG", "ClaimNo");
            excelApp.SetCellValue(1, "BH", "AccidentDate");
            excelApp.SetCellValue(1, "BI", "ClmSettDate");
            excelApp.SetCellValue(1, "BJ", "PayStatusCode");
            excelApp.SetCellValue(1, "BK", "ClaimMoney");
            excelApp.SetCellValue(1, "BL", "BackClaimMoney");
            excelApp.SetCellValue(1, "BM", "BackDate");
            excelApp.SetCellValue(1, "BN", "Currency");
            excelApp.SetCellValue(1, "BO", "ReComputationsDate");
            excelApp.SetCellValue(1, "BP", "AccountGetDate");
            excelApp.SetColumnTextType("Sheet1", 1);
            excelApp.SetColumnTextType("Sheet1", 2);
            excelApp.SetColumnTextType("Sheet1", 3);
            excelApp.SetColumnTextType("Sheet1", 4);
            excelApp.SetColumnTextType("Sheet1", 5);
            excelApp.SetColumnTextType("Sheet1", 6);
            excelApp.SetColumnTextType("Sheet1", 7);
            excelApp.SetColumnTextType("Sheet1", 8);
            excelApp.SetColumnTextType("Sheet1", 9);
            excelApp.SetColumnTextType("Sheet1", 17);
            excelApp.SetColumnTextType("Sheet1", 23);
            excelApp.SetColumnTextType("Sheet1", 24);
            excelApp.SetColumnTextType("Sheet1", 52);
            excelApp.SetColumnTextType("Sheet1", 53);
            excelApp.SetColumnTextType("Sheet1", 37);
            excelApp.SetColumnTextType("Sheet1", 59);

            excelApp.SetColumnDateType("Sheet1", 18);
            excelApp.SetColumnDateType("Sheet1", 19);
            excelApp.SetColumnDateType("Sheet1", 21);
            excelApp.SetColumnDateType("Sheet1", 60);
            excelApp.SetColumnDateType("Sheet1", 65);
            excelApp.SetColumnDateType("Sheet1", 67);
            excelApp.SetColumnDateType("Sheet1", 68);
            excelApp.SetColumnDateType("Sheet1", 61);

            excelApp.SetColumnDecimalsType("Sheet1", 63);
            excelApp.SetColumnDecimalsType("Sheet1", 64);

            excelApp.Save();
        }



        #endregion
    }

}
