using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.ExcelCore;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using System.IO;
using System.Drawing;
using HSBC.InsuranceDataAnalysis.ExcelCommon.Excel;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class ContractInfoBusiness
    {
        IExcel excelApp = new ExcelCore.ExcelCore();

        private const string RIContractName = "RI Contract Info";
        private const string ProductInfoName = "Product Info";

        private const string ContractName = "合同";
        private const string ChildContractName = "附约";

        public List<RIContractInfo> lstHanReModel
        {
            get;
            set;
        }

        public List<RIContractInfo> lstMuReModel
        {
            get;
            set;
        }

        public List<RIContractInfo> lstRGAModel
        {
            get;
            set;
        }

        public List<RIContractInfo> lstSwissReModel
        {
            get;
            set;
        }

        public List<HugeDisasterModel> lstHugeDisasterModel
        {
            get; set;
        }

        public List<ProductInfo> lstProductInfoModel
        {
            get; set;
        }

        //public List<TEMP_LMLiability> lstTEMP_LMLiabilityModel
        //{
        //    get; set;
        //}

        public List<TEMP_LMProduct> lstTEMP_LMProductModel
        {
            get; set;
        }

        public List<InsuranceReinsuranceStatement> lstInsuranceReinsuranceStatementModel
        {
            get; set;
        }

        public List<LRInsureContModel> lstLRInsureContModel
        {
            get; set;
        }

        public List<PolicyAlternationReportGroup> lstPolicyAlternationReportGroup
        {
            get; set;
        }

        public List<InforceBusinessListing> lstHR_LifeInforceBusinessListing
        {
            get; set;
        }

        public List<InforceBusinessListing> lstMR_HealthInforceBusinessListing
        {
            get; set;
        }

        public List<InforceBusinessListing> lstMR_LifeInforceBusinessListing
        {
            get; set;
        }
        public List<InforceBusinessListing> lstRGAInforceBusinessListing
        {
            get; set;
        }

        public List<InforceBusinessListing> lstSRInforceBusinessListing
        {
            get; set;
        }

        public List<InforceBusinessListing> lstInforceBusinessListing
        {
            get; set;
        }
        public List<RIMonthlyReportGroup> lstRIMonthlyReportGroup { get; set; }

        public List<TEMP_LCCont> lstTEMP_LCCont { get; set; }
        public List<TEMP_LCInsureAccTrace> lstTEMP_LCInsureAccTrace { get; set; }
        public List<TEMP_LCPolTransaction> lstTEMP_LCPolTransaction { get; set; }
        public List<ClaimSheetModel> lstClaimSheetModel
        { get; set; }
        public List<ClaimSheetModel> lstHR_LifeClaimSheetModel
        { get; set; }
        public List<ClaimSheetModel> lstMR_HealthClaimSheetModel
        { get; set; }
        public List<ClaimSheetModel> lstMR_LifeClaimSheetModel
        { get; set; }
        public List<ClaimSheetModel> lstRGAClaimSheetModel
        { get; set; }

        public List<ClaimSheetModel> lstSRClaimSheetModel
        { get; set; }

        public List<RIClaimReportGroup> lstRIClaimReportGroup
        { get; set; }

        public List<TEMP_LLClaimDetail> lstTEMP_LLClaimDetail
        { get; set; }

        public List<TEMP_LLClaimPolicy> lstTEMP_LLClaimPolicy
        { get; set; }

        public List<TEMP_LLClaimInfo> lstTEMP_LLClaimInfo
        { get; set; }


        public List<TEMP_LCGrpProduct> lstTEMP_LCGrpProduct { get; set; }

        public List<TEMP_LCProduct> lstTEMP_LCProduct { get; set; }

        public List<TEMP_LCProduct> lstTEMP_LCProductGroup { get; set; }

        public List<TEMP_LCInsureAcc> lstTEMP_LCInsureAcc { get; set; }

        public List<TEMP_LCInsured> lstTEMP_LCInsured { get; set; }

        public List<TEMP_LCInsured> lst_LCInsuredGroup { get; set; }

        public List<ZaiBaoProductInfo> lstZaiBaoProductInfo { get; set; }

        public List<LCGrpContGroup> lstLCGrpContGroup { get; set; }
        public List<TEMP_LLClaimDetail> lstLLClaimDetailGroup{ get; set; }

        public ContractInfoBusiness()
        {
            lstHanReModel = new List<RIContractInfo>();
            lstMuReModel = new List<RIContractInfo>();
            lstRGAModel = new List<RIContractInfo>();
            lstSwissReModel = new List<RIContractInfo>();
            lstHugeDisasterModel = new List<HugeDisasterModel>();
            lstProductInfoModel = new List<ProductInfo>();
            //lstTEMP_LMLiabilityModel = new List<TEMP_LMLiability>();
            lstTEMP_LMProductModel = new List<TEMP_LMProduct>();
            lstInsuranceReinsuranceStatementModel = new List<InsuranceReinsuranceStatement>();
            lstPolicyAlternationReportGroup = new List<PolicyAlternationReportGroup>();
            lstHR_LifeInforceBusinessListing = new List<InforceBusinessListing>();
            lstMR_HealthInforceBusinessListing = new List<InforceBusinessListing>();
            lstMR_LifeInforceBusinessListing = new List<InforceBusinessListing>();
            lstRGAInforceBusinessListing = new List<InforceBusinessListing>();
            lstSRInforceBusinessListing = new List<InforceBusinessListing>();
            lstInforceBusinessListing = new List<InforceBusinessListing>();
            lstRIMonthlyReportGroup = new List<RIMonthlyReportGroup>();
            lstTEMP_LCCont = new List<TEMP_LCCont>();//
            lstTEMP_LCInsureAccTrace = new List<TEMP_LCInsureAccTrace>();
            lstTEMP_LCPolTransaction = new List<TEMP_LCPolTransaction>();
            lstClaimSheetModel = new List<ClaimSheetModel>();
            lstHR_LifeClaimSheetModel = new List<ClaimSheetModel>();
            lstMR_HealthClaimSheetModel = new List<ClaimSheetModel>();
            lstMR_LifeClaimSheetModel = new List<ClaimSheetModel>();
            lstRGAClaimSheetModel = new List<ClaimSheetModel>();
            lstSRClaimSheetModel = new List<ClaimSheetModel>();
            lstRIClaimReportGroup = new List<RIClaimReportGroup>();
            lstTEMP_LLClaimDetail = new List<TEMP_LLClaimDetail>();
            lstTEMP_LLClaimPolicy = new List<TEMP_LLClaimPolicy>();
            lstTEMP_LLClaimInfo = new List<TEMP_LLClaimInfo>();
            lstTEMP_LCGrpProduct = new List<TEMP_LCGrpProduct>();
            lstTEMP_LCProduct = new List<TEMP_LCProduct>();
            lstTEMP_LCInsureAcc = new List<TEMP_LCInsureAcc>();
            lstTEMP_LCInsured = new List<TEMP_LCInsured>();
            lstZaiBaoProductInfo = new List<ZaiBaoProductInfo>();
            lstTEMP_LCProductGroup = new List<TEMP_LCProduct>();
lstLCGrpContGroup = new List<LCGrpContGroup>(); lstLLClaimDetailGroup = new List<TEMP_LLClaimDetail>();
lst_LCInsuredGroup = new List<TEMP_LCInsured>();
  }

        public void GetInformationDataFromExcel(string InformationExcelPath, string inputFilePath)
        {
            ProcessLogProxy.Normal("Start to get contract Info excel information");
            GetDataFromContractInfo(inputFilePath + @"\Contract Info.xlsx");
            GetDataFromProductInfo(inputFilePath + @"\Contract Info.xlsx");
            //GetDataFromTEMP_LMLiabilityInfo(InformationExcelPath + @"\TEMP_LMLiability.xlsx");
            GetDataFromTEMP_LMProductInfo(InformationExcelPath + @"\TEMP_LMProduct.xlsx");
            GetDataFromRIStatementStatistics(inputFilePath);
            GetInforceBusinessListingData(inputFilePath);
            GetPolicyAlternationReportGroupData(inputFilePath + @"\group\Policy alternation report-GROUP.csv");
            GetDataRIMonthlyReportGroup(inputFilePath + @"\group\RI Monthly report-GROUP.csv");
            GetRIClaimReportGroupData(inputFilePath + @"\group\RI Claim report-GROUP.csv");
            GetDataFromTEMPLCProductGroup(inputFilePath + @"\group\LCProduct_Group.xlsx");
            this.GetLCGrpContGroup(inputFilePath + @"\group\LCGrpCont_Group.xlsx");
            this.GetDataFromlstLCInsuredGroup(inputFilePath + @"\group\LCInsured_Group.xlsx");
            GetTEMP_LCInsureAccTraceData(InformationExcelPath + @"\TEMP_LCInsureAccTrace.xlsx");
            GetTEMP_LCPolTransactionData(InformationExcelPath + @"\TEMP_LCPolTransaction.xlsx");
            GetTEMP_LLClaimDetailData(InformationExcelPath + @"\TEMP_LLClaimDetail.xlsx");
            GetLLClaimDetailGroupData(InformationExcelPath + @"\group\LLClaimDetail_Group.xlsx");
            GetTEMP_LLClaimPolicyData(InformationExcelPath + @"\TEMP_LLClaimPolicy.xlsx");
            GetTEMP_LLClaimInfoData(InformationExcelPath + @"\TEMP_LLClaimInfo.xlsx");
            GetDataFromlstTEMPLCInsured(InformationExcelPath + @"\TEMP_LCInsured.xlsx");
            GetDataFromTEMPLCInsureAcc(InformationExcelPath + @"\TEMP_LCInsureAcc.xlsx");
            GetDataFromTEMPLCProduct(InformationExcelPath + @"\TEMP_LCProduct.xlsx");
            GetDataFromTEMPLCCont(InformationExcelPath + @"\TEMP_LCCont.xlsx");
            GetDataFromTEMPLCGrpProduct(InformationExcelPath + @"\TEMP_LCGrpProduct.xlsx");
            ProcessLogProxy.SuccessMessage("Get excel information Success");

        }
        private void GetDataFromContractInfo(string excelPath)
        {
            try
            {
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet(RIContractName);
                var allRows = excelApp.GetSheetByRow();

                int riContractRowsCount = allRows.Count;

                if (riContractRowsCount > 3)
                {
                    bool isHugeProduct = false;
                    for (int i = 4; i <= riContractRowsCount; i++)
                    {
                        var hugeProductCode = excelApp.GetCell(i, "B").Value;
                        if (hugeProductCode.Equals("Product Code"))
                        {
                            //巨灾保险
                            isHugeProduct = true;
                            continue;
                        }

                        if (isHugeProduct)
                        {
                            HugeDisasterModel tempHugeDisasterModel = new HugeDisasterModel();

                            tempHugeDisasterModel.ProductCode = excelApp.GetCell(i, "B").Value;
                            tempHugeDisasterModel.TypeI = excelApp.GetCell(i, "C").Value;
                            tempHugeDisasterModel.TypeII = excelApp.GetCell(i, "D").Value;
                            tempHugeDisasterModel.ReinsurerName = excelApp.GetCell(i, "E").Value;
                            tempHugeDisasterModel.BenefitReinsured = excelApp.GetCell(i, "F").Value;
                            tempHugeDisasterModel.RImethodI = excelApp.GetCell(i, "G").Value;
                            tempHugeDisasterModel.RImethodII = excelApp.GetCell(i, "H").Value;
                            tempHugeDisasterModel.Percentage = excelApp.GetCell(i, "I").Value;
                            tempHugeDisasterModel.Retention = excelApp.GetCell(i, "J").Value;

                            tempHugeDisasterModel.TreatyName = excelApp.GetCell(i, "K").Value;
                            tempHugeDisasterModel.ContOrAmendmentType = excelApp.GetCell(i, "L").Value;
                            //tempHugeDisasterModel.EffectiveDate = excelApp.GetCell(i, "M").Value;
                            tempHugeDisasterModel.EffectiveDate = this.ConvertStrToDate(excelApp.GetCell(i, "M").Value);
                            tempHugeDisasterModel.Reinsurer = excelApp.GetCell(i, "N").Value;
                            tempHugeDisasterModel.RIratio = excelApp.GetCell(i, "O").Value;
                            tempHugeDisasterModel.SignDate_Rein = excelApp.GetCell(i, "P").Value;
                            tempHugeDisasterModel.SignDate_INSH = excelApp.GetCell(i, "Q").Value;
                            tempHugeDisasterModel.RIcomm = excelApp.GetCell(i, "R").Value;

                            tempHugeDisasterModel.MinNoofDeath = excelApp.GetCell(i, "S").Value;
                            tempHugeDisasterModel.LimitPerEvent = excelApp.GetCell(i, "T").Value;
                            tempHugeDisasterModel.LimitPerYear = excelApp.GetCell(i, "U").Value;
                            tempHugeDisasterModel.MinPrem = excelApp.GetCell(i, "V").Value;
                            tempHugeDisasterModel.Reinstatement = excelApp.GetCell(i, "W").Value;
                            tempHugeDisasterModel.Remark = excelApp.GetCell(i, "Z").Value;
                            lstHugeDisasterModel.Add(tempHugeDisasterModel);
                        }
                        else
                        {
                            // 慕尼黑保险公司
                            string tempMuReContractSignName = excelApp.GetCell(i, "L").Value;
                            string tempMuReContractName = excelApp.GetCell(i, "K").Value;

                            if (!string.IsNullOrEmpty(tempMuReContractSignName)
                                && !string.IsNullOrEmpty(tempMuReContractName))
                            {
                                MuReModel tempContractInfo = new MuReModel();

                                this.CollectMuReContractData(lstMuReModel, tempContractInfo,
                                     tempMuReContractSignName, i);
                            }


                            // 汉诺威保险公司
                            string tempHanReContractSignName = excelApp.GetCell(i, "T").Value;
                            string tempHanReContractName = excelApp.GetCell(i, "S").Value;

                            if (!string.IsNullOrEmpty(tempHanReContractSignName)
                               && !string.IsNullOrEmpty(tempHanReContractName))
                            {
                                HanReModel tempContractInfo = new HanReModel();

                                this.CollectHanReContractData(lstHanReModel, tempContractInfo,
                                    tempHanReContractSignName, i);
                            }

                            // RGA美国再保险公司
                            string tempRgaContractSignName = excelApp.GetCell(i, "AC").Value;
                            string tempRgaContractName = excelApp.GetCell(i, "AB").Value;

                            if (!string.IsNullOrEmpty(tempRgaContractSignName)
                               && !string.IsNullOrEmpty(tempRgaContractName))
                            {
                                RGAModel tempContractInfo = new RGAModel();

                                this.CollectRgaReContractData(lstRGAModel, tempContractInfo,
                                    tempRgaContractSignName, i);
                            }

                            // 瑞士再保险股份有限公司
                            string tempSwissContractSignName = excelApp.GetCell(i, "AK").Value;
                            string tempSwissContractName = excelApp.GetCell(i, "AJ").Value;

                            if (!string.IsNullOrEmpty(tempSwissContractSignName)
                               && !string.IsNullOrEmpty(tempSwissContractName))
                            {
                                SwissReModel tempContractInfo = new SwissReModel();

                                this.CollectSwissReContractData(lstSwissReModel, tempContractInfo,
                                     tempSwissContractSignName, i);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        private void GetDataFromProductInfo(string excelPath)
        {
            try
            {
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet(ProductInfoName);
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    ProductInfo productInfo = new ProductInfo();
                    productInfo.ProductCode = excelApp.GetCell(i, "B").Value;
                    productInfo.ProductName = excelApp.GetCell(i, "C").Value;
                    productInfo.ProductCode1 = excelApp.GetCell(i, "D").Value;
                    productInfo.ProductType = excelApp.GetCell(i, "E").Value;
                    if (!string.IsNullOrWhiteSpace(productInfo.ProductCode))
                    {
                        lstProductInfoModel.Add(productInfo);
                    }
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        //private void GetDataFromTEMP_LMLiabilityInfo(string excelPath)
        //{
        //    try
        //    {
        //        ProcessLogProxy.Normal("Start to get TEMP_LMLiabilityInfo excel information");
        //        excelApp.OpenExcel(excelPath, true);
        //        excelApp.SelectSheet("Sheet1");
        //        var allRows = excelApp.GetSheetByRow();
        //        for (int i = 2; i <= allRows.Count; i++)
        //        {
        //            TEMP_LMLiability tEMP_LMLiability = new TEMP_LMLiability();

        //            tEMP_LMLiability.TransactionNo = excelApp.GetCell(i, "A").Value;
        //            tEMP_LMLiability.CompanyCode = excelApp.GetCell(i, "B").Value;
        //            tEMP_LMLiability.ProductCode = excelApp.GetCell(i, "C").Value;
        //            tEMP_LMLiability.ProductName = excelApp.GetCell(i, "D").Value;
        //            tEMP_LMLiability.LiabilityCode = excelApp.GetCell(i, "E").Value;
        //            tEMP_LMLiability.LiabilityName = excelApp.GetCell(i, "F").Value;
        //            tEMP_LMLiability.Classification = excelApp.GetCell(i, "G").Value;
        //            if (!string.IsNullOrWhiteSpace(tEMP_LMLiability.ProductCode))
        //            {
        //                lstTEMP_LMLiabilityModel.Add(tEMP_LMLiability);
        //            }
        //        }
        //        ProcessLogProxy.SuccessMessage("Get excel information Success");
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        excelApp.CloseExcel();
        //    }
        //}

        private void GetDataFromTEMP_LMProductInfo(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LMProductInfo excel information");
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    TEMP_LMProduct tEMP_LMProduct = new TEMP_LMProduct();
                    tEMP_LMProduct.TransactionNo = excelApp.GetCell(i, "A").Value;
                    tEMP_LMProduct.CompanyCode = excelApp.GetCell(i, "B").Value;
                    tEMP_LMProduct.ProductCode = excelApp.GetCell(i, "C").Value;
                    tEMP_LMProduct.ProductName = excelApp.GetCell(i, "D").Value;
                    tEMP_LMProduct.ProductAbbr = excelApp.GetCell(i, "E").Value;
                    tEMP_LMProduct.ProductEnName = excelApp.GetCell(i, "F").Value;
                    tEMP_LMProduct.PorductEnAbbr = excelApp.GetCell(i, "G").Value;
                    tEMP_LMProduct.InsuAccFlag = excelApp.GetCell(i, "H").Value;
                    tEMP_LMProduct.AssumIntRate = excelApp.GetCell(i, "I").Value;
                    tEMP_LMProduct.ProductType = excelApp.GetCell(i, "J").Value;
                    tEMP_LMProduct.InvestmentType = excelApp.GetCell(i, "K").Value;
                    tEMP_LMProduct.TermType = excelApp.GetCell(i, "L").Value;
                    tEMP_LMProduct.GPFlag = excelApp.GetCell(i, "M").Value;
                    tEMP_LMProduct.MainProductFlag = excelApp.GetCell(i, "N").Value;
                    tEMP_LMProduct.StopDate = excelApp.GetCell(i, "O").Value;
                    tEMP_LMProduct.SaleFlag = excelApp.GetCell(i, "P").Value;
                    tEMP_LMProduct.CircRiskCode = excelApp.GetCell(i, "Q").Value;
                    tEMP_LMProduct.ShortDurationProduct = excelApp.GetCell(i, "R").Value;
                    tEMP_LMProduct.ChangeFeeProduct = excelApp.GetCell(i, "S").Value;
                    tEMP_LMProduct.TaxRateValidDate = excelApp.GetCell(i, "T").Value;
                    tEMP_LMProduct.TaxRateExpiryDate = excelApp.GetCell(i, "U").Value;
                    tEMP_LMProduct.TaxRate = excelApp.GetCell(i, "V").Value;
                    tEMP_LMProduct.MinInterestRate = excelApp.GetCell(i, "W").Value;
                    tEMP_LMProduct.StartDate = this.ConvertStrToDate(excelApp.GetCell(i, "x").Value);
                    tEMP_LMProduct.OpenArea = excelApp.GetCell(i, "Y").Value;
                    tEMP_LMProduct.SalesChannels = excelApp.GetCell(i, "Z").Value;
                    //tEMP_LMProduct.SpecificBusiness = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.SpecificBusinessCode = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.HesitatePeriod = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.WaitingPeriod = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.LoanRatio = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.LoanPeriod = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MinAppAge = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MaxAppAge = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.AppSex = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.DivType = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MajorDiseasesNum = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MildDiseaseNum = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MajorDiseasesMaxBenefitNum = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MildDiseasesMaxBenefitNum = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MajorDiseasesBenefitType = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MajorDiseasesBenefitPolState = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.AnnuityType = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.AnnuityPeriod = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.PolicyFeeFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.InitialFeeFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.BreakThroughSocialSecurityFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.DesignatedHospitalFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.VIPClinicFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.SpecialHospital = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MedicalCoverageArea = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.DeductibleCategory = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.Deductible1 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.Deductible2 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.DeductibleRatio1 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.DeductibleRatio2 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.ClaimRatio1 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.ClaimRatio2 = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.GuaranteedRenewableFlag = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.GuaranteedRenewablePeriod = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.MaxGuaranteedRenewableAge = excelApp.GetCell(i, "A").Value;
                    //tEMP_LMProduct.OperationAllowanceFlag = excelApp.GetCell(i, "A").Value;
                    if (!string.IsNullOrWhiteSpace(tEMP_LMProduct.ProductCode))
                    {
                        lstTEMP_LMProductModel.Add(tEMP_LMProduct);
                    }
                }

                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        private void GetDataFromRIStatementStatistics(string excelPath)
        {
            ProcessLogProxy.Normal("Start to get RIStatement excel information");
            var filesPath = GetFilePath(excelPath);
            for (int i = 0; i < filesPath.Count; i++)
            {
                try
                {
                    excelApp.OpenExcel(filesPath[i], true);
                    excelApp.SelectSheet("Statement");
                    InsuranceReinsuranceStatement statementModel = new InsuranceReinsuranceStatement();
                    statementModel.ToCompanyName = excelApp.GetCell(4, "A").Value.Split(':')[1].ToString().Trim();
                    ReinsuranceParticulars DebitReinsuranceParticulars = new ReinsuranceParticulars();
                    DebitReinsuranceParticulars.ReinsurancePremiums = excelApp.GetCell(16, "B").Value;
                    statementModel.Debit = DebitReinsuranceParticulars;
                    ReinsuranceParticulars CreditReinsuranceParticulars = new ReinsuranceParticulars();
                    if (filesPath[i].Contains("_HR_"))
                    {
                        CreditReinsuranceParticulars.ReinsuranceCommissions = excelApp.GetCell(22, "C").Value;
                        CreditReinsuranceParticulars.ReinsuranceClaimAmounts = excelApp.GetCell(26, "C").Value;
                    }
                    else
                    {
                        CreditReinsuranceParticulars.ReinsuranceCommissions = excelApp.GetCell(20, "C").Value;
                        CreditReinsuranceParticulars.ReinsuranceClaimAmounts = excelApp.GetCell(24, "C").Value;
                    }
                    statementModel.FilePath = filesPath[i];
                    statementModel.Credit = CreditReinsuranceParticulars;
                    lstInsuranceReinsuranceStatementModel.Add(statementModel);
                    ProcessLogProxy.SuccessMessage("Get excel information Success");
                }
                catch (Exception ex)
                {

                    throw ex;
                }
                finally
                {
                    excelApp.CloseExcel();
                }
            }
        }

        public void GetInforceBusinessListingData(string excelPath)
        {
            ProcessLogProxy.Normal("Start to get  In force Business Listing excel information");
            var filesPath = GetFilePath(excelPath);
            for (int i = 0; i < filesPath.Count; i++)
            {
                try
                {
                    excelApp.OpenExcel(filesPath[i], true);
                    excelApp.SelectSheet("Statement");
                    var CompanyName = excelApp.GetCell(4, "A").Value.Split(':')[1].ToString().Trim();
                    excelApp.SelectSheet("In-force Business Listing");
                    var allRows = excelApp.GetSheetByRow();
                    #region Read In-force Business Listing sheet
                    for (int j = 4; j <= allRows.Count; j++)
                    {
                        InforceBusinessListing inforceBusinessListing = new InforceBusinessListing();
                        inforceBusinessListing.CompanyName = CompanyName;
                        if (!filesPath[i].Contains("MR_Health"))
                        {
                            inforceBusinessListing.PolicyNo = excelApp.GetCell(j, "A").Value;
                            string tempMemberCertificateNo = excelApp.GetCell(j, "B").Value;
                            inforceBusinessListing.MemberCertificateNo = tempMemberCertificateNo.PadLeft(8, '0');
                            inforceBusinessListing.Sex = excelApp.GetCell(j, "C").Value;
                            inforceBusinessListing.DateofBirth = excelApp.GetCell(j, "D").Value;
                            inforceBusinessListing.OccupationClass = excelApp.GetCell(j, "E").Value;
                            inforceBusinessListing.AgeofMemberWhenJoiningtheScheme = excelApp.GetCell(j, "F").Value;
                            inforceBusinessListing.ProductCode = excelApp.GetCell(j, "G").Value;
                            inforceBusinessListing.Coverage1 = excelApp.GetCell(j, "H").Value;
                            inforceBusinessListing.Attainedage = excelApp.GetCell(j, "I").Value;
                            inforceBusinessListing.ExtraMortality = excelApp.GetCell(j, "J").Value;
                            inforceBusinessListing.SumInsured = excelApp.GetCell(j, "K").Value;
                            inforceBusinessListing.InitialSumatRisk = excelApp.GetCell(j, "L").Value;
                            inforceBusinessListing.SumReinsured = excelApp.GetCell(j, "M").Value;
                            inforceBusinessListing.Retention = excelApp.GetCell(j, "N").Value;
                            inforceBusinessListing.MonthlyReinsurancePremium = excelApp.GetCell(j, "O").Value;
                            inforceBusinessListing.MonthlyReinsuranceCommission = excelApp.GetCell(j, "P").Value;
                            inforceBusinessListing.Coverage2 = excelApp.GetCell(j, "Q").Value;
                            inforceBusinessListing.ExtraMorbidity = excelApp.GetCell(j, "R").Value;
                            inforceBusinessListing.SumInsured2 = excelApp.GetCell(j, "S").Value;
                            inforceBusinessListing.InitialSumatRisk2 = excelApp.GetCell(j, "T").Value;
                            inforceBusinessListing.SumReinsured2 = excelApp.GetCell(j, "U").Value;
                            inforceBusinessListing.Retention2 = excelApp.GetCell(j, "V").Value;
                            inforceBusinessListing.MonthlyReinsurancePremium2 = excelApp.GetCell(j, "W").Value;
                            inforceBusinessListing.MonthlyReinsuranceCommission2 = excelApp.GetCell(j, "X").Value;
                            inforceBusinessListing.EffectiveDate = excelApp.GetCell(j, "Y").Value;
                            inforceBusinessListing.AutomaticorFacultative = excelApp.GetCell(j, "Z").Value;
                            inforceBusinessListing.TaxInd = excelApp.GetCell(j, "AA").Value;
                            inforceBusinessListing.RI_RATIO_1 = excelApp.GetCell(j, "AB").Value;
                            inforceBusinessListing.RI_RATIO_2 = excelApp.GetCell(j, "AC").Value;
                            inforceBusinessListing.IsMrHealth = false;
                        }
                        else
                        {
                            inforceBusinessListing.PolicyNo = excelApp.GetCell(j, "B").Value;

                            string tempMemberCertificateNo = excelApp.GetCell(j, "B").Value;
                            inforceBusinessListing.MemberCertificateNo = tempMemberCertificateNo.PadLeft(8, '0');

                            //inforceBusinessListing.MemberCertificateNo = excelApp.GetCell(j, "E").Value;
                            inforceBusinessListing.Sex = excelApp.GetCell(j, "G").Value;
                            //inforceBusinessListing.DateofBirth = excelApp.GetCell(j, "F").Value;
                            //inforceBusinessListing.OccupationClass = excelApp.GetCell(j, "E").Value;
                            //inforceBusinessListing.AgeofMemberWhenJoiningtheScheme = excelApp.GetCell(j, "F").Value;
                            inforceBusinessListing.ProductCode = excelApp.GetCell(j, "A").Value;
                            inforceBusinessListing.Coverage1 = "MI";
                            inforceBusinessListing.Attainedage = excelApp.GetCell(j, "H").Value;
                            //inforceBusinessListing.ExtraMortality = excelApp.GetCell(j, "J").Value;
                            inforceBusinessListing.SumInsured = excelApp.GetCell(j, "N").Value;
                            inforceBusinessListing.InitialSumatRisk = excelApp.GetCell(j, "O").Value;
                            inforceBusinessListing.SumReinsured = excelApp.GetCell(j, "P").Value;
                            //inforceBusinessListing.Retention = excelApp.GetCell(j, "N").Value;
                            inforceBusinessListing.MonthlyReinsurancePremium = excelApp.GetCell(j, "L").Value;
                            inforceBusinessListing.MonthlyReinsuranceCommission = excelApp.GetCell(j, "R").Value;
                            //inforceBusinessListing.Coverage2 = excelApp.GetCell(j, "Q").Value;
                            //inforceBusinessListing.ExtraMorbidity = excelApp.GetCell(j, "R").Value;
                            ////inforceBusinessListing.SumInsured	 = excelApp.GetCell(j, "S").Value;
                            ////inforceBusinessListing.InitialSumatRisk = excelApp.GetCell(j, "T").Value;
                            //inforceBusinessListing.SumReinsured = excelApp.GetCell(j, "U").Value;
                            ////inforceBusinessListing.Retention = excelApp.GetCell(j, "V").Value;
                            ////inforceBusinessListing.MonthlyReinsurancePremium = excelApp.GetCell(j, "W").Value;
                            ////inforceBusinessListing.MonthlyReinsuranceCommission	 = excelApp.GetCell(j, "X").Value;
                            //inforceBusinessListing.EffectiveDate = excelApp.GetCell(j, "Y").Value;
                            //inforceBusinessListing.AutomaticorFacultative = "A";
                            inforceBusinessListing.IsMrHealth = true;
                        }

                        if (string.IsNullOrWhiteSpace(inforceBusinessListing.PolicyNo))
                        {
                            break;
                        }
                        if (filesPath[i].Contains("HR_life"))
                        {
                            lstHR_LifeInforceBusinessListing.Add(inforceBusinessListing);
                        }
                        else if (filesPath[i].Contains("MR_Health"))
                        {
                            lstMR_HealthInforceBusinessListing.Add(inforceBusinessListing);
                        }
                        else if (filesPath[i].Contains("MR_life"))
                        {
                            lstMR_LifeInforceBusinessListing.Add(inforceBusinessListing);
                        }
                        else if (filesPath[i].Contains("_RGA"))
                        {
                            lstRGAInforceBusinessListing.Add(inforceBusinessListing);
                        }
                        else if (filesPath[i].Contains("_SR"))
                        {
                            lstSRInforceBusinessListing.Add(inforceBusinessListing);
                        }
                        lstInforceBusinessListing.Add(inforceBusinessListing);
                    }
                    #endregion read

                    excelApp.SelectSheet("Claim");

                    #region Reade Claim sheet
                    allRows = excelApp.GetSheetByRow();
                    for (int j = 4; j <= allRows.Count; j++)
                    {
                        ClaimSheetModel claimSheetModel = new ClaimSheetModel();
                        claimSheetModel.CompanyName = CompanyName;

                        claimSheetModel.Product = excelApp.GetCell(j, "A").Value;
                        claimSheetModel.PolicyNo = excelApp.GetCell(j, "B").Value;
                        claimSheetModel.GroupName = excelApp.GetCell(j, "C").Value;
                        claimSheetModel.MembersCertificateNo = excelApp.GetCell(j, "D").Value;
                        claimSheetModel.Membereffectivedate = excelApp.GetCell(j, "E").Value;
                        claimSheetModel.Memberexpire = excelApp.GetCell(j, "F").Value;
                        claimSheetModel.CauseOfClaim = excelApp.GetCell(j, "G").Value;
                        claimSheetModel.AdmissionServiceDate = excelApp.GetCell(j, "H").Value;
                        claimSheetModel.Discharge = excelApp.GetCell(j, "L").Value;
                        claimSheetModel.PaymentDate = excelApp.GetCell(j, "M").Value;
                        claimSheetModel.PaidAmount = excelApp.GetCell(j, "N").Value;
                        claimSheetModel.PaidAmountCurrency = excelApp.GetCell(j, "O").Value;
                        claimSheetModel.RecoveryAmount = excelApp.GetCell(j, "P").Value;

                        if (string.IsNullOrWhiteSpace(claimSheetModel.PolicyNo))
                        {
                            break;
                        }
                        if (filesPath[i].Contains("HR_life"))
                        {
                            if (j >= 5) lstHR_LifeClaimSheetModel.Add(claimSheetModel);
                            else continue;
                        }
                        else if (filesPath[i].Contains("MR_Health"))
                        {

                            lstMR_HealthClaimSheetModel.Add(claimSheetModel);
                        }
                        else if (filesPath[i].Contains("MR_life"))
                        {
                            lstMR_LifeClaimSheetModel.Add(claimSheetModel);
                        }
                        else if (filesPath[i].Contains("_RGA"))
                        {
                            if (j >= 5) lstRGAClaimSheetModel.Add(claimSheetModel);
                            else continue;
                        }
                        else if (filesPath[i].Contains("_SR"))
                        {
                            if (j >= 5) lstSRClaimSheetModel.Add(claimSheetModel);
                            else continue;
                        }
                        lstClaimSheetModel.Add(claimSheetModel);
                    }
                    #endregion

                    ProcessLogProxy.SuccessMessage("Get excel information Success");
                }
                catch (Exception ex)
                {

                    throw ex;
                }
                finally
                {
                    excelApp.CloseExcel();
                }
            }
        }


        public void GetPolicyAlternationReportGroupData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get Policy Alternation Report Group excel information");
                try
                {
                    CheckExcelFile(excelPath);
                }
                catch (Exception ex)
                {

                    ProcessLogProxy.Error(ex.Message);
                    return;
                }
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Policy alternation report-GROUP");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    PolicyAlternationReportGroup model = new PolicyAlternationReportGroup();
                    model.Day = excelApp.GetCell(i, "A").Value;
                    model.Chdrcoy = excelApp.GetCell(i, "B").Value;
                    model.ChdrNum = excelApp.GetCell(i, "C").Value;
                    model.ProdTyp = excelApp.GetCell(i, "F").Value;
                    model.LiabilityCode = excelApp.GetCell(i, "F").Value;
                    model.SumSi = excelApp.GetCell(i, "BS").Value;
                    model.Pprem = excelApp.GetCell(i, "AY").Value;
                    model.Clntnum = excelApp.GetCell(i, "I").Value;
                    model.ProductCode = excelApp.GetCell(i, "BE").Value;


                    if (!string.IsNullOrWhiteSpace(model.Day))
                    {
                        lstPolicyAlternationReportGroup.Add(model);
                    }
                    else
                    {
                        break;
                    }
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        private static void CheckExcelFile(string excelPath)
        {
            if (!File.Exists(excelPath))
            {
                throw new Exception("Not find excel " + excelPath.Split('\\').Last());
            }
        }

        public void GetRIClaimReportGroupData(string excelPath)
        {
            try
            {

                ProcessLogProxy.Normal("Start to get RIClaim Report Group excel information");
                try
                {
                    CheckExcelFile(excelPath);
                }
                catch (Exception ex)
                {

                    ProcessLogProxy.Error(ex.Message);
                    return;
                }
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("RI Claim report-GROUP");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    RIClaimReportGroup model = new RIClaimReportGroup();
                    model.Clamnomap = excelApp.GetCell(i, "A").Value;
                    model.Chdrnum = excelApp.GetCell(i, "B").Value;
                    model.ProdTyp = excelApp.GetCell(i, "C").Value;
                    model.Clntcoy = excelApp.GetCell(i, "D").Value;
                    model.Clntnum = excelApp.GetCell(i, "E").Value;
                    model.Gcsts = excelApp.GetCell(i, "F").Value;
                    model.Planno = excelApp.GetCell(i, "G").Value;
                    model.Gcfrpdte = excelApp.GetCell(i, "H").Value;
                    model.Gcdthclm = excelApp.GetCell(i, "I").Value;
                    model.Gcauthby = excelApp.GetCell(i, "J").Value;
                    model.Dateauth = excelApp.GetCell(i, "K").Value;
                    model.Gcoprscd = excelApp.GetCell(i, "L").Value;
                    model.Apaidamt = excelApp.GetCell(i, "M").Value;
                    model.Grskcls = excelApp.GetCell(i, "N").Value;
                    model.Claimcond = excelApp.GetCell(i, "O").Value;
                    model.Longdesc = excelApp.GetCell(i, "P").Value;
                    model.Clmadmstf = excelApp.GetCell(i, "Q").Value;
                    model.Gccauscd = excelApp.GetCell(i, "R").Value;
                    model.Payeeno = excelApp.GetCell(i, "S").Value;
                    model.Appclam = excelApp.GetCell(i, "T").Value;
                    model.Payclam = excelApp.GetCell(i, "U").Value;
                    model.Clmrcusr01 = excelApp.GetCell(i, "V").Value;
                    model.Chdholder = excelApp.GetCell(i, "W").Value;
                    model.Zbkind03 = excelApp.GetCell(i, "X").Value;
                    model.Zdistric03 = excelApp.GetCell(i, "Y").Value;
                    model.Srcebus03 = excelApp.GetCell(i, "Z").Value;
                    model.Incurrdate = excelApp.GetCell(i, "AA").Value;
                    model.Clnamemap = excelApp.GetCell(i, "AB").Value;
                    model.Cltsexmap = excelApp.GetCell(i, "AC").Value;
                    model.Cltdobmap = excelApp.GetCell(i, "AD").Value;
                    model.Agemap = excelApp.GetCell(i, "AE").Value;
                    model.Cnttype03 = excelApp.GetCell(i, "AF").Value;
                    model.Cntbranc03 = excelApp.GetCell(i, "AG").Value;
                    model.Dteatt = excelApp.GetCell(i, "AH").Value;
                    model.Dtetrm = excelApp.GetCell(i, "AI").Value;
                    model.Dtevisit = excelApp.GetCell(i, "AJ").Value;
                    model.Gcdiagcd = excelApp.GetCell(i, "AK").Value;
                    model.Dateto04 = excelApp.GetCell(i, "AL").Value;
                    model.Effdate04 = excelApp.GetCell(i, "AM").Value;
                    model.Occpcode04 = excelApp.GetCell(i, "AN").Value;
                    model.Paydate = excelApp.GetCell(i, "AO").Value;
                    model.F42 = excelApp.GetCell(i, "AP").Value;
                    model.MunichReLf = excelApp.GetCell(i, "AQ").Value;
                    model.HannoverReLf = excelApp.GetCell(i, "AR").Value;
                    model.MunichReMd = excelApp.GetCell(i, "AS").Value;
                    model.PolicyNo = excelApp.GetCell(i, "AT").Value;
                    model.ProductCode = excelApp.GetCell(i, "AU").Value;
                    if (!string.IsNullOrWhiteSpace(model.Clamnomap))
                    {
                        lstRIClaimReportGroup.Add(model);
                    }
                    else
                    {
                        break;
                    }
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        #region shangjunqi modify
        private void GetDataRIMonthlyReportGroup(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get RI Monthly Report Group excel information");
                try
                {
                    CheckExcelFile(excelPath);
                }
                catch (Exception ex)
                {

                    ProcessLogProxy.Error(ex.Message);
                    return;
                }
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("RI Monthly report-GROUP");
                var allRows = excelApp.GetSheetByRow();

                int riMonthlyReportGroupCount = allRows.Count;

                if (riMonthlyReportGroupCount > 1)
                {
                    for (int i = 2; i <= riMonthlyReportGroupCount; i++)
                    {
                        RIMonthlyReportGroup tempModel = new RIMonthlyReportGroup();

                        tempModel.ChdrNumber = excelApp.GetCell(i, "C").Value;
                        tempModel.Mbrno = excelApp.GetCell(i, "D").Value;
                        tempModel.Prodtyp = excelApp.GetCell(i, "F").Value;

                        tempModel.ProductCode = excelApp.GetCell(i, "BE").Value;

                        tempModel.SumSi = excelApp.GetCell(i, "bF").Value;
                        tempModel.Pprem = excelApp.GetCell(i, "AY").Value;
                        tempModel.Clntnum = excelApp.GetCell(i, "I").Value;
                        tempModel.RIAnnualizedPremiumTot = excelApp.GetCell(i, "BK").Value;
                        tempModel.RICommissionTot = excelApp.GetCell(i, "BQ").Value;
                        tempModel.ReinsuranceCommssion = excelApp.GetCell(i, "BP").Value;

                        lstRIMonthlyReportGroup.Add(tempModel);
                    }
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }
        #endregion


        /// <summary>
        /// 1.6.5 TEMP_LCInsureAccTrace
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetTEMP_LCInsureAccTraceData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LCInsureAccTrace excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                int riMonthlyReportGroupCount = allRows.Count;
                if (riMonthlyReportGroupCount > 1)
                {
                    for (int i = 2; i <= riMonthlyReportGroupCount; i++)
                    {
                        TEMP_LCInsureAccTrace tempModel = new TEMP_LCInsureAccTrace();

                        tempModel.PolicyNo = excelApp.GetCell(i, "F").Value;
                        if (string.IsNullOrWhiteSpace(tempModel.PolicyNo))
                        {
                            break;
                        }
                        lstTEMP_LCInsureAccTrace.Add(tempModel);
                    }
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        /// <summary>
        /// 1.6.5 TEMP_LCPolTransaction
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetTEMP_LCPolTransactionData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LCPolTransaction excel information");
                CheckExcelFile(excelPath);

                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCPolTransaction = _excel.Read<TEMP_LCPolTransaction>(excel).ToList();
                //excelApp.OpenExcel(excelPath, true);
                //excelApp.SelectSheet("Sheet1");
                //var allRows = excelApp.GetSheetByRow();
                //for (int i = 2; i <= allRows.Count; i++)
                //{
                //    TEMP_LCPolTransaction tempModel = new TEMP_LCPolTransaction();
                //    tempModel.GrpPolicyNo = excelApp.GetCell(i, "D").Value;
                //    tempModel.PolicyNo = excelApp.GetCell(i, "E").Value;
                //    tempModel.EndorAcceptNo = excelApp.GetCell(i, "R").Value;
                //    tempModel.EndorsementNo = excelApp.GetCell(i, "S").Value;
                //    if (string.IsNullOrWhiteSpace(tempModel.PolicyNo))
                //    {
                //        break;
                //    }
                //    lstTEMP_LCPolTransaction.Add(tempModel);
                //}
                excel.Close();
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void GetLCGrpContGroup(string excelPath)
        {
            ProcessLogProxy.Normal("Start to get LCGrpContGroup excel information");
            CheckExcelFile(excelPath);

            var _excel = new ExcelHelper();
            ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
            lstLCGrpContGroup = _excel.Read<LCGrpContGroup>(excel).ToList();
            excel.Close();
            ProcessLogProxy.SuccessMessage("Get excel information Success");
        }

        /// <summary>
        /// 1.6.6 TEMP_LLClaimDetail 
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetTEMP_LLClaimDetailData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LLClaimDetail excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    TEMP_LLClaimDetail tempModel = new TEMP_LLClaimDetail();
                    tempModel.ClmCaseNo = excelApp.GetCell(i, "C").Value;
                    tempModel.GrpPolicyNo = excelApp.GetCell(i, "E").Value;
                    tempModel.PolicyNo = excelApp.GetCell(i, "G").Value;
                    tempModel.GetLiabilityCode = excelApp.GetCell(i, "M").Value;
                    tempModel.GetLiabilityName = excelApp.GetCell(i, "O").Value;
                    tempModel.BenefitType = excelApp.GetCell(i, "J").Value;
                    tempModel.DeductibleType = excelApp.GetCell(i, "W").Value;
                    tempModel.Deductible = excelApp.GetCell(i, "X").Value;
                    tempModel.ClaimRatio = excelApp.GetCell(i, "Y").Value;
                    if (string.IsNullOrWhiteSpace(tempModel.PolicyNo))
                    {
                        break;
                    }
                    lstTEMP_LLClaimDetail.Add(tempModel);
                }

                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        /// <summary>
        /// 1.6.6 TEMP_LLClaimPolicy
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetTEMP_LLClaimPolicyData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LLClaimPolicy excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    TEMP_LLClaimPolicy tempModel = new TEMP_LLClaimPolicy();

                    tempModel.ClaimNo = excelApp.GetCell(i, "C").Value;
                    tempModel.PolicyNo = excelApp.GetCell(i, "G").Value;
                    tempModel.PayStatusCode = excelApp.GetCell(i, "Y").Value;
                    if (string.IsNullOrWhiteSpace(tempModel.PolicyNo))
                    {
                        break;
                    }
                    lstTEMP_LLClaimPolicy.Add(tempModel);
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        /// <summary>
        /// 1.6.6  TEMP_LLClaimInfo
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetTEMP_LLClaimInfoData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LLClaimInfo excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    TEMP_LLClaimInfo tempModel = new TEMP_LLClaimInfo();
                    tempModel.ClaimNo = excelApp.GetCell(i, "E").Value;
                    tempModel.AccidentDate = excelApp.GetCell(i, "AD").Value;
                    tempModel.ClmSettDate = excelApp.GetCell(i, "BM").Value;
                    if (string.IsNullOrWhiteSpace(tempModel.ClaimNo))
                    {
                        break;
                    }
                    lstTEMP_LLClaimInfo.Add(tempModel);
                }
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        /// <summary>
        /// 1.6.6 LLClaimDetailGroup 
        /// </summary>
        /// <param name="excelPath"></param>
        private void GetLLClaimDetailGroupData(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMP_LLClaimDetail excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("Sheet1");
                var allRows = excelApp.GetSheetByRow();
                for (int i = 2; i <= allRows.Count; i++)
                {
                    TEMP_LLClaimDetail tempModel = new TEMP_LLClaimDetail();
                    tempModel.ClmCaseNo = excelApp.GetCell(i, "C").Value;
                    tempModel.GrpPolicyNo = excelApp.GetCell(i, "E").Value;
                    tempModel.PolicyNo = excelApp.GetCell(i, "G").Value;
                    tempModel.GetLiabilityCode = excelApp.GetCell(i, "M").Value;
                    tempModel.GetLiabilityName = excelApp.GetCell(i, "O").Value;
                    tempModel.BenefitType = excelApp.GetCell(i, "J").Value;
                    tempModel.DeductibleType = excelApp.GetCell(i, "W").Value;
                    tempModel.Deductible = excelApp.GetCell(i, "X").Value;
                    tempModel.ClaimRatio = excelApp.GetCell(i, "Y").Value;
                    if (string.IsNullOrWhiteSpace(tempModel.PolicyNo))
                    {
                        break;
                    }
                    lstLLClaimDetailGroup.Add(tempModel);
                }

                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        #region shangjunqi ADD
        private void GetDataFromTEMPLCGrpProduct(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCGrpProduct excel information");
                CheckExcelFile(excelPath);
                excelApp.OpenExcel(excelPath, true);
                excelApp.SelectSheet("团体险种表");
                var allRows = excelApp.GetSheetByRow();

                int tempLCGrpProductCount = allRows.Count;

                if (tempLCGrpProductCount > 1)
                {
                    for (int i = 2; i <= tempLCGrpProductCount; i++)
                    {
                        TEMP_LCGrpProduct tempModel = new TEMP_LCGrpProduct();

                        tempModel.GrpProductNo = excelApp.GetCell(i, "F").Value;

                        lstTEMP_LCGrpProduct.Add(tempModel);
                    }
                }

                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.CloseExcel();
            }
        }

        private void GetDataFromTEMPLCCont(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCCont excel information");

                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCCont = _excel.Read<TEMP_LCCont>(excel).ToList();


                //CheckExcelFile(excelPath);
                //excelApp.OpenExcel(excelPath, true);
                //excelApp.SelectSheet("Sheet1");
                //var allRows = excelApp.GetSheetByRow();

                //int tempLCContCount = allRows.Count;

                //if (tempLCContCount > 1)
                //{
                //    for (int i = 2; i <= tempLCContCount; i++)
                //    {
                //        TEMP_LCCont tempModel = new TEMP_LCCont();

                //        tempModel.GrpPolicyNo = excelApp.GetCell(i, "C").Value;
                //        tempModel.PolicyNo = excelApp.GetCell(i, "D").Value;
                //        tempModel.RenewalTimes = excelApp.GetCell(i, "BR").Value;
                //        tempModel.ManageCom = excelApp.GetCell(i, "H").Value;
                //        tempModel.SignDate = excelApp.GetCell(i, "AE").Value;
                //        tempModel.Premium = excelApp.GetCell(i, "AH").Value;

                //        lstTEMP_LCCont.Add(tempModel);
                //    }
                //}
                ProcessLogProxy.SuccessMessage("Get excel information Success");
                excel.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetDataFromTEMPLCProduct(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCProduct excel information");

                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCProduct = _excel.Read<TEMP_LCProduct>(excel).ToList();
                excel.Close();
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetDataFromTEMPLCProductGroup(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCProductGroup excel information");

                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCProductGroup = _excel.Read<TEMP_LCProduct>(excel).ToList();
                excel.Close();
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetDataFromTEMPLCInsureAcc(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCInsureAcc excel information");

                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCInsureAcc = _excel.Read<TEMP_LCInsureAcc>(excel).ToList();

                //CheckExcelFile(excelPath);
                //excelApp.OpenExcel(excelPath, true);
                //excelApp.SelectSheet("Sheet1");
                //var allRows = excelApp.GetSheetByRow();

                //int allRowCount = allRows.Count;

                //if (allRowCount > 1)
                //{
                //    for (int i = 2; i <= allRowCount; i++)
                //    {
                //        TEMP_LCInsureAcc tempModel = new TEMP_LCInsureAcc();

                //        tempModel.AccountValue = excelApp.GetCell(i, "Q").Value;
                //        tempModel.PolicyNo = excelApp.GetCell(i, "E").Value;
                //        tempModel.ProductNo = excelApp.GetCell(i, "G").Value;

                //        lstTEMP_LCInsureAcc.Add(tempModel);
                //    }
                //}
                ProcessLogProxy.SuccessMessage("Get excel information Success");
                excel.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GetDataFromlstTEMPLCInsured(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get TEMPLCInsured excel information");
                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lstTEMP_LCInsured = _excel.Read<TEMP_LCInsured>(excel).ToList();
                excel.Close();
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void GetDataFromlstLCInsuredGroup(string excelPath)
        {
            try
            {
                ProcessLogProxy.Normal("Start to get LCInsured_Group excel information");
                var _excel = new ExcelHelper();
                ExcelReflectionHelper excel = new ExcelReflectionHelper(false, excelPath);
                lst_LCInsuredGroup = _excel.Read<TEMP_LCInsured>(excel).ToList();
                excel.Close();
                ProcessLogProxy.SuccessMessage("Get excel information Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion

        private void CollectMuReContractData(List<RIContractInfo> lstContainer, RIContractInfo currentModel,
            string tempMuReContractSignName, int rowIndex)
        {
            string tempProductCode = excelApp.GetCell(rowIndex, "B").Value;

            //if (!string.IsNullOrEmpty(tempProductCode))
            //{
            currentModel.ProductCode = tempProductCode;
            currentModel.TypeI = excelApp.GetCell(rowIndex, "C").Value;
            currentModel.TypeII = excelApp.GetCell(rowIndex, "D").Value;
            currentModel.ReinsurerName = excelApp.GetCell(rowIndex, "E").Value;
            currentModel.BenefitReinsured = excelApp.GetCell(rowIndex, "F").Value;
            currentModel.RImethodI = excelApp.GetCell(rowIndex, "G").Value;
            currentModel.RImethodII = excelApp.GetCell(rowIndex, "H").Value;
            currentModel.Percentage = excelApp.GetCell(rowIndex, "I").Value;
            currentModel.Retention = excelApp.GetCell(rowIndex, "J").Value;

            currentModel.TreatyName = excelApp.GetCell(rowIndex, "K").Value;
            currentModel.ContOrAmendmentType = excelApp.GetCell(rowIndex, "L").Value;
            currentModel.EffectiveDate = excelApp.GetCell(rowIndex, "M").Value;
            currentModel.Reinsurer = excelApp.GetCell(rowIndex, "N").Value;
            currentModel.RIratio = excelApp.GetCell(rowIndex, "O").Value;
            currentModel.SignDate_Rein = excelApp.GetCell(rowIndex, "P").Value;
            currentModel.SignDate_INSH = excelApp.GetCell(rowIndex, "Q").Value;
            currentModel.RIcomm = excelApp.GetCell(rowIndex, "R").Value;
            //}

            //  合同
            if (tempMuReContractSignName.Equals(ContractName))
            {
                currentModel.ContractTypeSign = "M";
                currentModel.lstChildRIContractInfo = new List<RIContractInfo>();
                lstContainer.Add(currentModel);
            }
            else
            {
                // 附约
                if (lstContainer.Count > 0)
                {
                    currentModel.ContractTypeSign = "T";
                    lstContainer[lstContainer.Count - 1].lstChildRIContractInfo.Add(currentModel);
                }
            }
        }

        private void CollectHanReContractData(List<RIContractInfo> lstContainer, RIContractInfo currentModel,
           string tempHanReContractSignName, int rowIndex)
        {
            string tempProductCode = excelApp.GetCell(rowIndex, "B").Value;

            //if (!string.IsNullOrEmpty(tempProductCode))
            //{
            currentModel.ProductCode = tempProductCode;
            currentModel.TypeI = excelApp.GetCell(rowIndex, "C").Value;
            currentModel.TypeII = excelApp.GetCell(rowIndex, "D").Value;
            currentModel.ReinsurerName = excelApp.GetCell(rowIndex, "E").Value;
            currentModel.BenefitReinsured = excelApp.GetCell(rowIndex, "F").Value;
            currentModel.RImethodI = excelApp.GetCell(rowIndex, "G").Value;
            currentModel.RImethodII = excelApp.GetCell(rowIndex, "H").Value;
            currentModel.Percentage = excelApp.GetCell(rowIndex, "I").Value;
            currentModel.Retention = excelApp.GetCell(rowIndex, "J").Value;

            currentModel.TreatyName = excelApp.GetCell(rowIndex, "S").Value;
            currentModel.ContOrAmendmentType = excelApp.GetCell(rowIndex, "T").Value;
            currentModel.EffectiveDate = excelApp.GetCell(rowIndex, "U").Value;
            currentModel.Reinsurer = excelApp.GetCell(rowIndex, "V").Value;
            currentModel.RIratio = excelApp.GetCell(rowIndex, "W").Value;
            currentModel.SignDate_Rein = excelApp.GetCell(rowIndex, "X").Value;
            currentModel.SignDate_INSH = excelApp.GetCell(rowIndex, "Y").Value;
            currentModel.RIcomm = excelApp.GetCell(rowIndex, "Z").Value;
            //}

            //  合同
            if (tempHanReContractSignName.Equals(ContractName))
            {
                currentModel.lstChildRIContractInfo = new List<RIContractInfo>();
                currentModel.ContractTypeSign = "M";
                lstContainer.Add(currentModel);
            }
            else
            {
                // 附约
                if (lstContainer.Count > 0)
                {
                    currentModel.ContractTypeSign = "T";
                    lstContainer[lstContainer.Count - 1].lstChildRIContractInfo.Add(currentModel);
                }
            }
        }

        private void CollectRgaReContractData(List<RIContractInfo> lstContainer, RIContractInfo currentModel,
           string tempRgaReContractSignName, int rowIndex)
        {
            string tempProductCode = excelApp.GetCell(rowIndex, "B").Value;

            //if (!string.IsNullOrEmpty(tempProductCode))
            //{
            currentModel.ProductCode = tempProductCode;
            currentModel.TypeI = excelApp.GetCell(rowIndex, "C").Value;
            currentModel.TypeII = excelApp.GetCell(rowIndex, "D").Value;
            currentModel.ReinsurerName = excelApp.GetCell(rowIndex, "E").Value;
            currentModel.BenefitReinsured = excelApp.GetCell(rowIndex, "F").Value;
            currentModel.RImethodI = excelApp.GetCell(rowIndex, "G").Value;
            currentModel.RImethodII = excelApp.GetCell(rowIndex, "H").Value;
            currentModel.Percentage = excelApp.GetCell(rowIndex, "I").Value;
            currentModel.Retention = excelApp.GetCell(rowIndex, "J").Value;

            currentModel.TreatyName = excelApp.GetCell(rowIndex, "AB").Value;
            currentModel.ContOrAmendmentType = excelApp.GetCell(rowIndex, "AC").Value;
            currentModel.EffectiveDate = excelApp.GetCell(rowIndex, "AD").Value;
            currentModel.Reinsurer = excelApp.GetCell(rowIndex, "AE").Value;
            currentModel.RIratio = excelApp.GetCell(rowIndex, "AF").Value;
            currentModel.SignDate_Rein = excelApp.GetCell(rowIndex, "AG").Value;
            currentModel.SignDate_INSH = excelApp.GetCell(rowIndex, "AH").Value;
            currentModel.RIcomm = excelApp.GetCell(rowIndex, "AI").Value;
            //}

            //  合同
            if (tempRgaReContractSignName.Equals(ContractName))
            {
                currentModel.ContractTypeSign = "M";
                currentModel.lstChildRIContractInfo = new List<RIContractInfo>();
                lstContainer.Add(currentModel);
            }
            else
            {
                // 附约
                if (lstContainer.Count > 0)
                {
                    currentModel.ContractTypeSign = "T";
                    lstContainer[lstContainer.Count - 1].lstChildRIContractInfo.Add(currentModel);
                }
            }
        }

        private void CollectSwissReContractData(List<RIContractInfo> lstContainer, RIContractInfo currentModel,
           string tempSwissReContractSignName, int rowIndex)
        {
            string tempProductCode = excelApp.GetCell(rowIndex, "B").Value;

            //if (!string.IsNullOrEmpty(tempProductCode))
            //{
            currentModel.ProductCode = tempProductCode;
            currentModel.TypeI = excelApp.GetCell(rowIndex, "C").Value;
            currentModel.TypeII = excelApp.GetCell(rowIndex, "D").Value;
            currentModel.ReinsurerName = excelApp.GetCell(rowIndex, "E").Value;
            currentModel.BenefitReinsured = excelApp.GetCell(rowIndex, "F").Value;
            currentModel.RImethodI = excelApp.GetCell(rowIndex, "G").Value;
            currentModel.RImethodII = excelApp.GetCell(rowIndex, "H").Value;
            currentModel.Percentage = excelApp.GetCell(rowIndex, "I").Value;
            currentModel.Retention = excelApp.GetCell(rowIndex, "J").Value;

            currentModel.TreatyName = excelApp.GetCell(rowIndex, "AJ").Value;
            currentModel.ContOrAmendmentType = excelApp.GetCell(rowIndex, "AK").Value;
            currentModel.EffectiveDate = excelApp.GetCell(rowIndex, "AL").Value;
            currentModel.Reinsurer = excelApp.GetCell(rowIndex, "AM").Value;
            currentModel.RIratio = excelApp.GetCell(rowIndex, "AN").Value;
            currentModel.SignDate_Rein = excelApp.GetCell(rowIndex, "AO").Value;
            currentModel.SignDate_INSH = excelApp.GetCell(rowIndex, "AP").Value;
            currentModel.RIcomm = excelApp.GetCell(rowIndex, "AQ").Value;
            //}

            //  合同
            if (tempSwissReContractSignName.Equals(ContractName))
            {
                currentModel.ContractTypeSign = "M";
                currentModel.lstChildRIContractInfo = new List<RIContractInfo>();
                lstContainer.Add(currentModel);
            }
            else
            {
                // 附约
                if (lstContainer.Count > 0)
                {
                    currentModel.ContractTypeSign = "T";
                    lstContainer[lstContainer.Count - 1].lstChildRIContractInfo.Add(currentModel);
                }
            }
        }

        private DateTime ConvertStrToDate(string str)
        {
            DateTime result;
            //bool convertResult = DateTime.TryParse(str, out result);
            result = DateTime.Parse(str);
            //if (convertResult)
            //{
            //    return result;
            //}
            //else
            //{
            //    return null;
            //}
            return result;
        }

        private List<string> GetFilePath(string filePath)
        {
            List<string> filesPath = new List<string>();
            DirectoryInfo dir = new DirectoryInfo(filePath);
            foreach (FileInfo file in dir.GetFiles("RI Statement & Statistics*", SearchOption.TopDirectoryOnly))//第二个参数表示搜索包含子目录中的文件；
            {
                filesPath.Add(filePath + "//" + file.Name);
            }
            return filesPath;
        }


        private string GetDateStr(string date)
        {
            string result = string.Empty;
            if (!string.IsNullOrEmpty(date))
            {
                result = DateTime.Parse(date.Trim()).ToString("yyyy-MM-dd");
            }
            return result;
        }
    }
}
