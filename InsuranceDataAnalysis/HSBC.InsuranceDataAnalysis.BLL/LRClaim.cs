using HSBC.InsuranceDataAnalysis.ExcelCore;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LRClaim
    {
        List<LRClaimModel> lRClaimModelList = new List<LRClaimModel>();
        private string origanizationCode = "000131";

        Reinsurer reinsurer = new Reinsurer();

        public void WriteLRClaimSheet(ContractInfoBusiness contractInfoBusiness, string OutPutFolderPath, string dateyyyymm)
        {
            IExcel excelApp = new ExcelCore.ExcelCore();
            int serialNumber = 1;
            try
            {
                ProcessLogProxy.Normal("Start building Claim excel");
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LRClaim + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LRClaim);//创建模板
                GetGroupLRClaimData(contractInfoBusiness, dateyyyymm, ref serialNumber);
                GetLRClaimData(contractInfoBusiness, dateyyyymm, ref serialNumber);
                excelApp.OpenExcel(excelPath, false);
                for (int i = 0; i < lRClaimModelList.Count; i++)
                {
                    var model = lRClaimModelList[i];
                    excelApp.SetCellValue(i + 2, "A", model.TransactionNo);
                    excelApp.SetCellValue(i + 2, "B", model.CompanyCode);
                    excelApp.SetCellValue(i + 2, "C", model.GrpPolicyNo);
                    excelApp.SetCellValue(i + 2, "D", model.GrpProductNo);
                    excelApp.SetCellValue(i + 2, "E", model.PolicyNo);
                    excelApp.SetCellValue(i + 2, "F", model.ProductNo);
                    excelApp.SetCellValue(i + 2, "G", model.GPFlag);
                    excelApp.SetCellValue(i + 2, "H", model.MainProductNo);
                    excelApp.SetCellValue(i + 2, "I", model.MainProductFlag);
                    excelApp.SetCellValue(i + 2, "J", model.ProductCode);
                    excelApp.SetCellValue(i + 2, "K", model.LiabilityCode);
                    excelApp.SetCellValue(i + 2, "L", model.LiabilityName);
                    excelApp.SetCellValue(i + 2, "M", model.GetLiabilityCode);
                    excelApp.SetCellValue(i + 2, "N", model.GetLiabilityName);
                    excelApp.SetCellValue(i + 2, "O", model.BenefitType);
                    excelApp.SetCellValue(i + 2, "P", model.TermType);
                    excelApp.SetCellValue(i + 2, "Q", model.ManageCom);
                    excelApp.SetCellValue(i + 2, "R", model.SignDate);
                    excelApp.SetCellValue(i + 2, "S", model.EffDate);
                    excelApp.SetCellValue(i + 2, "T", model.PolYear);
                    excelApp.SetCellValue(i + 2, "U", model.InvalidDate);
                    excelApp.SetCellValue(i + 2, "V", model.UWConclusion);
                    excelApp.SetCellValue(i + 2, "W", model.PolStatus);
                    excelApp.SetCellValue(i + 2, "X", model.Status);
                    excelApp.SetCellValue(i + 2, "Y", model.BasicSumInsured);
                    excelApp.SetCellValue(i + 2, "Z", model.RiskAmnt);
                    excelApp.SetCellValue(i + 2, "AA", model.Premium);
                    excelApp.SetCellValue(i + 2, "AB", model.DeductibleType);
                    excelApp.SetCellValue(i + 2, "AC", model.Deductible);
                    excelApp.SetCellValue(i + 2, "AD", model.ClaimRatio);
                    excelApp.SetCellValue(i + 2, "AE", model.AccountValue);
                    excelApp.SetCellValue(i + 2, "AF", model.FacultativeFlag);
                    excelApp.SetCellValue(i + 2, "AG", model.AnonymousFlag);
                    excelApp.SetCellValue(i + 2, "AH", model.WaiverFlag);
                    excelApp.SetCellValue(i + 2, "AI", model.WaiverPrem);
                    excelApp.SetCellValue(i + 2, "AJ", model.FinalCashValue);
                    excelApp.SetCellValue(i + 2, "AK", model.InsuredNo);
                    excelApp.SetCellValue(i + 2, "AL", model.InsuredName);
                    excelApp.SetCellValue(i + 2, "AM", model.InsuredSex);
                    excelApp.SetCellValue(i + 2, "AN", model.InsuredCertType);
                    excelApp.SetCellValue(i + 2, "AO", model.InsuredCertNo);
                    excelApp.SetCellValue(i + 2, "AP", model.OccupationType);
                    excelApp.SetCellValue(i + 2, "AQ", model.AppntAge);
                    excelApp.SetCellValue(i + 2, "AR", model.PreAge);
                    excelApp.SetCellValue(i + 2, "AS", model.FinalLiabilityReserve);
                    excelApp.SetCellValue(i + 2, "AT", model.ProfessionalFee);
                    excelApp.SetCellValue(i + 2, "AU", model.SubStandardFee);
                    excelApp.SetCellValue(i + 2, "AV", model.EMRate);
                    excelApp.SetCellValue(i + 2, "AW", model.ProjectFlag);
                    excelApp.SetCellValue(i + 2, "AX", model.InsurePeoples);
                    excelApp.SetCellValue(i + 2, "AY", model.SaparateFlag);
                    excelApp.SetCellValue(i + 2, "AZ", model.ReInsuranceContNo);
                    excelApp.SetCellValue(i + 2, "BA", model.ReinsurerCode);
                    excelApp.SetCellValue(i + 2, "BB", model.ReinsurerName);
                    excelApp.SetCellValue(i + 2, "BC", model.ReinsurMode);
                    excelApp.SetCellValue(i + 2, "BD", model.ReinsuranceAmnt);
                    excelApp.SetCellValue(i + 2, "BE", model.RetentionAmount);
                    excelApp.SetCellValue(i + 2, "BF", model.QuotaSharePercentage);
                    excelApp.SetCellValue(i + 2, "BG", model.ClaimNo);
                    excelApp.SetCellValue(i + 2, "BH", model.AccidentDate);
                    excelApp.SetCellValue(i + 2, "BI", model.ClmSettDate);
                    excelApp.SetCellValue(i + 2, "BJ", model.PayStatusCode);
                    excelApp.SetCellValue(i + 2, "BK", model.ClaimMoney);
                    excelApp.SetCellValue(i + 2, "BL", model.BackClaimMoney);
                    excelApp.SetCellValue(i + 2, "BM", model.BackDate);
                    excelApp.SetCellValue(i + 2, "BN", model.Currency);
                    excelApp.SetCellValue(i + 2, "BO", model.ReComputationsDate);
                    excelApp.SetCellValue(i + 2, "BP", model.AccountGetDate);
                }
                excelApp.SetSheetAutoFit("Sheet1");
                excelApp.Save();
                excelApp.Close();
                ProcessLogProxy.SuccessMessage("Build Success");
            }
            catch (Exception ex)
            {
                ProcessLogProxy.Error(ex.Message);
                ProcessLogProxy.Error("Build fail");
            }

        }



        /// <summary>
        /// 得到个人信息
        /// </summary>
        /// <param name="businessModel"></param>
        /// <param name="yyyymm"></param>
        private void GetLRClaimData(ContractInfoBusiness businessModel, string yearMonthDay, ref int serialNumber)
        {
            for (int i = 0; i < businessModel.lstClaimSheetModel.Count; i++)
            {
                var model = businessModel.lstClaimSheetModel[i];

                LRClaimModel lRClaimModel = new LRClaimModel();
                //交易编码
                lRClaimModel.TransactionNo = CommFuns.GetTransactionNo(serialNumber, yearMonthDay);

                //保险机构代码
                lRClaimModel.CompanyCode = origanizationCode;

                //团体保单号
                lRClaimModel.GrpPolicyNo = "";

                //团体保单险种号码
                lRClaimModel.GrpProductNo = ""; //kong

                //个人保单号
                lRClaimModel.PolicyNo = model.PolicyNo;

                //个单保险险种号码  //TODO 个人 ProductNo
                var tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo) &&
                 e.ProductCode.Equal(model.Product)).FirstOrDefault();
                lRClaimModel.ProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.ProductNo;

                //保单团个性质代码
                lRClaimModel.GPFlag = "01";

                //主险保险险种号码  //TODO 个人
                lRClaimModel.MainProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductNo;

                //主附险性质代码  //TODO 个人
                lRClaimModel.MainProductFlag = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductFlag;

                //产品编码
                lRClaimModel.ProductCode = model.Product;

                //责任代码
                lRClaimModel.LiabilityCode = model.CauseOfClaim;

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(lRClaimModel.LiabilityCode)).FirstOrDefault();
                lRClaimModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;

                //给付责任代码 //AAATODO ClaimNo
                var tEMP_LLClaimDetail = businessModel.lstTEMP_LLClaimDetail.Where(A => A.ClmCaseNo == lRClaimModel.ClaimNo).FirstOrDefault();
                lRClaimModel.GetLiabilityCode = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.GetLiabilityCode;

                //给付责任名称
                lRClaimModel.GetLiabilityName = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.GetLiabilityName;

                //赔付责任类型代码
                lRClaimModel.BenefitType = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.BenefitType;

                //保险期限类型
                var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == lRClaimModel.ProductCode).FirstOrDefault();
                lRClaimModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                ///管理机构代码  //TODO 个人
                var tempLCCont = businessModel.lstTEMP_LCCont.Where(e => e.PolicyNo.Equals(lRClaimModel.PolicyNo)).FirstOrDefault();
                lRClaimModel.ManageCom = tempLCCont == null ? string.Empty : tempLCCont.ManageCom;

                //签单日期 //TODO 个人
                DateTime tempSignDate;
                string strSignDate = string.Empty;
                if (tempLCCont != null)
                {
                    bool convertResult = DateTime.TryParse(tempLCCont.SignDate, out tempSignDate);

                    if (convertResult)
                    {
                        strSignDate = tempSignDate.ToString("yyyy-MM-dd");
                    }
                }
                lRClaimModel.SignDate = strSignDate;

                //保险责任生效日期 //TODO 个人
                lRClaimModel.EffDate = tempLCProduct == null ? string.Empty : tempLCProduct.EffDate;

                //保单年度
                if (!string.IsNullOrEmpty(strSignDate))
                {
                    int currentYear = int.Parse(yearMonthDay.Substring(0, 4));
                    int signDateYear = int.Parse(strSignDate.Substring(0, 4));
                    lRClaimModel.PolYear = (currentYear - signDateYear).ToString();
                }
                else
                {
                    lRClaimModel.PolYear = "0";
                }

                //保险责任终止日期   //TODO 个人
                lRClaimModel.InvalidDate = tempLCProduct == null ? string.Empty : tempLCProduct.InvalidDate;

                //核保结论代码  //TODO 个人
                lRClaimModel.UWConclusion = tempLCProduct == null ? string.Empty : tempLCProduct.UWConclusion;

                //保单状态代码 //TODO 个人
                lRClaimModel.PolStatus = "01";

                //保单险种状态代码 //TODO 个人
                lRClaimModel.Status = "01";

                //基本保额
                var newTEMPLCProduct= businessModel.lstTEMP_LCProduct.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo) && e.ProductCode.Equal(lRClaimModel.ProductNo)).FirstOrDefault();
                lRClaimModel.BasicSumInsured = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.BasicSumInsured.Trim());

                //风险保额
                lRClaimModel.RiskAmnt = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.RiskAmnt.Trim());

                //保费
                lRClaimModel.Premium = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.Premium.Trim());

                //免赔类型代码
                lRClaimModel.DeductibleType = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.DeductibleType;

                //免赔额
                lRClaimModel.Deductible = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.Deductible;

                //赔付比例
                lRClaimModel.ClaimRatio = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.ClaimRatio;

                //保险账户价值 //TODO 个人
                var tempInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo)
                  && e.ProductNo.Equal(lRClaimModel.ProductNo)).FirstOrDefault();

                lRClaimModel.AccountValue = tempInsureAcc == null ?
                        string.Empty : Common.ConvertToStrToStrDecimal(tempInsureAcc.AccountValue);

                //临分标记
                lRClaimModel.FacultativeFlag = ConfigInformation.TextValue;

                //无名单标志
                lRClaimModel.AnonymousFlag = "0";

                //豁免险标志
                lRClaimModel.WaiverFlag = "否";

                //所需豁免剩余保费
                lRClaimModel.WaiverPrem = "0";

                //期末现金价值
                lRClaimModel.FinalCashValue = ConfigInformation.NumberValue;

                //被保人客户号
                lRClaimModel.InsuredNo = model.MembersCertificateNo;

                //被保人姓名 //TODO 个人
                var tempInsured = businessModel.lstTEMP_LCInsured.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo)
                   && e.InsuredNo.Equal(lRClaimModel.InsuredNo)).FirstOrDefault();

                lRClaimModel.InsuredName = tempInsured == null ? string.Empty : tempInsured.InsuredName;

                //被保人性别 //TODO 个人
                lRClaimModel.InsuredSex = tempInsured == null ? string.Empty : tempInsured.InsuredSex;

                //被保人证件类型 //TODO 个人
                lRClaimModel.InsuredCertType = tempInsured == null ? string.Empty : tempInsured.InsuredCertType;

                //被保人证件编码 //TODO 个人
                lRClaimModel.InsuredCertNo = tempInsured == null ? string.Empty : tempInsured.InsuredCertNo;

                //职业代码 //TODO 个人
                lRClaimModel.OccupationType = tempInsured == null ? string.Empty : tempInsured.OccupationType;

                //投保年龄 //TODO 个人
                lRClaimModel.AppntAge = tempInsured == null ? string.Empty : tempInsured.AppAge;

                //当前年龄
                lRClaimModel.PreAge = ConfigInformation.TextValue;

                //期末责任准备金
                lRClaimModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //职业加费金额 //TODO 个人
                lRClaimModel.ProfessionalFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.ProfessionalFee);

                //次标准体加费金额 //TODO 个人
                lRClaimModel.SubStandardFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.SubStandardFee);

                //EM加点 //TODO 个人
                lRClaimModel.EMRate = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.EMRate);

                //建工险标志
                lRClaimModel.ProjectFlag = ConfigInformation.TextValue;

                //投保总人数
                lRClaimModel.InsurePeoples = "1";

                //再保险公司名称 //TODO  个人
                lRClaimModel.ReinsurerName = model.CompanyName;

                //再保险公司代码
                lRClaimModel.ReinsurerCode = reinsurer.GetReinsurerInforByName(lRClaimModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
                  e.ReinsurerCode.Equal(lRClaimModel.ReinsurerCode)
                   && e.ProductCode.Equals(lRClaimModel.ProductCode) && e.LiabilityCode.Equals(lRClaimModel.LiabilityCode)).FirstOrDefault();

                lRClaimModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReInsuranceContNo;

                //分保方式
                lRClaimModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReinsurMode;

                //分出标记
                string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                      templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";

                lRClaimModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                //分保保额
                lRClaimModel.ReinsuranceAmnt = ConfigInformation.NumberValue;

                //自留额
                lRClaimModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       Common.ConvertToStrToStrDecimal(templstZaiBaoProductInfo.RetentionAmount);

                //分保比例
                lRClaimModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.QuotaSharePercentage;

                //赔案号
                var TEMP_LLClaimPolicyModel = businessModel.lstTEMP_LLClaimPolicy.Where(A => A.PolicyNo == lRClaimModel.PolicyNo).ToList();
                lRClaimModel.ClaimNo = TEMP_LLClaimPolicyModel.Count() == 0 ? "" : TEMP_LLClaimPolicyModel.First().ClaimNo;
                //出险日期
                var TEMP_LLClaimInfoModel = businessModel.lstTEMP_LLClaimInfo.Where(A => A.ClaimNo == lRClaimModel.ClaimNo).ToList();
                lRClaimModel.AccidentDate = TEMP_LLClaimInfoModel.Count() == 0 ? "" : TEMP_LLClaimInfoModel.First().AccidentDate;
                //结案日期
                lRClaimModel.ClmSettDate = TEMP_LLClaimInfoModel.Count() == 0 ? "" : TEMP_LLClaimInfoModel.First().ClmSettDate;
                //理赔结论代码
                lRClaimModel.PayStatusCode = TEMP_LLClaimPolicyModel.Count() == 0 ? "" : TEMP_LLClaimPolicyModel.First().PayStatusCode;
                //实际赔款金额
                lRClaimModel.ClaimMoney = model.PaidAmount;
                //摊回赔款金额
                lRClaimModel.BackClaimMoney = model.RecoveryAmount;
                //摊回日期
                lRClaimModel.BackDate = GetLastDayOfMonth(yearMonthDay);
                //货币代码
                lRClaimModel.Currency = "CNY";
                //分保计算日期
                lRClaimModel.ReComputationsDate = GetLastDayOfMonth(yearMonthDay);
                //账单归属日期
                lRClaimModel.AccountGetDate = GetLastDayOfMonth(yearMonthDay);

                serialNumber++;
                lRClaimModelList.Add(lRClaimModel);
            }
        }


        /// <summary>
        /// 得到团体信息
        /// </summary>
        /// <param name="businessModel"></param>
        /// <param name="yyyymm"></param>
        private void GetGroupLRClaimData(ContractInfoBusiness businessModel, string yyyymm, ref int serialNumber)
        {
            for (int i = 0; i < businessModel.lstRIClaimReportGroup.Count; i++)
            {
                var model = businessModel.lstRIClaimReportGroup[i];
                LRClaimModel lRClaimModel = new LRClaimModel();
                //交易编码
                lRClaimModel.TransactionNo = CommFuns.GetTransactionNo(serialNumber, yyyymm);//已赋值
                //保险机构代码
                lRClaimModel.CompanyCode = origanizationCode;
                //团体保单号
                lRClaimModel.GrpPolicyNo = model.Chdrnum;

                //团体保单险种号码
                lRClaimModel.GrpProductNo = model.ProdTyp;//16.4

                //个人保单号
                lRClaimModel.PolicyNo = model.PolicyNo;

                //个单保险险种号码//TODO 团体
                var tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo) &&
                 e.ProductCode.Equal(model.ProdTyp)).FirstOrDefault();
                lRClaimModel.ProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.ProductNo;

                //保单团个性质代码
                lRClaimModel.GPFlag = "02";

                //主险保险险种号码//TODO 团体
                lRClaimModel.MainProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductNo; ;//16.4

                //主附险性质代码//TODO 团体
                lRClaimModel.MainProductFlag = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductFlag;//16.4

                //产品编码
                lRClaimModel.ProductCode = model.ProductCode;

                //责任代码
                lRClaimModel.LiabilityCode = model.Claimcond;

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(model.Claimcond)).FirstOrDefault();
                lRClaimModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;//16.4

                //给付责任代码
                lRClaimModel.GetLiabilityCode = lRClaimModel.LiabilityCode;

                //给付责任名称
                lRClaimModel.GetLiabilityName = lRClaimModel.LiabilityName;

                //赔付责任类型代码//TODO AAAAA
                var tEMP_LLClaimDetail = businessModel.lstTEMP_LLClaimDetail.Where(A => A.ClmCaseNo == lRClaimModel.ClaimNo).FirstOrDefault();
                lRClaimModel.BenefitType = tEMP_LLClaimDetail == null ? "" : tEMP_LLClaimDetail.BenefitType;

                //保险期限类型
                var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == lRClaimModel.ProductCode).FirstOrDefault();
                lRClaimModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                ///管理机构代码  //TODO 团体
                var tempLCCont = businessModel.lstTEMP_LCCont.Where(a => a.GrpPolicyNo == lRClaimModel.GrpPolicyNo).FirstOrDefault();
                lRClaimModel.ManageCom = tempLCCont == null ? string.Empty : tempLCCont.ManageCom;

                //签单日期   //TODO 团体
                DateTime tempSignDate;
                string strSignDate = string.Empty;
                if (tempLCCont != null)
                {
                    bool convertResult = DateTime.TryParse(tempLCCont.SignDate, out tempSignDate);

                    if (convertResult)
                    {
                        strSignDate = tempSignDate.ToString("yyyy-MM-dd");
                    }
                }
                lRClaimModel.SignDate = strSignDate;

                //保险责任生效日期   //TODO 团体
                lRClaimModel.EffDate = tempLCProduct == null ? string.Empty : tempLCProduct.EffDate;

                //保单年度
                if (!string.IsNullOrEmpty(strSignDate))
                {
                    int currentYear = int.Parse(yyyymm.Substring(0, 4));
                    int signDateYear = int.Parse(strSignDate.Substring(0, 4));
                    lRClaimModel.PolYear = (currentYear - signDateYear).ToString();
                }
                else
                {
                    lRClaimModel.PolYear = "0";
                }

                //保险责任终止日期   //TODO 团体
                lRClaimModel.InvalidDate = tempLCProduct == null ? string.Empty : tempLCProduct.InvalidDate;

                //核保结论代码   //TODO 团体
                lRClaimModel.UWConclusion = tempLCProduct == null ? string.Empty : tempLCProduct.UWConclusion;

                //保单状态代码   //TODO 团体
                lRClaimModel.PolStatus = "01";

                //保单险种状态代码   //TODO 团体
                lRClaimModel.Status = "01";

                //基本保额
                var lCProduct_Group = businessModel.lstTEMP_LCProductGroup.Where(A => A.GrpPolicyNo == lRClaimModel.GrpPolicyNo && A.PolicyNo == lRClaimModel.PolicyNo && A.ProductNo == lRClaimModel.GrpProductNo).FirstOrDefault();
                lRClaimModel.BasicSumInsured = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.BasicSumInsured.Trim());
                                                                                                                                        
                //风险保额
                lRClaimModel.RiskAmnt = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.RiskAmnt.Trim());

                //保费
                lRClaimModel.Premium = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.Premium.Trim());

                //免赔类型代码
                var LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.ClmCaseNo == lRClaimModel.ClaimNo).FirstOrDefault();
                lRClaimModel.DeductibleType = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.DeductibleType;

                //免赔额
                lRClaimModel.Deductible = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.Deductible;

                //赔付比例
                lRClaimModel.ClaimRatio = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.ClaimRatio;

                //保险账户价值 //TODO 团体
                var tempInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo)
                   && e.ProductNo.Equal(lRClaimModel.ProductNo)).FirstOrDefault();

                lRClaimModel.AccountValue = tempInsureAcc == null ?
                        string.Empty : Common.ConvertToStrToStrDecimal(tempInsureAcc.AccountValue);

                //临分标记
                lRClaimModel.FacultativeFlag = ConfigInformation.TextValue;

                //无名单标志
                lRClaimModel.AnonymousFlag = "0";

                //豁免险标志
                lRClaimModel.WaiverFlag = "否";

                //所需豁免剩余保费
                lRClaimModel.WaiverPrem = "0";
                //期末现金价值
                lRClaimModel.FinalCashValue = ConfigInformation.NumberValue;

                //被保人客户号
                lRClaimModel.InsuredNo = model.Clntnum;

                //被保人姓名  //TODO 团体
                var tempInsured = businessModel.lstTEMP_LCInsured.Where(e => e.PolicyNo.Equal(lRClaimModel.PolicyNo)
                   && e.InsuredNo.Equal(lRClaimModel.InsuredNo)).FirstOrDefault();

                lRClaimModel.InsuredName = tempInsured == null ? string.Empty : tempInsured.InsuredName; ;

                //被保人性别  //TODO 团体
                lRClaimModel.InsuredSex = tempInsured == null ? string.Empty : tempInsured.InsuredSex;

                //被保人证件类型  //TODO 团体
                lRClaimModel.InsuredCertType = tempInsured == null ? string.Empty : tempInsured.InsuredCertType;

                //被保人证件编码  //TODO 团体
                lRClaimModel.InsuredCertNo = tempInsured == null ? string.Empty : tempInsured.InsuredCertNo;

                //职业代码  //TODO 团体
                lRClaimModel.OccupationType = tempInsured == null ? string.Empty : tempInsured.OccupationType;

                //投保年龄  //TODO 团体
                lRClaimModel.AppntAge = tempInsured == null ? string.Empty : tempInsured.AppAge;

                //当前年龄
                lRClaimModel.PreAge = ConfigInformation.TextValue;

                //期末责任准备金 
                lRClaimModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //职业加费金额  //TODO 团体
                lRClaimModel.ProfessionalFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.ProfessionalFee);

                //次标准体加费金额  //TODO 团体
                lRClaimModel.SubStandardFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.SubStandardFee);

                //EM加点  //TODO 团体
                lRClaimModel.EMRate = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.EMRate);

                //建工险标志
                lRClaimModel.ProjectFlag = ConfigInformation.TextValue;

                //投保总人数
                lRClaimModel.InsurePeoples = "1";

                //再保险公司名称 //TODO 团体
                lRClaimModel.ReinsurerName = Common.DefaultCommanyName;

                //再保险公司代码
                lRClaimModel.ReinsurerCode = reinsurer.GetReinsurerInforByName(lRClaimModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
                   e.ReinsurerCode.Equal(lRClaimModel.ReinsurerCode)
                    && e.ProductCode.Equals(lRClaimModel.ProductCode) && e.LiabilityCode.Equals(lRClaimModel.LiabilityCode)).FirstOrDefault();

                lRClaimModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReInsuranceContNo;

                //分保方式
                lRClaimModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReinsurMode;

                //分出标记
                string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                      templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";

                lRClaimModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                //分保保额
                lRClaimModel.ReinsuranceAmnt = ConfigInformation.NumberValue;

                //自留额
                lRClaimModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       Common.ConvertToStrToStrDecimal(templstZaiBaoProductInfo.RetentionAmount);

                //分保比例
                lRClaimModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.QuotaSharePercentage;

                //赔案号 
                LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.PolicyNo == lRClaimModel.PolicyNo&& A.GrpPolicyNo == lRClaimModel.GrpPolicyNo).FirstOrDefault();
                lRClaimModel.ClaimNo = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.ClmCaseNo;

                //出险日期
                var TEMP_LLClaimInfoModel = businessModel.lstTEMP_LLClaimInfo.Where(A => A.ClaimNo == lRClaimModel.ClaimNo).ToList();
                lRClaimModel.AccidentDate = model.Incurrdate;

                //结案日期
                lRClaimModel.ClmSettDate = TEMP_LLClaimInfoModel.Count() == 0 ? "" : TEMP_LLClaimInfoModel.First().ClmSettDate;

                //理赔结论代码
              //  lRClaimModel.PayStatusCode = TEMP_LLClaimPolicyModel.Count() == 0 ? "" : TEMP_LLClaimPolicyModel.First().PayStatusCode;

                //实际赔款金额
                lRClaimModel.ClaimMoney = model.Payclam;

                //摊回赔款金额
                lRClaimModel.BackClaimMoney = model.Clmrcusr01;

                //摊回日期
                lRClaimModel.BackDate = GetLastDayOfMonth(yyyymm);

                //货币代码
                lRClaimModel.Currency = "CNY";

                //分保计算日期
                lRClaimModel.ReComputationsDate = GetLastDayOfMonth(yyyymm);

                //账单归属日期
                lRClaimModel.AccountGetDate = GetLastDayOfMonth(yyyymm);

                serialNumber++;

                lRClaimModelList.Add(lRClaimModel);
            }

        }

        private string GetLastDayOfMonth(string yyyymm)
        {
            var date = yyyymm.Substring(0, 4) + "-" + yyyymm.Substring(4, 2) + "-01";
            return Convert.ToDateTime(date).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
        }
    }
}
