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
                var tempModel = businessModel.lstClaimSheetModel[i];

                LRClaimModel currentModel = new LRClaimModel();
                //交易编码
                currentModel.TransactionNo = CommFuns.GetTransactionNo(serialNumber, yearMonthDay);

                //保险机构代码
                currentModel.CompanyCode = origanizationCode;

                //团体保单号
                currentModel.GrpPolicyNo = "";

                //团体保单险种号码
                currentModel.GrpProductNo = ""; //kong

                //个人保单号
                currentModel.PolicyNo = tempModel.PolicyNo.PadLeft(8,'0');

                //主附险性质代码
                currentModel.MainProductFlag = CommFuns.GetMainProductFlag(tempModel.Product);

                // 个单保险险种号码
                var tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e =>
                  e.PolicyNo.Equal(currentModel.PolicyNo) &&
                   e.ProductCode.Equal(tempModel.Product) &&
                   e.MainProductFlag.Equals(currentModel.MainProductFlag)).FirstOrDefault();
                currentModel.ProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.ProductNo;

                //保单团个性质代码
                currentModel.GPFlag = "01";

                // 主险保险险种号码
                tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo) && e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();
                currentModel.MainProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductNo;

               

                //产品编码
                currentModel.ProductCode = tempModel.Product;

                //责任代码
                currentModel.LiabilityCode = tempModel.CauseOfClaim;

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(currentModel.LiabilityCode)).FirstOrDefault();
                currentModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;

                //赔案号
                var TEMPLLClaimDetail = businessModel.lstTEMP_LLClaimDetail.Where(A => A.PolicyNo == currentModel.PolicyNo && A.ProductNo == currentModel.ProductNo).FirstOrDefault();
                currentModel.ClaimNo = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.ClmCaseNo;

                //给付责任代码 
                TEMPLLClaimDetail = businessModel.lstTEMP_LLClaimDetail.Where(A => A.ClmCaseNo == currentModel.ClaimNo).FirstOrDefault();
                currentModel.GetLiabilityCode = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.GetLiabilityCode;

                //给付责任名称
                currentModel.GetLiabilityName = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.GetLiabilityName;

                //赔付责任类型代码
                currentModel.BenefitType = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.BenefitType;

                //保险期限类型
                var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == currentModel.ProductCode).FirstOrDefault();
                currentModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                //管理机构代码
                var tempLCCont = businessModel.lstTEMP_LCCont.Where(e => e.PolicyNo.Equals(tempModel.PolicyNo)).FirstOrDefault();
                currentModel.ManageCom = tempLCCont == null ? string.Empty : tempLCCont.ManageCom;

                //签单日期
                DateTime tempSignDate;
                string strSignDate = string.Empty;
                if (tempLCCont != null)
                {
                    bool convertResult = DateTime.TryParse(tempLCCont.SignDate, out tempSignDate);

                    if (convertResult)
                    {
                        strSignDate = tempSignDate.ToString("yyyy/MM/dd");
                    }
                }
                currentModel.SignDate = strSignDate;

                //保险责任生效日期
                currentModel.EffDate = tempLCProduct == null ? string.Empty : tempLCProduct.EffDate;



                //保单年度
                if (!string.IsNullOrEmpty(strSignDate))
                {
                    int currentYear = int.Parse(yearMonthDay.Substring(0, 4));
                    int signDateYear = int.Parse(strSignDate.Substring(0, 4));
                    currentModel.PolYear = (currentYear - signDateYear).ToString();
                }
                else
                {
                    currentModel.PolYear = "0";
                }

                //保险责任终止日期
                currentModel.InvalidDate = tempLCProduct == null ? string.Empty : tempLCProduct.InvalidDate;

                //核保结论代码
                currentModel.UWConclusion = tempLCProduct == null ? string.Empty : tempLCProduct.UWConclusion;

                //保单状态代码
                currentModel.PolStatus = "01";

                //保单险种状态代码
                currentModel.Status = "01";

                //基本保额
                var newTEMPLCProduct = businessModel.lstTEMP_LCProduct.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo) && e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();
                currentModel.BasicSumInsured = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.BasicSumInsured.Trim());

                //风险保额
                currentModel.RiskAmnt = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.RiskAmnt.Trim());

                //保费
                currentModel.Premium = newTEMPLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(newTEMPLCProduct.Premium.Trim());

                //免赔类型代码
                currentModel.DeductibleType = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.DeductibleType;

                //免赔额
                currentModel.Deductible = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.Deductible;

                //赔付比例
                currentModel.ClaimRatio = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.ClaimRatio;

                //保险账户价值 //TODO 个人
                var tempInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
                  && e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                currentModel.AccountValue = tempInsureAcc == null ?
                        string.Empty : Common.ConvertToStrToStrDecimal(tempInsureAcc.AccountValue);

                //临分标记
                currentModel.FacultativeFlag = ConfigInformation.TextValue;

                //无名单标志
                currentModel.AnonymousFlag = "0";

                //豁免险标志
                currentModel.WaiverFlag = "0";

                //所需豁免剩余保费
                currentModel.WaiverPrem = "0";

                //期末现金价值
                currentModel.FinalCashValue = ConfigInformation.NumberValue;

                //被保人客户号
                currentModel.InsuredNo = tempModel.MembersCertificateNo.PadLeft(8, '0'); ;

                //被保人姓名
                var tempInsured = businessModel.lstTEMP_LCInsured.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
                && e.InsuredNo.Equal(currentModel.InsuredNo)).FirstOrDefault();
                currentModel.InsuredName = tempInsured == null ? string.Empty : tempInsured.InsuredName;

                //被保人性别
                currentModel.InsuredSex = tempInsured == null ? string.Empty : tempInsured.InsuredSex;

                //被保人证件类型
                currentModel.InsuredCertType = tempInsured == null ? string.Empty : tempInsured.InsuredCertType;

                //被保人证件编码
                currentModel.InsuredCertNo = tempInsured == null ? string.Empty : tempInsured.InsuredCertNo;

                //职业代码
                currentModel.OccupationType = tempInsured == null ? string.Empty : tempInsured.OccupationType;

                //投保年龄
                currentModel.AppntAge = tempInsured == null ? string.Empty : tempInsured.AppAge;

                //当前年龄
                currentModel.PreAge = ConfigInformation.TextValue;

                //期末责任准备金
                currentModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //职业加费金额
                currentModel.ProfessionalFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.ProfessionalFee);

                //次标准体加费金额
                currentModel.SubStandardFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.SubStandardFee);

                //EM加点
                currentModel.EMRate = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.EMRate);

                //建工险标志
                currentModel.ProjectFlag = ConfigInformation.TextValue;

                //投保总人数
                currentModel.InsurePeoples = "1";

                //再保险公司名称
                currentModel.ReinsurerName = tempModel.CompanyName;

                //再保险公司代码
                currentModel.ReinsurerCode = reinsurer.GetReinsurerInforByName(currentModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
                  e.ReinsurerCode.Equal(currentModel.ReinsurerCode)
                   && e.ProductCode.Equals(currentModel.ProductCode) && e.LiabilityCode.Equals(currentModel.LiabilityCode)).FirstOrDefault();

                currentModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReInsuranceContNo;

                //分保方式
                currentModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReinsurMode;

                //分出标记
                string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                      templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";

                currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                //分保保额
                currentModel.ReinsuranceAmnt = ConfigInformation.NumberValue;

                //自留额
                currentModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       Common.ConvertToStrToStrDecimal(templstZaiBaoProductInfo.RetentionAmount);

                //分保比例
                currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : templstZaiBaoProductInfo.QuotaSharePercentage;

                //赔案号 前面已经赋值 lRClaimModel.ClaimNo 

                //出险日期
                var TEMP_LLClaimInfoModel = businessModel.lstTEMP_LLClaimInfo.Where(A => A.ClaimNo == currentModel.ClaimNo).FirstOrDefault();
                currentModel.AccidentDate = TEMP_LLClaimInfoModel == null ? "" : TEMP_LLClaimInfoModel.AccidentDate;

                //结案日期
                currentModel.ClmSettDate = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.ClmSettDate;

                //理赔结论代码
                currentModel.PayStatusCode = TEMPLLClaimDetail == null ? "" : TEMPLLClaimDetail.PayStatusCode;

                //实际赔款金额
                currentModel.ClaimMoney = tempModel.PaidAmount;

                //摊回赔款金额
                currentModel.BackClaimMoney = tempModel.RecoveryAmount;

                //摊回日期
                currentModel.BackDate = GetLastDayOfMonth(yearMonthDay);

                //货币代码
                currentModel.Currency = "CNY";

                //分保计算日期
                currentModel.ReComputationsDate = GetLastDayOfMonth(yearMonthDay);

                //账单归属日期
                currentModel.AccountGetDate = GetLastDayOfMonth(yearMonthDay);

                serialNumber++;
                lRClaimModelList.Add(currentModel);
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
                var tempModel = businessModel.lstRIClaimReportGroup[i];
                LRClaimModel currentModel = new LRClaimModel();
                //交易编码
                currentModel.TransactionNo = CommFuns.GetTransactionNo(serialNumber, yyyymm);//已赋值
                //保险机构代码
                currentModel.CompanyCode = origanizationCode;
                //团体保单号
                currentModel.GrpPolicyNo = tempModel.Chdrnum;

                //团体保单险种号码
                currentModel.GrpProductNo = tempModel.ProdTyp;//16.4

                //个人保单号
                currentModel.PolicyNo = tempModel.PolicyNo.PadLeft(7,'0');


                //个单保险险种号码
                currentModel.ProductNo = tempModel.ProdTyp;


                //保单团个性质代码
                currentModel.GPFlag = "02";

                // 主险保险险种号码
                var tempLCProductGroup = businessModel.lstTEMP_LCProductGroup.Where(e =>
                e.GrpPolicyNo.Equals(currentModel.GrpPolicyNo) &&
                e.PolicyNo.Substring(1,7).Equal(currentModel.PolicyNo) &&
                e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();
                currentModel.MainProductNo = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductNo;

                //主附险性质代码
                currentModel.MainProductFlag = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductFlag;

                //产品编码
                currentModel.ProductCode = tempModel.ProductCode;

                //责任代码
                currentModel.LiabilityCode = tempModel.Claimcond;

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(tempModel.Claimcond)).FirstOrDefault();
                currentModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;//16.4

                //给付责任代码
                currentModel.GetLiabilityCode = currentModel.LiabilityCode;

                //给付责任名称
                currentModel.GetLiabilityName = currentModel.LiabilityName;

                //赔案号 
                var LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.PolicyNo.PadLeft(7,'0') == currentModel.PolicyNo && A.GrpPolicyNo == currentModel.GrpPolicyNo).FirstOrDefault();
                currentModel.ClaimNo = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.ClmCaseNo;

                //赔付责任类型代码
                LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.ClmCaseNo == currentModel.ClaimNo).FirstOrDefault();
                currentModel.BenefitType = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.BenefitType;

                //保险期限类型
                var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == currentModel.ProductCode).FirstOrDefault();
                currentModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                //管理机构代码
                var tempLcGrpContGroup = businessModel.lstLCGrpContGroup.Where(e => e.GrpPolicyNo.Equal(currentModel.GrpPolicyNo)).FirstOrDefault();
                currentModel.ManageCom = tempLcGrpContGroup == null ? ConfigInformation.TextValue : tempLcGrpContGroup.ManageCom;

                //签单日期
                DateTime tempSignDate;
                string strSignDate = string.Empty;
                if (tempLcGrpContGroup != null)
                {
                    bool convertResult = DateTime.TryParse(tempLcGrpContGroup.SignDate, out tempSignDate);

                    if (convertResult)
                    {
                        strSignDate = tempSignDate.ToString("yyyy-MM-dd");
                    }
                }
                currentModel.SignDate = strSignDate;

                //保险责任生效日期
                currentModel.EffDate = tempLcGrpContGroup == null ? string.Empty : tempLcGrpContGroup.EffDate;

                //保单年度
                if (!string.IsNullOrEmpty(strSignDate))
                {
                    int currentYear = int.Parse(yyyymm.Substring(0, 4));
                    int signDateYear = int.Parse(strSignDate.Substring(0, 4));
                    currentModel.PolYear = (currentYear - signDateYear).ToString();
                }
                else
                {
                    currentModel.PolYear = "0";
                }


                //保险责任终止日期
                currentModel.InvalidDate = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.InvalidDate;

                //核保结论代码
                currentModel.UWConclusion = "10";

                //保单状态代码
                currentModel.PolStatus = "03";

                //保单险种状态代码
                currentModel.Status = "03";

                //基本保额
                var lCProduct_Group = businessModel.lstTEMP_LCProductGroup.Where(A => A.GrpPolicyNo== currentModel.GrpPolicyNo && A.PolicyNo.Substring(1, 7) == currentModel.PolicyNo && A.ProductNo == currentModel.GrpProductNo).FirstOrDefault();
                currentModel.BasicSumInsured = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.BasicSumInsured.Trim());

                //风险保额
                currentModel.RiskAmnt = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.RiskAmnt.Trim());

                //保费
                currentModel.Premium = lCProduct_Group == null ? string.Empty : Common.ConvertToStrToStrDecimal(lCProduct_Group.Premium.Trim());


                //免赔类型代码
                LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.ClmCaseNo == currentModel.ClaimNo).FirstOrDefault();
                currentModel.DeductibleType = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.DeductibleType;

                //免赔额
                currentModel.Deductible = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.Deductible;

                //赔付比例
                currentModel.ClaimRatio = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.ClaimRatio;

                //保险账户价值 //TODO 团体
                var tempInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
                   && e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                currentModel.AccountValue = "0";

                //临分标记
                currentModel.FacultativeFlag = ConfigInformation.TextValue;

                //无名单标志
                currentModel.AnonymousFlag = "0";

                //豁免险标志
                currentModel.WaiverFlag = "0";

                //所需豁免剩余保费
                currentModel.WaiverPrem = "0";
                //期末现金价值
                currentModel.FinalCashValue = ConfigInformation.NumberValue;

                //被保人客户号
                currentModel.InsuredNo = tempModel.Clntnum.PadLeft(8,'0');

                //被保人姓名
                var tempInsuredGroup = businessModel.lst_LCInsuredGroup.Where(e => e.PolicyNo.Substring(1,7).Equal(currentModel.PolicyNo)
                && e.GrpPolicyNo.Equal(currentModel.GrpPolicyNo)).FirstOrDefault();
                currentModel.InsuredName = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.InsuredName;

                //被保人性别
                currentModel.InsuredSex = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.InsuredSex;

                //被保人证件类型
                currentModel.InsuredCertType = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.InsuredCertType;

                //被保人证件编码
                currentModel.InsuredCertNo = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.InsuredCertNo;

                //职业代码
                currentModel.OccupationType = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.OccupationType;

                //投保年龄
                currentModel.AppntAge = tempInsuredGroup == null ? string.Empty : tempInsuredGroup.AppAge;

                //当前年龄
                currentModel.PreAge = ConfigInformation.TextValue;

                //期末责任准备金 
                currentModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //职业加费金额
                currentModel.ProfessionalFee = "0";

                //次标准体加费金额
                currentModel.SubStandardFee = "0";

                //EM加点
                currentModel.EMRate = tempLCProductGroup == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProductGroup.EMRate);

                //建工险标志
                currentModel.ProjectFlag = ConfigInformation.TextValue;

                //投保总人数
                currentModel.InsurePeoples = "1";

                //再保险公司名称 
                currentModel.ReinsurerName = Common.DefaultCommanyName;

                //再保险公司代码
                currentModel.ReinsurerCode = reinsurer.GetReinsurerInforByName(currentModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
                   e.ReinsurerCode.Equal(currentModel.ReinsurerCode)
                    && e.ProductCode.Equals(currentModel.ProductCode) && e.LiabilityCode.Equals(currentModel.LiabilityCode)).FirstOrDefault();

                currentModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReInsuranceContNo;

                //分保方式
                currentModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.ReinsurMode;

                //分出标记
                string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                      templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";

                currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                //分保保额
                currentModel.ReinsuranceAmnt = ConfigInformation.NumberValue;

                //自留额
                currentModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       Common.ConvertToStrToStrDecimal(templstZaiBaoProductInfo.RetentionAmount);

                //分保比例
                currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.QuotaSharePercentage;

                //赔案号  lRClaimModel.ClaimNo  前面赋值

                //出险日期
                currentModel.AccidentDate =tempModel.Incurrdate.Substring(0,4)+"/"+ tempModel.Incurrdate.Substring(4, 2)+"/"+ tempModel.Incurrdate.Substring(6, 2);

                //结案日期
                LLClaimDetailGroup = businessModel.lstLLClaimDetailGroup.Where(A => A.ClmCaseNo == currentModel.ClaimNo).FirstOrDefault();
                currentModel.ClmSettDate = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.ClmSettDate;

                //理赔结论代码
                currentModel.PayStatusCode = LLClaimDetailGroup == null ? "" : LLClaimDetailGroup.PayStatusCode;

                //实际赔款金额
                currentModel.ClaimMoney = tempModel.Payclam;

                //摊回赔款金额
                currentModel.BackClaimMoney = tempModel.Clmrcusr01;

                //摊回日期
                currentModel.BackDate = GetLastDayOfMonth(yyyymm);

                //货币代码
                currentModel.Currency = "CNY";

                //分保计算日期
                currentModel.ReComputationsDate = GetLastDayOfMonth(yyyymm);

                //账单归属日期
                currentModel.AccountGetDate = GetLastDayOfMonth(yyyymm);

                serialNumber++;

                lRClaimModelList.Add(currentModel);
            }

        }

        private string GetLastDayOfMonth(string yyyymm)
        {
            var date = yyyymm.Substring(0, 4) + "-" + yyyymm.Substring(4, 2) + "-01";
            return Convert.ToDateTime(date).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
        }
    }
}
