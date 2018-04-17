using HSBC.InsuranceDataAnalysis.ExcelCore;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LREdor
    {
        List<LREdorModel> lREdorModelList = new List<LREdorModel>();

        private string origanizationCode = "000131";
        public void WriteLREdorSheet(ContractInfoBusiness contractInfoBusiness, string OutPutFolderPath, string dateyyyymm)
        {
            IExcel excelApp = new ExcelCore.ExcelCore();
            int serialNumber = 0;
            try
            {

                ProcessLogProxy.Normal("Start building LREdor excel");
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LREdor + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LREdor);//创建模板
                GetGroupLREdorData(contractInfoBusiness, dateyyyymm);
                GetLREdorData(contractInfoBusiness, dateyyyymm);
                excelApp.OpenExcel(excelPath, false);
                for (int i = 0; i < lREdorModelList.Count; i++)
                {
                    serialNumber++;
                    var model = lREdorModelList[i];
                    excelApp.SetCellValue(i + 2, "A", CommFuns.GetTransactionNo(i + 1, dateyyyymm));
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
                    excelApp.SetCellValue(i + 2, "M", model.Classification);
                    excelApp.SetCellValue(i + 2, "N", model.TermType);
                    excelApp.SetCellValue(i + 2, "O", model.ManageCom);
                    excelApp.SetCellValue(i + 2, "P", model.SignDate);
                    excelApp.SetCellValue(i + 2, "Q", model.EffDate);
                    excelApp.SetCellValue(i + 2, "R", model.PolYear);
                    excelApp.SetCellValue(i + 2, "S", model.InvalidDate);
                    excelApp.SetCellValue(i + 2, "T", model.UWConclusion);
                    excelApp.SetCellValue(i + 2, "U", model.PolStatus);
                    excelApp.SetCellValue(i + 2, "V", model.Status);
                    excelApp.SetCellValue(i + 2, "W", model.BasicSumInsured);
                    excelApp.SetCellValue(i + 2, "X", model.RiskAmnt);
                    excelApp.SetCellValue(i + 2, "Y", model.Premium);
                    excelApp.SetCellValue(i + 2, "Z", model.AccountValue);
                    excelApp.SetCellValue(i + 2, "AA", model.FacultativeFlag);
                    excelApp.SetCellValue(i + 2, "AB", model.AnonymousFlag);
                    excelApp.SetCellValue(i + 2, "AC", model.WaiverFlag);
                    excelApp.SetCellValue(i + 2, "AD", model.WaiverPrem);
                    excelApp.SetCellValue(i + 2, "AE", model.FinalCashValue);
                    excelApp.SetCellValue(i + 2, "AF", model.FinalLiabilityReserve);
                    excelApp.SetCellValue(i + 2, "AG", model.InsuredNo);
                    excelApp.SetCellValue(i + 2, "AH", model.InsuredName);
                    excelApp.SetCellValue(i + 2, "AI", model.InsuredSex);
                    excelApp.SetCellValue(i + 2, "AJ", model.InsuredCertType);
                    excelApp.SetCellValue(i + 2, "AK", model.InsuredCertNo);
                    excelApp.SetCellValue(i + 2, "AL", model.OccupationType);
                    excelApp.SetCellValue(i + 2, "AM", model.AppntAge);
                    excelApp.SetCellValue(i + 2, "AN", model.PreAge);
                    excelApp.SetCellValue(i + 2, "AO", model.ProfessionalFee);
                    excelApp.SetCellValue(i + 2, "AP", model.SubStandardFee);
                    excelApp.SetCellValue(i + 2, "AQ", model.EMRate);
                    excelApp.SetCellValue(i + 2, "AR", model.ProjectFlag);
                    excelApp.SetCellValue(i + 2, "AS", model.InsurePeoples);
                    excelApp.SetCellValue(i + 2, "AT", model.EndorAcceptNo);
                    excelApp.SetCellValue(i + 2, "AU", model.EndorsementNo);
                    excelApp.SetCellValue(i + 2, "AV", model.EdorType);
                    excelApp.SetCellValue(i + 2, "AW", model.EdorValiDate);
                    excelApp.SetCellValue(i + 2, "AX", model.EdorConfDate);
                    excelApp.SetCellValue(i + 2, "AY", model.EdorMoney);
                    excelApp.SetCellValue(i + 2, "AZ", model.SaparateFlag);
                    excelApp.SetCellValue(i + 2, "BA", model.ReInsuranceContNo);
                    excelApp.SetCellValue(i + 2, "BB", model.ReinsurerCode);
                    excelApp.SetCellValue(i + 2, "BC", model.ReinsurerName);
                    excelApp.SetCellValue(i + 2, "BD", model.ReinsurMode);
                    excelApp.SetCellValue(i + 2, "BE", model.QuotaSharePercentage);
                    excelApp.SetCellValue(i + 2, "BF", model.PreInsuredAge);
                    excelApp.SetCellValue(i + 2, "BG", model.PreBasicSumInsured);
                    excelApp.SetCellValue(i + 2, "BH", model.PreRiskAmnt);
                    excelApp.SetCellValue(i + 2, "BI", model.PreReinsuranceAmnt);
                    excelApp.SetCellValue(i + 2, "BJ", model.PreRetentionAmount);
                    excelApp.SetCellValue(i + 2, "BK", model.PrePremium);
                    excelApp.SetCellValue(i + 2, "BL", model.PreAccountValue);
                    excelApp.SetCellValue(i + 2, "BM", model.PreWaiverPrem);
                    excelApp.SetCellValue(i + 2, "BN", model.ProjectAcreageChange);
                    excelApp.SetCellValue(i + 2, "BO", model.ProjectCostChange);
                    excelApp.SetCellValue(i + 2, "BP", model.ReinsuranceAmntChange);
                    excelApp.SetCellValue(i + 2, "BQ", model.RetentionAmount);
                    excelApp.SetCellValue(i + 2, "BR", model.ReinsurancePremiumChange);
                    excelApp.SetCellValue(i + 2, "BS", model.ReinsuranceCommssionChange);
                    excelApp.SetCellValue(i + 2, "BT", model.Currency);
                    excelApp.SetCellValue(i + 2, "BU", model.ReComputationsDate);
                    excelApp.SetCellValue(i + 2, "BV", model.AccountGetDate);

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
        private void GetLREdorData(ContractInfoBusiness businessModel, string yyyymm)
        {
            for (int i = 0; i < businessModel.lstInforceBusinessListing.Count; i++)
            {
                var tempModel = businessModel.lstInforceBusinessListing[i];
                LREdorModel currentModel = new LREdorModel();
                //交易编码
                currentModel.TransactionNo = "";

                //保险机构代码
                currentModel.CompanyCode = origanizationCode;

                //团体保单号
                currentModel.GrpPolicyNo = "";

                //团体保单险种号码
                currentModel.GrpProductNo = "";

                //个人保单号
                currentModel.PolicyNo = tempModel.PolicyNo;
                if (businessModel.lstTEMP_LCPolTransaction.Where(A => A.PolicyNo == tempModel.PolicyNo).Count() == 0) { continue; }

                //主附险性质代码
                currentModel.MainProductFlag = this.GetMainProductFlag(tempModel.ProductCode);

                //个单保险险种号码
                var tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e =>
                    e.PolicyNo.Equal(currentModel.PolicyNo) &&
                    e.ProductCode.Equal(tempModel.ProductCode) &&
                    e.MainProductFlag.Equals(currentModel.MainProductFlag)).FirstOrDefault();

                currentModel.ProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.ProductNo;

                //保单团个性质代码
                currentModel.GPFlag = "01";

                //主险保险险种号码
                tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e =>
                  e.PolicyNo.Equal(currentModel.PolicyNo) &&
                   e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                currentModel.MainProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.MainProductNo;

                //产品编码
                currentModel.ProductCode = tempModel.ProductCode;

                //责任代码
                currentModel.LiabilityCode = tempModel.Coverage1;

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(tempModel.Coverage1)).FirstOrDefault();
                currentModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;

                //责任分类代码
                currentModel.Classification = tempCategory == null ? string.Empty : tempCategory.LiabilityCategoryCode;

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

                //PolYear 所跑数据年份减去签单日期年份
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
                currentModel.InvalidDate = tempLCProduct == null ? string.Empty : tempLCProduct.InvalidDate;

                //核保结论代码
                currentModel.UWConclusion = tempLCProduct == null ? string.Empty : tempLCProduct.UWConclusion;

                //保单状态代码
                currentModel.PolStatus = "01";

                //保单险种状态代码
                currentModel.Status = "01";

                //基本保额
                currentModel.BasicSumInsured = Common.ConvertToStrToStrDecimal(tempModel.SumInsured);

                // 风险保额
                currentModel.RiskAmnt = Common.ConvertToStrToStrDecimal(tempModel.InitialSumatRisk);

                //保费
                currentModel.Premium = tempLCCont == null ? string.Empty :
                    Common.ConvertToStrToStrDecimal(tempLCCont.Premium);

                //保险账户价值
                var tempLstInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
&& e.ProductNo.Equal(currentModel.ProductNo));

                decimal tempAccountTotal = 0m;
                foreach (var temp in tempLstInsureAcc)
                {
                    if (!string.IsNullOrWhiteSpace(temp.AccountValue))
                    {
                        tempAccountTotal += decimal.Parse(temp.AccountValue.Trim());
                    }
                }

                string strTempAccountTotal = tempAccountTotal.ToString("0.00");

                currentModel.AccountValue = strTempAccountTotal;

                //临分标记
                if (tempModel.IsMrHealth)
                {
                    currentModel.FacultativeFlag = "0";
                }
                else
                {
                    currentModel.FacultativeFlag = tempModel.AutomaticorFacultative.Equals("A") ? "0" : "1";
                }

                //无名单标志
                currentModel.AnonymousFlag = "0";

                //豁免险标志
                currentModel.WaiverFlag = "0";

                //所需豁免剩余保费
                currentModel.WaiverPrem = "0";

                //期末现金价值
                currentModel.FinalCashValue = ConfigInformation.NumberValue;

                //期末责任准备金
                currentModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //被保人客户号
                currentModel.InsuredNo = tempModel.MemberCertificateNo;

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
                currentModel.PreAge = tempModel.Attainedage;

                //职业加费金额
                currentModel.ProfessionalFee = tempLCProduct == null ? ConfigInformation.NumberValue : Common.ConvertToStrToStrDecimal(tempLCProduct.ProfessionalFee);

                //次标准体加费金额
                currentModel.SubStandardFee = tempLCProduct == null ? ConfigInformation.NumberValue : Common.ConvertToStrToStrDecimal(tempLCProduct.SubStandardFee);

                //EM加点
                currentModel.EMRate = tempLCProduct == null ? ConfigInformation.NumberValue : Common.ConvertToStrToStrDecimal(tempLCProduct.EMRate);

                //建工险标志
                currentModel.ProjectFlag = ConfigInformation.TextValue;

                // 投保总人数
                currentModel.InsurePeoples = "1";
                //保全受理号码
                var listLCPolTransaction = businessModel.lstTEMP_LCPolTransaction.Where(A => A.PolicyNo == tempModel.PolicyNo).ToList();
                currentModel.EndorAcceptNo = listLCPolTransaction.Count == 0 ? "" : listLCPolTransaction.First().EndorAcceptNo;
                //保全批单号码
                currentModel.EndorsementNo = listLCPolTransaction.Count == 0 ? "" : listLCPolTransaction.First().EndorsementNo;
                //保全项目类型
                currentModel.EdorType = ConfigInformation.TextValue;
                //保全生效日期
                currentModel.EdorValiDate = ConfigInformation.TextValue;
                //保全确认日期
                currentModel.EdorConfDate = ConfigInformation.TextValue;
                //保全发生费用
                currentModel.EdorMoney = ConfigInformation.NumberValue;

                //再保险公司名称 
                currentModel.ReinsurerName = tempModel.CompanyName;

                //再保险公司代码
                currentModel.ReinsurerCode = new Reinsurer().GetReinsurerInforByName(currentModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
                e.ReinsurerCode.Equal(currentModel.ReinsurerCode)
                 && e.ProductCode.Equals(currentModel.ProductCode) && e.LiabilityCode.Equals(currentModel.LiabilityCode)).FirstOrDefault();

                currentModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.ReInsuranceContNo;

                // 分保方式
                currentModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.ReinsurMode;

                //分出标记
                string tempQuotaSharePercentage = (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                    templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";
                currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                //分保比例
                currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.QuotaSharePercentage;

                //变更前被保人投保年龄
                currentModel.PreInsuredAge = ConfigInformation.NumberValue;

                //变更前基本保额
                currentModel.PreBasicSumInsured = ConfigInformation.NumberValue;

                //变更前风险保额
                currentModel.PreRiskAmnt = ConfigInformation.NumberValue;

                //变更前分保保额
                currentModel.PreReinsuranceAmnt = ConfigInformation.NumberValue;

                //变更前自留额
                currentModel.PreRetentionAmount = ConfigInformation.NumberValue;

                //变更前保费
                currentModel.PrePremium = ConfigInformation.NumberValue;

                //变更前账户价值
                currentModel.PreAccountValue = ConfigInformation.NumberValue;

                //变更前所需豁免剩余保费
                currentModel.PreWaiverPrem = ConfigInformation.NumberValue;

                //建筑面积变化量
                currentModel.ProjectAcreageChange = ConfigInformation.NumberValue;

                //工程造价变化量
                currentModel.ProjectCostChange = ConfigInformation.NumberValue;

                //变更后分保保额
                currentModel.ReinsuranceAmntChange = ConfigInformation.NumberValue;

                //变更后自留额
                currentModel.RetentionAmount = ConfigInformation.NumberValue;

                //变更分保费
                currentModel.ReinsurancePremiumChange = ConfigInformation.NumberValue;

                //变更分保佣金
                currentModel.ReinsuranceCommssionChange = ConfigInformation.NumberValue;

                //货币代码
                currentModel.Currency = "CNY";

                //分保计算日期
                currentModel.ReComputationsDate = GetLastDayOfMonth(yyyymm);

                //账单归属日期
                currentModel.AccountGetDate = GetLastDayOfMonth(yyyymm);

                lREdorModelList.Add(currentModel);
            }


        }


        /// <summary>
        /// 得到团体信息
        /// </summary>
        /// <param name="businessModel"></param>
        /// <param name="yyyymm"></param>
        private void GetGroupLREdorData(ContractInfoBusiness businessModel, string yyyymm)
        {
            for (int i = 0; i < businessModel.lstPolicyAlternationReportGroup.Count; i++)
            {
                var tempModel = businessModel.lstPolicyAlternationReportGroup[i];
                LREdorModel currentModel = new LREdorModel();
                //交易编码
                currentModel.TransactionNo = "";
                //保险机构代码
                currentModel.CompanyCode = origanizationCode;

                //团体保单号
                currentModel.GrpPolicyNo = tempModel.ChdrNum;

                //团体保单险种号码
                //var tempLCGrpProduct = businessModel.lstTEMP_LCGrpProduct.Where(e => e.GrpPolicyNo == currentModel.GrpPolicyNo
                //  && e.ProductCode == tempModel.ProdTyp).FirstOrDefault();
                //currentModel.GrpProductNo = tempLCGrpProduct != null ? tempLCGrpProduct.GrpProductNo : string.Empty;
                currentModel.GrpProductNo = tempModel.ProdTyp;

                //个人保单号
                //var tempLCCont = businessModel.lstTEMP_LCCont.Where(a => a.GrpPolicyNo == currentModel.GrpPolicyNo).ToList().FirstOrDefault();
                string tempMbrno = string.IsNullOrWhiteSpace(tempModel.Mbrno) ? string.Empty : tempModel.Mbrno.Trim() + "00";
                currentModel.PolicyNo = tempMbrno.PadLeft(7, '0');

                // 个单保险险种号码
                currentModel.ProductNo = currentModel.GrpProductNo;

                //保单团个性质代码
                currentModel.GPFlag = "02";

                //主险保险险种号码
                var tempLCProductGroup = businessModel.lstTEMP_LCProductGroup.Where(e =>
                    e.GrpPolicyNo.Equals(currentModel.GrpPolicyNo) &&
                    e.PolicyNo.Equal(currentModel.PolicyNo) &&
                    e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                currentModel.MainProductNo = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductNo;

                //主附险性质代码
                currentModel.MainProductFlag = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductFlag;

                //产品编码
                string tempProductCode = string.IsNullOrWhiteSpace(tempModel.ProductCode) ? string.Empty : tempModel.ProductCode.Trim();
                currentModel.ProductCode = (tempProductCode.Equals("GIP")
                    || tempProductCode.Equals("GIP")) ? "GHB" : tempProductCode;
                //currentModel.ProductCode = tempModel.ProdTyp;

                //责任代码
                currentModel.LiabilityCode = Common.GetLiabilityCode(currentModel.ProductCode);

                //责任名称
                var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(currentModel.LiabilityCode)).FirstOrDefault();
                currentModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;

                //责任分类代码
                currentModel.Classification = tempCategory == null ? string.Empty : tempCategory.LiabilityCategoryCode;

                //保险期限类型
                var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == currentModel.ProductCode).FirstOrDefault();
                currentModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                //管理机构代码
                var tempLcGrpContGroup = businessModel.lstLCGrpContGroup.Where(e => e.GrpPolicyNo.Equal(currentModel.GrpPolicyNo)).FirstOrDefault();
                currentModel.ManageCom = tempLcGrpContGroup == null ? string.Empty : tempLcGrpContGroup.ManageCom;

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
                currentModel.EffDate = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.EffDate;

                //PolYear 所跑数据年份减去签单日期年份
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
                currentModel.BasicSumInsured = Common.ConvertToStrToStrDecimal(tempModel.SumSi);

                // 风险保额
                currentModel.RiskAmnt = ConfigInformation.TextValue;

                //保费
                currentModel.Premium = Common.ConvertToStrToStrDecimal(tempModel.Pprem);

                //保险账户价值
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

                //期末责任准备金
                currentModel.FinalLiabilityReserve = ConfigInformation.NumberValue;

                //被保人客户号
                currentModel.InsuredNo = tempModel.Clntnum;

                //被保人姓名
                var tempInsuredGroup = businessModel.lst_LCInsuredGroup.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
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
                currentModel.PreAge = ConfigInformation.NumberValue;

                //职业加费金额
                currentModel.ProfessionalFee = "0";

                //次标准体加费金额
                currentModel.SubStandardFee = "0";

                //EM加点
                if (tempLCProductGroup != null)
                {
                    currentModel.EMRate = Common.ConvertToStrToStrDecimal(tempLCProductGroup.EMRate);
                }
                else
                {
                    currentModel.EMRate = string.Empty;
                }

                //建工险标志
                currentModel.ProjectFlag = ConfigInformation.TextValue;

                //投保总人数
                currentModel.InsurePeoples = "1";

                //保全受理号码
                var listLCPolTransaction = businessModel.lstTEMP_LCPolTransaction.Where(A => A.PolicyNo == currentModel.PolicyNo).ToList();
                currentModel.EndorAcceptNo = listLCPolTransaction.Count == 0 ? "" : listLCPolTransaction.First().EndorAcceptNo;

                //保全批单号码
                currentModel.EndorsementNo = listLCPolTransaction.Count == 0 ? "" : listLCPolTransaction.First().EndorsementNo;

                //保全项目类型
                currentModel.EdorType = ConfigInformation.TextValue;
                //保全生效日期
                currentModel.EdorValiDate = ConfigInformation.TextValue;
                //保全确认日期
                currentModel.EdorConfDate = ConfigInformation.TextValue;
                //保全发生费用
                currentModel.EdorMoney = ConfigInformation.NumberValue;

                //分出标记
                currentModel.SaparateFlag = "";

                //分出标记
                var templstZaiBaoProductInfo = businessModel.lstZaiBaoProductInfo.Where(e =>
               e.ReinsurerCode.Equal(currentModel.ReinsurerCode)
                && e.ProductCode.Equals(currentModel.ProductCode) && e.LiabilityCode.Equals(currentModel.LiabilityCode)).FirstOrDefault();

                string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                    templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";
                currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;


                //再保险公司名称
                currentModel.ReinsurerName = Common.DefaultCommanyName;/////////////////////////////////////

                //再保险公司代码
                Reinsurer Reinsurer = new Reinsurer();
                currentModel.ReinsurerCode = Reinsurer.GetReinsurerInforByName(currentModel.ReinsurerName).ReinsurerCode;

                //再保险合同号码
                currentModel.ReInsuranceContNo = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.ReInsuranceContNo;

                // 分保方式
                currentModel.ReinsurMode = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.ReinsurMode;

                //分保比例
                currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                    templstZaiBaoProductInfo.QuotaSharePercentage;

                //变更前被保人投保年龄
                currentModel.PreInsuredAge = ConfigInformation.NumberValue;

                //变更前基本保额
                currentModel.PreBasicSumInsured = ConfigInformation.NumberValue;

                //变更前风险保额
                currentModel.PreRiskAmnt = ConfigInformation.NumberValue;

                //变更前分保保额
                currentModel.PreReinsuranceAmnt = ConfigInformation.NumberValue;

                //变更前自留额
                currentModel.PreRetentionAmount = ConfigInformation.NumberValue;

                //变更前保费
                currentModel.PrePremium = ConfigInformation.NumberValue;

                //变更前账户价值
                currentModel.PreAccountValue = ConfigInformation.NumberValue;

                //变更前所需豁免剩余保费
                currentModel.PreWaiverPrem = ConfigInformation.NumberValue;

                //建筑面积变化量
                currentModel.ProjectAcreageChange = ConfigInformation.NumberValue;

                //工程造价变化量
                currentModel.ProjectCostChange = ConfigInformation.NumberValue;

                //变更后分保保额
                currentModel.ReinsuranceAmntChange = ConfigInformation.NumberValue;

                //变更后自留额
                currentModel.RetentionAmount = ConfigInformation.NumberValue;

                //变更分保费
                currentModel.ReinsurancePremiumChange = ConfigInformation.NumberValue;

                //变更分保佣金
                currentModel.ReinsuranceCommssionChange = ConfigInformation.NumberValue;

                //货币代码
                currentModel.Currency = "CNY";

                //分保计算日期
                currentModel.ReComputationsDate = GetLastDayOfMonth(yyyymm);

                //账单归属日期
                currentModel.AccountGetDate = GetLastDayOfMonth(yyyymm);

                lREdorModelList.Add(currentModel);
            }

        }

        private string GetLastDayOfMonth(string yyyymm)
        {
            var date = yyyymm.Substring(0, 4) + "-" + yyyymm.Substring(4, 2) + "-01";
            return Convert.ToDateTime(date).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
        }

        private string GetMainProductFlag(string productCode)
        {
            string mainProductFlag = string.Empty;
            productCode = string.IsNullOrWhiteSpace(productCode) ? string.Empty : productCode.Trim().ToUpper();
            switch (productCode)
            {
                case "HC2":
                    mainProductFlag = "2";
                    break;
                case "MI1":
                    mainProductFlag = "2";
                    break;
                case "MI2":
                    mainProductFlag = "2";
                    break;
                case "MI3":
                    mainProductFlag = "2";
                    break;
                case "MM1":
                    mainProductFlag = "2";
                    break;
                default:
                    mainProductFlag = "1";
                    break;
            }
            return mainProductFlag;
        }
    }
}
