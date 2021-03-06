﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using HSBC.InsuranceDataAnalysis.ExcelCore;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LRCont
    {
        List<LRContModel> lRContModelList = new List<LRContModel>();
        Reinsurer reinsurer = new Reinsurer();

        public void WriteLRContSheet(ContractInfoBusiness contractInfoBusiness,
            string OutPutFolderPath, string dateyyyymm)
        {
            IExcel excelApp = new ExcelCore.ExcelCore();
            try
            {
                ProcessLogProxy.Normal("Start building LRCont excel");
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LRCont + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LRCont);//创建模板
                int serialNumber = 1;
                SetGroupDataToModel(contractInfoBusiness, dateyyyymm, ref serialNumber);
                SetIndividualDataToModel(contractInfoBusiness, dateyyyymm, ref serialNumber);
                excelApp.OpenExcel(excelPath, false);
                for (int i = 0; i < lRContModelList.Count; i++)
                {
                    var model = lRContModelList[i];

                    excelApp.SetCellValue(i + 2, "A", model.TransactionNo);
                    excelApp.SetCellValue(i + 2, "B", model.CompanyCode);
                    excelApp.SetCellValue(i + 2, "C", model.GrpPolicyNo);

                    if (model.IsGroup && string.IsNullOrWhiteSpace(model.GrpProductNo))
                    {
                        Color errorBackGroudColor = Color.Yellow;
                        excelApp.SetCellBackgroundColor(i + 2, 4, errorBackGroudColor);
                    }
                    else
                    {
                        excelApp.SetCellValue(i + 2, "D", model.GrpProductNo);
                    }

                    excelApp.SetCellValue(i + 2, "E", model.PolicyNo);
                    excelApp.SetCellValue(i + 2, "F", model.ProductNo);
                    excelApp.SetCellValue(i + 2, "G", model.GPFlag);
                    excelApp.SetCellValue(i + 2, "H", model.MainProductNo);
                    excelApp.SetCellValue(i + 2, "I", model.MainProductFlag);
                    excelApp.SetCellValue(i + 2, "J", model.ProductCode);
                    excelApp.SetCellValue(i + 2, "K", model.LiabilityCode);
                    excelApp.SetCellValue(i + 2, "L", model.LiabilityName);
                    excelApp.SetCellValue(i + 2, "M", model.Classification);
                    excelApp.SetCellValue(i + 2, "N", model.EventType);
                    excelApp.SetCellValue(i + 2, "O", model.RenewalTimes);
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
                    excelApp.SetCellValue(i + 2, "AB", model.AccountValue);
                    excelApp.SetCellValue(i + 2, "AC", model.FacultativeFlag);
                    excelApp.SetCellValue(i + 2, "AD", model.AnonymousFlag);
                    excelApp.SetCellValue(i + 2, "AE", model.WaiverFlag);
                    excelApp.SetCellValue(i + 2, "AF", model.WaiverPrem);
                    excelApp.SetCellValue(i + 2, "AG", model.FinalCashValue);
                    excelApp.SetCellValue(i + 2, "AH", model.FinalLiabilityReserve);
                    excelApp.SetCellValue(i + 2, "AI", model.InsuredNo);
                    excelApp.SetCellValue(i + 2, "AJ", model.InsuredName);
                    excelApp.SetCellValue(i + 2, "AK", model.InsuredSex);
                    excelApp.SetCellValue(i + 2, "AL", model.InsuredCertType);
                    excelApp.SetCellValue(i + 2, "AM", model.InsuredCertNo);
                    excelApp.SetCellValue(i + 2, "AN", model.OccupationType);
                    excelApp.SetCellValue(i + 2, "AO", model.AppntAge);
                    excelApp.SetCellValue(i + 2, "AP", model.PreAge);
                    excelApp.SetCellValue(i + 2, "AQ", model.ProfessionalFee);
                    excelApp.SetCellValue(i + 2, "AR", model.SubStandardFee);
                    excelApp.SetCellValue(i + 2, "AS", model.EMRate);
                    excelApp.SetCellValue(i + 2, "AT", model.ProjectFlag);
                    excelApp.SetCellValue(i + 2, "AU", model.InsurePeoples);
                    excelApp.SetCellValue(i + 2, "AV", model.SaparateFlag);
                    excelApp.SetCellValue(i + 2, "AW", model.ReInsuranceContNo);
                    excelApp.SetCellValue(i + 2, "AX", model.ReinsurerCode);
                    excelApp.SetCellValue(i + 2, "AY", model.ReinsurerName);
                    excelApp.SetCellValue(i + 2, "AZ", model.ReinsurMode);
                    excelApp.SetCellValue(i + 2, "BA", model.ReinsuranceAmnt);
                    excelApp.SetCellValue(i + 2, "BB", model.RetentionAmount);
                    excelApp.SetCellValue(i + 2, "BC", model.Currency);
                    excelApp.SetCellValue(i + 2, "BD", model.QuotaSharePercentage);
                    excelApp.SetCellValue(i + 2, "BE", model.ReinsurancePremium);
                    excelApp.SetCellValue(i + 2, "BF", model.ReinsuranceCommssion);
                    excelApp.SetCellValue(i + 2, "BG", model.ReComputationsDate);
                    excelApp.SetCellValue(i + 2, "BH", model.AccountGetDate);
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


        private void SetGroupDataToModel(ContractInfoBusiness businessModel, string yearMonthDay,
            ref int serialNumber)
        {
            if (businessModel.lstRIMonthlyReportGroup.Count > 0)
            {
                foreach (var tempModel in businessModel.lstRIMonthlyReportGroup)
                {
                    LRContModel currentModel = new LRContModel();

                    currentModel.IsGroup = true;
                    currentModel.TransactionNo = CommFuns.GetTransactionNo4(serialNumber, yearMonthDay);
                    currentModel.CompanyCode = CommFuns.OriganizationCode;
                    currentModel.GrpPolicyNo = tempModel.ChdrNumber;

                    //var tempLCGrpProduct = businessModel.lstTEMP_LCGrpProduct.Where(e => e.GrpPolicyNo == currentModel.GrpPolicyNo
                    //&& e.ProductCode == tempModel.Prodtyp).FirstOrDefault();

                    //若为团单则按右侧位置取数,inputFile: yyyymm\group\"RI Monthly report-GROUP.csv" -> Col F
                    //if (tempLCGrpProduct != null)
                    //{
                    //    currentModel.GrpProductNo = tempLCGrpProduct.GrpProductNo;
                    //}
                    //else
                    //{
                    //    currentModel.GrpProductNo = string.Empty;
                    //}

                    currentModel.GrpProductNo = tempModel.Prodtyp;

                    //个人保单号
                    //var tempLCCont = businessModel.lstTEMP_LCCont.Where(e => e.PolicyNo.Equals(currentModel.GrpPolicyNo)).FirstOrDefault();
                    //currentModel.PolicyNo = tempLCCont == null ? string.Empty : (tempLCCont.PolicyNo + "00").PadLeft(7, '0');

                    currentModel.PolicyNo =(tempModel.Mbrno + "00").PadLeft(7, '0');

                    // 个单保险险种号码
                    currentModel.ProductNo = tempModel.Prodtyp;

                    //保单团个性质代码
                    currentModel.GPFlag = "02";

                    // 主险保险险种号码
                    var tempLCProductGroup = businessModel.lstTEMP_LCProductGroup.Where(e =>
                    e.GrpPolicyNo.Equals(currentModel.GrpPolicyNo) &&
                    e.PolicyNo.Equal(currentModel.PolicyNo.PadLeft(8,'0')) &&
                    e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                    currentModel.MainProductNo = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductNo;

                    //主附险性质代码
                    currentModel.MainProductFlag = tempLCProductGroup == null ? string.Empty : tempLCProductGroup.MainProductFlag;

                    //产品编码
                    string tempProductCode = string.IsNullOrWhiteSpace(tempModel.ProductCode) ? string.Empty : tempModel.ProductCode.Trim();
                    currentModel.ProductCode = (tempProductCode.Equals("GIP")
                        || tempProductCode.Equals("GOP") || tempProductCode.Equals("GMA")) ? "GHB" : tempProductCode;

                    //责任代码
                    //var templstZaiBaoProductInfo2 = businessModel.lstZaiBaoProductInfo.Where(e => e.ProductCode.Equals(tempModel.ProductCode)).FirstOrDefault();
                    currentModel.LiabilityCode = Common.GetLiabilityCode(currentModel.ProductCode);

                    //责任名称
                    var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryCode.Equal(currentModel.LiabilityCode)).FirstOrDefault();
                    currentModel.LiabilityName = tempCategory == null ? string.Empty : tempCategory.CategoryName;

                    //责任分类代码
                    currentModel.Classification = tempCategory == null ? string.Empty : tempCategory.LiabilityCategoryCode;

                    // 续期续保次数
                    currentModel.RenewalTimes = "0";

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
                            strSignDate = tempSignDate.ToString("yyyy/MM/dd");
                        }
                    }
                    currentModel.SignDate = strSignDate;

                    //业务类型 
                    if (!string.IsNullOrEmpty(strSignDate))
                    {
                        bool checkResult = Common.CheckEventType(yearMonthDay, strSignDate);

                        if (checkResult)
                        {
                            currentModel.EventType = "01";
                        }
                        else
                        {
                            if (!currentModel.ProductCode.ToUpper().Equal("HBA"))
                            {
                                currentModel.EventType = "02";
                            }
                            else
                            {
                                currentModel.EventType = "03";
                            }
                        }
                    }
                    else
                    {
                        currentModel.EventType = string.Empty;
                    }

                    //保险责任生效日期
                    currentModel.EffDate = tempLcGrpContGroup == null ? string.Empty : tempLcGrpContGroup.EffDate;

                    //PolYear 所跑数据年份减去签单日期年份
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
                    currentModel.RiskAmnt = ConfigInformation.NumberValue;

                    //保费
                    currentModel.Premium = Common.ConvertToStrToStrDecimal(tempModel.Pprem);

                    //保险账户价值
                    //var tempInsureAcc = businessModel.lstTEMP_LCInsureAcc.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo)
                    //&& e.ProductNo.Equal(currentModel.ProductNo)).FirstOrDefault();

                    currentModel.AccountValue = "0";

                    //临分标记
                    currentModel.FacultativeFlag = string.Empty;

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
                    var tempInsuredGroup = businessModel.lst_LCInsuredGroup.Where(e => e.PolicyNo.Equal(currentModel.PolicyNo.PadLeft(8,'0'))
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

                    //被保人数
                    currentModel.InsurePeoples = "1";

                    //再保险公司名称 
                    currentModel.ReinsurerName = Common.DefaultCommanyName;

                    //再保险公司代码
                    var tempReinsurer = reinsurer.GetReinsurerInforByName(currentModel.ReinsurerName);
                    currentModel.ReinsurerCode = tempReinsurer == null ? string.Empty : tempReinsurer.ReinsurerCode;

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
                    if (templstZaiBaoProductInfo == null)
                    {
                        currentModel.SaparateFlag = string.Empty;
                    }
                    else
                    {
                        string tempQuotaSharePercentage = (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                    templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";
                        currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;
                    }




                    //分保保额
                    currentModel.ReinsuranceAmnt = ConfigInformation.NumberValue;

                    //自留额
                    currentModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       templstZaiBaoProductInfo.RetentionAmount;

                    //货币代码
                    currentModel.Currency = "156";

                    //分保比例
                    currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.QuotaSharePercentage;

                    //分保费
                    currentModel.ReinsurancePremium = Common.ConvertToStrToStrDecimal(tempModel.RIAnnualizedPremiumTot);

                    //分保佣金
                    currentModel.ReinsuranceCommssion = Common.ConvertToStrToStrDecimal(tempModel.ReinsuranceCommssion);

                    //ReComputationsDate
                    currentModel.ReComputationsDate = Common.GetLastDayOfMonth(yearMonthDay);

                    //账单归属日期
                    currentModel.AccountGetDate = Common.GetLastDayOfMonth(yearMonthDay);

                    serialNumber++;

                    lRContModelList.Add(currentModel);
                }
            }
        }

        private void SetIndividualDataToModel(ContractInfoBusiness businessModel, string yearMonthDay,
          ref int serialNumber)
        {
            if (businessModel.lstInforceBusinessListing.Count > 0)
            {
                foreach (var tempModel in businessModel.lstInforceBusinessListing)
                {
                    LRContModel currentModel = new LRContModel();

                    //交易编码
                    currentModel.TransactionNo = CommFuns.GetTransactionNo4(serialNumber, yearMonthDay);

                    //保险机构代码
                    currentModel.CompanyCode = CommFuns.OriganizationCode;

                    //团体保单号
                    currentModel.GrpPolicyNo = string.Empty;

                    //团体保单险种号码
                    currentModel.GrpProductNo = string.Empty;

                    //个人保单号
                    string tempPolicyNo = string.IsNullOrWhiteSpace(tempModel.PolicyNo) ? string.Empty : tempModel.PolicyNo.Trim();
                    currentModel.PolicyNo = tempPolicyNo.PadLeft(8, '0');

                    //主附险性质代码
                    currentModel.MainProductFlag = this.GetMainProductFlag(tempModel.ProductCode);

                    // 个单保险险种号码
                    var tempLCProduct = businessModel.lstTEMP_LCProduct.Where(e =>
                    e.PolicyNo.Equal(currentModel.PolicyNo) &&
                     e.ProductCode.Equal(tempModel.ProductCode) &&
                     e.MainProductFlag.Equals(currentModel.MainProductFlag)).FirstOrDefault();

                    currentModel.ProductNo = tempLCProduct == null ? string.Empty : tempLCProduct.ProductNo;

                    //保单团个性质代码
                    currentModel.GPFlag = "01";

                    // 主险保险险种号码
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

                    // 续期续保次数
                    var tempLCCont = businessModel.lstTEMP_LCCont.Where(e => e.PolicyNo.Equals(tempModel.PolicyNo)).FirstOrDefault();
                    currentModel.RenewalTimes = tempLCCont == null ? string.Empty : tempLCCont.RenewalTimes;

                    //保险期限类型
                    var tempProductModel = businessModel.lstTEMP_LMProductModel.Where(e => e.ProductCode == currentModel.ProductCode).FirstOrDefault();
                    currentModel.TermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                    //管理机构代码
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

                    //业务类型 
                    if (!string.IsNullOrEmpty(strSignDate))
                    {
                        bool checkResult = Common.CheckEventType(yearMonthDay, strSignDate);

                        if (checkResult)
                        {
                            currentModel.EventType = "01";
                        }
                        else
                        {
                            if (!currentModel.ProductCode.ToUpper().Equal("HBA"))
                            {
                                currentModel.EventType = "02";
                            }
                            else
                            {
                                currentModel.EventType = "03";
                            }
                        }
                    }
                    else
                    {
                        currentModel.EventType = string.Empty;
                    }

                    //保险责任生效日期
                    currentModel.EffDate = tempLCProduct == null ? string.Empty : tempLCProduct.EffDate;

                    //PolYear 所跑数据年份减去签单日期年份
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
                    currentModel.ProfessionalFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.ProfessionalFee);

                    //次标准体加费金额
                    currentModel.SubStandardFee = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.SubStandardFee);

                    //EM加点
                    currentModel.EMRate = tempLCProduct == null ? string.Empty : Common.ConvertToStrToStrDecimal(tempLCProduct.EMRate);

                    //建工险标志
                    currentModel.ProjectFlag = ConfigInformation.TextValue;

                    //被保人数
                    currentModel.InsurePeoples = "1";

                    //再保险公司名称
                    currentModel.ReinsurerName = tempModel.CompanyName;

                    //再保险公司代码
                    var tempReinsurer = reinsurer.GetReinsurerInforByName(currentModel.ReinsurerName);
                    currentModel.ReinsurerCode = tempReinsurer == null ? string.Empty : tempReinsurer.ReinsurerCode;

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
                    string tempQuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty : (templstZaiBaoProductInfo.QuotaSharePercentage == "0" ||
                        templstZaiBaoProductInfo.QuotaSharePercentage == "0.00") ? "0" : "1";
                    currentModel.SaparateFlag = templstZaiBaoProductInfo == null ? string.Empty : tempQuotaSharePercentage;

                    //分保保额
                    if (tempModel.IsMrHealth)
                    {
                        currentModel.ReinsuranceAmnt = tempModel.SumReinsured;
                    }
                    else
                    {
                        string para1 = string.IsNullOrWhiteSpace(tempModel.SumReinsured) ? string.Empty : tempModel.SumReinsured.Trim();
                        string para2 = string.IsNullOrWhiteSpace(tempModel.SumReinsured2) ? string.Empty : tempModel.SumReinsured2.Trim();
                        currentModel.ReinsuranceAmnt = this.GetReinsuranceAmnt(para1, para2);
                    }

                    //自留额
                    currentModel.RetentionAmount = templstZaiBaoProductInfo == null ? string.Empty :
                       templstZaiBaoProductInfo.RetentionAmount;

                    //货币代码
                    currentModel.Currency = "156";

                    //分保比例
                    currentModel.QuotaSharePercentage = templstZaiBaoProductInfo == null ? string.Empty :
                        templstZaiBaoProductInfo.QuotaSharePercentage;

                    //分保费
                    if (tempModel.IsMrHealth)
                    {
                        currentModel.ReinsurancePremium = Common.ConvertToStrToStrDecimal(tempModel.MonthlyReinsurancePremium);
                    }
                    else
                    {
                        string para1 = string.IsNullOrWhiteSpace(tempModel.MonthlyReinsurancePremium) ?
                            string.Empty : tempModel.MonthlyReinsurancePremium.Trim();
                        string para2 = string.IsNullOrWhiteSpace(tempModel.MonthlyReinsurancePremium2) ?
                            string.Empty : tempModel.MonthlyReinsurancePremium2.Trim();

                        currentModel.ReinsurancePremium = this.GetReinsuranceAmnt(para1, para2);
                    }

                    //分保佣金
                    if (tempModel.IsMrHealth)
                    {
                        currentModel.ReinsuranceCommssion = Common.ConvertToStrToStrDecimal(tempModel.MonthlyReinsuranceCommission);
                    }
                    else
                    {
                        string para1 = string.IsNullOrWhiteSpace(tempModel.MonthlyReinsuranceCommission) ?
                                             string.Empty : tempModel.MonthlyReinsuranceCommission.Trim();
                        string para2 = string.IsNullOrWhiteSpace(tempModel.MonthlyReinsuranceCommission2) ?
                            string.Empty : tempModel.MonthlyReinsuranceCommission2.Trim();

                        currentModel.ReinsuranceCommssion = this.GetReinsuranceAmnt(para1, para2);
                    }

                    //ReComputationsDate
                    currentModel.ReComputationsDate = Common.GetLastDayOfMonth(yearMonthDay);

                    //账单归属日期
                    currentModel.AccountGetDate = Common.GetLastDayOfMonth(yearMonthDay);

                    serialNumber++;

                    lRContModelList.Add(currentModel);
                }
            }
        }

        private string GetReinsuranceAmnt(string sumReinsured, string sumReinsured2)
        {
            string result = string.Empty;

            if (string.IsNullOrWhiteSpace(sumReinsured))
            {
                sumReinsured = "0";
            }

            if (string.IsNullOrWhiteSpace(sumReinsured2))
            {
                sumReinsured2 = "0";
            }

            result = (decimal.Parse(sumReinsured.Trim()) + decimal.Parse(sumReinsured2.Trim())).ToString("0.00");

            return result;
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
