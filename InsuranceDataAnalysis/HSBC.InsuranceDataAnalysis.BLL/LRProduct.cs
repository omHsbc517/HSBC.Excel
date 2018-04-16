using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.ExcelCore;
using System.Text.RegularExpressions;
using HSBC.InsuranceDataAnalysis.Utils;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LRProduct
    {

        IExcel excelApp;
        Reinsurer reinsurer = new Reinsurer();
        private const string LRProductSheetName = "Sheet1";
        private const string origanizationCode = "000131";
        #region shangjunqi
        public void InputDataToLRProductSheet(ContractInfoBusiness contractInfoBusiness,
            string OutPutFolderPath, string yearMonthDay)
        {
            excelApp = new ExcelCore.ExcelCore();
            try
            {
                //contractInfoBusiness.lstLRInsureContModel
                ProcessLogProxy.Normal("Start building LRProduct excel");
                var lstHanReModel = contractInfoBusiness.lstHanReModel;
                var lstMuReModel = contractInfoBusiness.lstMuReModel;
                var lstRGAModel = contractInfoBusiness.lstRGAModel;
                var lstSwissReModel = contractInfoBusiness.lstSwissReModel;
                var lstHugeDisasterModel = contractInfoBusiness.lstHugeDisasterModel;
                var lstTEMP_LMProductModel = contractInfoBusiness.lstTEMP_LMProductModel;
                var lstProductInfoModel = contractInfoBusiness.lstProductInfoModel;
                //var lstTEMP_LMLiabilityModel = contractInfoBusiness.lstTEMP_LMLiabilityModel;
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LRProduct + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LRProduct);
                excelApp.OpenExcel(excelPath, false);
                excelApp.SelectSheet(LRProductSheetName);

                int serialNumber = 1;
                int startRowIndex = 2;

                this.Test2(lstHanReModel, yearMonthDay, ref serialNumber, ref startRowIndex,
                    lstTEMP_LMProductModel, contractInfoBusiness);
                this.Test2(lstMuReModel, yearMonthDay, ref serialNumber, ref startRowIndex,
                    lstTEMP_LMProductModel, contractInfoBusiness);
                this.Test2(lstRGAModel, yearMonthDay, ref serialNumber, ref startRowIndex,
                    lstTEMP_LMProductModel, contractInfoBusiness);
                this.Test2(lstSwissReModel, yearMonthDay, ref serialNumber, ref startRowIndex,
                    lstTEMP_LMProductModel, contractInfoBusiness);

                foreach (var temp in lstHugeDisasterModel)
                {
                    var tempList = lstTEMP_LMProductModel.Where(e => e.StartDate <= temp.EffectiveDate);

                    var tempModel = (from lmProduct in tempList
                                     join productInfo in lstProductInfoModel
                                     on lmProduct.ProductCode equals productInfo.ProductCode
                                     select new
                                     {
                                         productCode = lmProduct.ProductCode,
                                         productName = lmProduct.ProductName,
                                         productType = lmProduct.ProductType,
                                         TermType = lmProduct.TermType
                                     }).ToList();

                    int contractOrder = 0;
                    foreach (var temp1 in tempModel)
                    {
                        this.Test3(yearMonthDay, serialNumber, startRowIndex, temp, contractOrder,
                            temp1.productCode, temp1.productName, temp1.productType,
                            temp1.TermType, contractInfoBusiness);

                        contractOrder += 1;
                        serialNumber++;
                        startRowIndex++; 
                    }

                    //foreach (var temp1 in tempModel)
                    //{
                    //    this.Test3(yearMonthDay, serialNumber, ref startRowIndex, temp, contractOrder,
                    //        temp1.productCode, temp1.productName, temp1.productType, temp1.LiabilityCode,
                    //        temp1.TermType, contractInfoBusiness);

                    //    contractOrder += 1;
                    //}
                }
                ProcessLogProxy.SuccessMessage("Build Success");
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                excelApp.SetSheetAutoFit(LRProductSheetName);
                excelApp.Save();
                excelApp.Close();
            }
        }

        private string GetRiMethodICodeByName(string name)
        {
            string result = string.Empty;

            if (name.Trim().Equals("溢额"))
            {
                result = "1";
            }
            else if (name.Trim().Equals("成数"))
            {
                result = "2";
            }
            else if (name.Trim().Equals("超赔") || name.Trim().Equals("非比例"))
            {
                result = "4";
            }
            else if (name.Trim().Contains("溢额") && name.Trim().Contains("成数"))
            {
                result = "3";
            }

            return result;
        }

        private void Test2(List<RIContractInfo> lstHanReModel, string yearMonthDay,
            ref int serialNumber, ref int startRowIndex,
            List<TEMP_LMProduct> lstTEMP_LMProductModel,
            ContractInfoBusiness contractInfoBusiness)
        {
            //  从第二行开始输入数据
            if (lstHanReModel.Count > 0)
            {
                int contractOrder = 1;
                foreach (var temp in lstHanReModel)
                {
                    var tempProductModel = lstTEMP_LMProductModel.Where(e => e.ProductCode == temp.ProductCode).FirstOrDefault();
                    string tempProductType = tempProductModel == null ? string.Empty : tempProductModel.ProductType;
                    var tempTermType = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                    //var tempLMLiabilityModel = lstTEMP_LMLiabilityModel.Where(e => e.ProductCode == temp.ProductCode).FirstOrDefault();
                    //string tempLiabilityCode = tempLMLiabilityModel == null ? string.Empty : tempLMLiabilityModel.LiabilityCode;

                    var tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryName.Equals(temp.BenefitReinsured)).FirstOrDefault();
                    string tempLiabilityCode = tempCategory == null ? string.Empty : tempCategory.CategoryCode;

                    if (!string.IsNullOrWhiteSpace(temp.ProductCode))
                    {
                        this.Test(yearMonthDay, serialNumber, startRowIndex, temp,
                            contractOrder, 0, tempProductType, tempLiabilityCode, tempTermType, contractInfoBusiness);
                        startRowIndex++;
                        serialNumber++;
                    }

                    bool isAdd = false;
                    if (temp.lstChildRIContractInfo != null && temp.lstChildRIContractInfo.Count > 0)
                    {
                        for (int i = 0; i < temp.lstChildRIContractInfo.Count; i++)
                        {
                            var tempChildModel = temp.lstChildRIContractInfo[i];

                            tempProductModel = lstTEMP_LMProductModel.Where(e => e.ProductCode == tempChildModel.ProductCode).FirstOrDefault();
                            string tempProductType2 = tempProductModel == null ? string.Empty : tempProductModel.ProductType;
                            var tempTermType2 = tempProductModel == null ? string.Empty : tempProductModel.TermType;

                            //tempLMLiabilityModel = lstTEMP_LMLiabilityModel.Where(e => e.ProductCode == tempChildModel.ProductCode).FirstOrDefault();
                            //string tempLiabilityCode2 = tempLMLiabilityModel == null ? string.Empty : tempLMLiabilityModel.LiabilityCode;

                            tempCategory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryName.Equals(tempChildModel.BenefitReinsured)).FirstOrDefault();
                            string tempLiabilityCode2 = tempCategory == null ? string.Empty : tempCategory.CategoryCode;

                            if (!string.IsNullOrWhiteSpace(tempChildModel.ProductCode))
                            {
                                this.Test(yearMonthDay, serialNumber, startRowIndex, tempChildModel,
                                                contractOrder, i + 1, tempProductType2, tempLiabilityCode2, tempTermType2, contractInfoBusiness, false);
                                serialNumber++;
                                startRowIndex++;
                            }
                        }
                    }
                    else
                    {
                        isAdd = true;
                    }

                    contractOrder++;
                    if (isAdd)
                    {
                        startRowIndex++;
                        serialNumber++;
                    }
                }
            }
        }

        private void Test(string yearMonthDay, int serialNumber, int rowIndex,
            RIContractInfo temp, int contractOrder, int currentIndex,
            string productType, string liabilityCode, string termType,
            ContractInfoBusiness contractInfoBusiness, bool isMainContract = true)
        {
            ZaiBaoProductInfo tempModel = new ZaiBaoProductInfo();

            string currentTransactionNo = CommFuns.GetTransactionNo(serialNumber, yearMonthDay);

            excelApp.SetCellValue(rowIndex, "A", currentTransactionNo);
            excelApp.SetCellValue(rowIndex, "B", origanizationCode);

            var tempEntity = reinsurer.GetReinsurerInforByName(temp.Reinsurer);
            var tempCompanyCode = tempEntity == null ? string.Empty : tempEntity.ReinsurerCode;

            string currentReInsuranceContNo = string.Empty;
            string mainContractCode = string.Empty;

            #region  废代码
            //string mainContractCode = "RICN" + tempCompanyCode + "M"
            //    + contractOrder.ToString().PadLeft(2, '0') + "000";

            //if (isMainContract)
            //{
            //    currentReInsuranceContNo = "RICN" + tempCompanyCode + temp.ContractTypeSign
            //    + contractOrder.ToString().PadLeft(2, '0') + "000";
            //}
            //else
            //{
            //    currentReInsuranceContNo = "RICN" + tempCompanyCode + temp.ContractTypeSign
            //    + contractOrder.ToString().PadLeft(2, '0') + currentIndex.ToString().PadLeft(3, '0');
            //}
            #endregion

            var referenceEntity = contractInfoBusiness.lstLRInsureContModel.Where(e => e.ReInsuranceContName.Equals(temp.TreatyName)
             && e.ReinsurerCode.Equals(tempCompanyCode)).FirstOrDefault();

            currentReInsuranceContNo = referenceEntity == null ? string.Empty : referenceEntity.ReInsuranceContNo;
            mainContractCode = referenceEntity == null ? string.Empty : referenceEntity.MainReInsuranceContNo;

            excelApp.SetCellValue(rowIndex, "C", currentReInsuranceContNo);
            excelApp.SetCellValue(rowIndex, "D", temp.TreatyName);

            excelApp.SetCellValue(rowIndex, "E", string.Empty);
            excelApp.SetCellValue(rowIndex, "F", mainContractCode);

            string tempContractType = temp.ContractTypeSign.Equals("M") ? "1" : "2";

            excelApp.SetCellValue(rowIndex, "G", tempContractType);

            excelApp.SetCellValue(rowIndex, "H", temp.ProductCode);

            excelApp.SetCellValue(rowIndex, "I", temp.ReinsurerName);

            //string tempGpfFlag = string.Empty;
            string productCodeFirstChar = string.Empty;

            if (!string.IsNullOrEmpty(temp.ProductCode))
            {
                productCodeFirstChar = temp.ProductCode.Trim().Substring(0, 1);

                productCodeFirstChar = productCodeFirstChar.Equals("G") ? "02" : "01";

                excelApp.SetCellValue(rowIndex, "J", productCodeFirstChar);
            }

            excelApp.SetCellValue(rowIndex, "K", productType);
            excelApp.SetCellValue(rowIndex, "L", liabilityCode);

            excelApp.SetCellValue(rowIndex, "M", temp.BenefitReinsured);
            excelApp.SetCellValue(rowIndex, "N", tempCompanyCode);

            //excelApp.SetCellValue(rowIndex, "O", temp.Reinsurer);
            excelApp.SetCellValue(rowIndex, "O", tempEntity == null ? string.Empty : tempEntity.ReinsurerChineseName);

            excelApp.SetCellValue(rowIndex, "P", temp.RIratio);

            string tempMethodCode = this.GetRiMethodICodeByName(temp.RImethodI);

            excelApp.SetCellValue(rowIndex, "Q", tempMethodCode);

            excelApp.SetCellValue(rowIndex, "R", "04");
            excelApp.SetCellValue(rowIndex, "S", termType);

            string currentRetentionAmount = this.GetRetentionAmount(tempMethodCode, temp.Retention);
            excelApp.SetCellValue(rowIndex, "T", currentRetentionAmount);

            string currentRetentionPercentage = this.GetRetentionPercentage(tempMethodCode, temp.Retention);
            excelApp.SetCellValue(rowIndex, "U", currentRetentionPercentage);

            string QuotaSharePercentage = "0";
            if (tempMethodCode.Equals("2") || tempMethodCode.Equals("3"))
            {
                QuotaSharePercentage = (1 - decimal.Parse(currentRetentionPercentage)).ToString("0.00");
            }
            excelApp.SetCellValue(rowIndex, "V", QuotaSharePercentage);

            tempModel.TransactionNo = currentTransactionNo;
            tempModel.CompanyCode = origanizationCode;
            tempModel.ReInsuranceContNo = currentReInsuranceContNo;
            tempModel.ReInsuranceContName = temp.TreatyName;

            tempModel.ReInsuranceContTitle = ConfigInformation.TextValue;
            tempModel.MainReInsuranceContNo = mainContractCode;
            tempModel.ContOrAmendmentType = tempContractType;
            tempModel.ProductCode = temp.ProductCode;

            tempModel.ProductName = temp.ReinsurerName;
            tempModel.GPFlag = productCodeFirstChar;
            tempModel.ProductType = productType;
            tempModel.LiabilityCode = liabilityCode;

            tempModel.LiabilityName = temp.BenefitReinsured;
            tempModel.ReinsurerCode = tempCompanyCode;
            tempModel.ReinsurerName = temp.Reinsurer;
            tempModel.ReinsuranceShare = temp.RIratio;

            tempModel.ReinsurMode = tempMethodCode;
            tempModel.ReInsuranceType = "04";
            tempModel.TermType = termType;
            tempModel.RetentionAmount = currentRetentionAmount;

            tempModel.RetentionPercentage = currentRetentionPercentage;
            tempModel.QuotaSharePercentage = QuotaSharePercentage;

            contractInfoBusiness.lstZaiBaoProductInfo.Add(tempModel);
        }

        private void Test3(string yearMonthDay, int serialNumber, int rowIndex,
      HugeDisasterModel temp, int contractOrder, string productCode,
      string productName, string productType,  string termType, ContractInfoBusiness business)
        {
            ZaiBaoProductInfo tempModel = new ZaiBaoProductInfo();
            string currentTransactionNo = CommFuns.GetTransactionNo(serialNumber, yearMonthDay);

            excelApp.SetCellValue(rowIndex, "A", currentTransactionNo);
            excelApp.SetCellValue(rowIndex, "B", origanizationCode);

            var tempEntity = reinsurer.GetReinsurerInforByName(temp.Reinsurer);
            var tempCompanyCode = tempEntity == null ? string.Empty : tempEntity.ReinsurerCode;

            //string currentReInsuranceContNo = "RICN" + tempCompanyCode + "M"
            //    + contractOrder.ToString().PadLeft(2, '0') + "000";

            var referenceEntity = business.lstLRInsureContModel.Where(e => e.ReInsuranceContName.Equals(temp.TreatyName)
         && e.ReinsurerCode.Equals(tempCompanyCode)).FirstOrDefault();

            string currentReInsuranceContNo = referenceEntity == null ? string.Empty : referenceEntity.ReInsuranceContNo;
            //mainContractCode = referenceEntity == null ? string.Empty : referenceEntity.MainReInsuranceContNo;

            excelApp.SetCellValue(rowIndex, "C", currentReInsuranceContNo);
            excelApp.SetCellValue(rowIndex, "D", temp.TreatyName);
            excelApp.SetCellValue(rowIndex, "E", string.Empty);
            excelApp.SetCellValue(rowIndex, "F", currentReInsuranceContNo);

            excelApp.SetCellValue(rowIndex, "G", "1");

            excelApp.SetCellValue(rowIndex, "H", productCode);

            excelApp.SetCellValue(rowIndex, "I", productName);

            //string tempGpfFlag = string.Empty;
            string productCodeFirstChar = string.Empty;

            if (!string.IsNullOrEmpty(productCode))
            {
                productCodeFirstChar = productCode.Trim().Substring(0, 1);

                productCodeFirstChar = productCodeFirstChar.Equals("G") ? "02" : "01";

                excelApp.SetCellValue(rowIndex, "J", productCodeFirstChar);
            }

            excelApp.SetCellValue(rowIndex, "K", productType);

            var tempCateGory = PersonalLiabilityCategory.LstCategory.Where(e => e.CategoryName.Equals(temp.BenefitReinsured)).FirstOrDefault();
            //excelApp.SetCellValue(rowIndex, "L", liabilityCode);
            excelApp.SetCellValue(rowIndex, "L", tempCateGory == null ? string.Empty : tempCateGory.CategoryCode);

            excelApp.SetCellValue(rowIndex, "M", temp.BenefitReinsured);
            excelApp.SetCellValue(rowIndex, "N", tempCompanyCode);

            excelApp.SetCellValue(rowIndex, "O", tempEntity == null ? string.Empty : tempEntity.ReinsurerChineseName);
            //excelApp.SetCellValue(rowIndex, "O", temp.Reinsurer);

            excelApp.SetCellValue(rowIndex, "P", temp.RIratio);

            string tempMethodCode = this.GetRiMethodICodeByName(temp.RImethodI);

            excelApp.SetCellValue(rowIndex, "Q", tempMethodCode);

            excelApp.SetCellValue(rowIndex, "R", "01");
            excelApp.SetCellValue(rowIndex, "S", termType);
            excelApp.SetCellValue(rowIndex, "T", "0");
            excelApp.SetCellValue(rowIndex, "U", "0");
            excelApp.SetCellValue(rowIndex, "V", "0");

            tempModel.TransactionNo = currentTransactionNo;
            tempModel.CompanyCode = origanizationCode;
            tempModel.ReInsuranceContNo = currentReInsuranceContNo;
            tempModel.ReInsuranceContName = temp.TreatyName;

            tempModel.ReInsuranceContTitle = ConfigInformation.TextValue;
            tempModel.MainReInsuranceContNo = currentReInsuranceContNo;
            tempModel.ContOrAmendmentType = "1";
            tempModel.ProductCode = productCode;

            tempModel.ProductName = productName;
            tempModel.GPFlag = productCodeFirstChar;
            tempModel.ProductType = productType;
            tempModel.LiabilityCode = tempCateGory == null ? string.Empty : tempCateGory.CategoryCode;

            tempModel.LiabilityName = temp.BenefitReinsured;
            tempModel.ReinsurerCode = tempCompanyCode;
            tempModel.ReinsurerName = temp.Reinsurer;
            tempModel.ReinsuranceShare = temp.RIratio;

            tempModel.ReinsurMode = tempMethodCode;
            tempModel.ReInsuranceType = "01";
            tempModel.TermType = termType;
            tempModel.RetentionAmount = "0";

            tempModel.RetentionPercentage = "0";
            tempModel.QuotaSharePercentage = "0";


            business.lstZaiBaoProductInfo.Add(tempModel);
        }

        private string GetRetentionAmount(string reinsurMode, string retention)
        {
            string retentionAmount = string.Empty;
            decimal tempResult = 0.0m;
            switch (reinsurMode)
            {
                case "1":
                    retentionAmount = retention;
                    break;
                case "3":
                    string[] tempList = retention.Split(',');
                    if (tempList.Length == 2)
                    {
                        retentionAmount = tempList[1].Replace(")", "");
                    }
                    break;
                default:
                    retentionAmount = "0";
                    break;
            }

            if (retentionAmount.Contains("万"))
            {
                Regex r = new Regex("\\d+\\.?\\d*");
                bool ismatch = r.IsMatch(retentionAmount);
                MatchCollection mc = r.Matches(retentionAmount);

                string result = string.Empty;
                for (int i = 0; i < mc.Count; i++)
                {
                    result += mc[i];//匹配结果是完整的数字，此处可以不做拼接的
                }
                tempResult = decimal.Parse(result) * 10000;

                retentionAmount = Convert.ToDecimal(tempResult).ToString("0.00");

                if (retentionAmount.Contains("."))
                {
                    retentionAmount = retentionAmount.TrimEnd('0').TrimEnd('.');
                }
            }
            return retentionAmount;
        }

        private string GetRetentionPercentage(string reinsurMode, string retention)
        {
            string retentionAmount = "0";
            decimal tempResult = 0.0m;
            bool isDivision = false;

            switch (reinsurMode)
            {
                case "2":
                    if (retention.Contains("%"))
                    {
                        retentionAmount = retention.Replace("%", "");
                        isDivision = true;
                    }
                    else
                    {
                        retentionAmount = retention;
                    }
                    break;
                case "3":
                    Regex rg = new Regex(@"\(.+\)");
                    var matchResult = rg.Match(retention);

                    if (matchResult.Success)
                    {
                        retentionAmount = matchResult.Value;

                        string[] tempList = retentionAmount.Split(',');
                        if (tempList.Length == 2)
                        {
                            retentionAmount = tempList[0].Replace("(", "").Replace("SA", "").Replace("%", "");
                            isDivision = true;
                        }
                    }
                    break;
                default:
                    retentionAmount = "0";
                    break;
            }

            if (isDivision)
            {
                tempResult = decimal.Parse(retentionAmount) / 100;
            }
            else
            {
                tempResult = decimal.Parse(retentionAmount);
            }
            
            retentionAmount = Convert.ToDecimal(tempResult).ToString("0.00");

            if (retentionAmount.Contains("."))
            {
                retentionAmount = retentionAmount.TrimEnd('0').TrimEnd('.');
            }
            return retentionAmount;
        }
        #endregion
    }
}
