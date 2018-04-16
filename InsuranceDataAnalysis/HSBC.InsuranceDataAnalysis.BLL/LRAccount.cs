using HSBC.InsuranceDataAnalysis.ExcelCore;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LRAccount
    {
        private const string origanizationCode = "000131";
        List<LRAccountModel> listLRAccount = new List<LRAccountModel>();
        public void WriteLRAccountSheet(ContractInfoBusiness contractInfoBusiness, string OutPutFolderPath, string dateyyyymm)
        {
            IExcel excelApp = new ExcelCore.ExcelCore();
            try
            {
                ProcessLogProxy.Normal("Start building LRAccount excel");
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LRAccount + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LRAccount);//创建模板
                GetLRAccountData(contractInfoBusiness, dateyyyymm);
                excelApp.OpenExcel(excelPath, false);
                for (int i = 0; i < listLRAccount.Count; i++)
                {
                    var model = listLRAccount[i];
                    excelApp.SetCellValue("Sheet1", i + 2, "A", model.TransactionNo);
                    excelApp.SetCellValue("Sheet1", i + 2, "B", model.CompanyCode);
                    excelApp.SetCellValue("Sheet1", i + 2, "C", model.AccountID);
                    excelApp.SetCellValue("Sheet1", i + 2, "D", model.AccountingPeriodfrom);
                    excelApp.SetCellValue("Sheet1", i + 2, "E", model.AccountingPeriodto);
                    excelApp.SetCellValue("Sheet1", i + 2, "F", model.ReinsurerCode);
                    excelApp.SetCellValue("Sheet1", i + 2, "G", model.ReinsurerName);
                    excelApp.SetCellValue("Sheet1", i + 2, "H", model.ReInsuranceContNo);
                    excelApp.SetCellValue("Sheet1", i + 2, "I", model.ReInsuranceContName);
                    excelApp.SetCellValue("Sheet1", i + 2, "J", model.Currency);
                    excelApp.SetCellValue("Sheet1", i + 2, "K", model.ReinsurancePremium);
                    excelApp.SetCellValue("Sheet1", i + 2, "L", model.ReinsuranceCommssionRate);
                    excelApp.SetCellValue("Sheet1", i + 2, "M", model.ReinsuranceCommssion);
                    excelApp.SetCellValue("Sheet1", i + 2, "N", model.ReturnReinsurancePremium);
                    excelApp.SetCellValue("Sheet1", i + 2, "O", model.ReturnReinsuranceCommssion);
                    excelApp.SetCellValue("Sheet1", i + 2, "P", model.ReturnSurrenderPay);
                    excelApp.SetCellValue("Sheet1", i + 2, "Q", model.ReturnClaimPay);
                    excelApp.SetCellValue("Sheet1", i + 2, "R", model.ReturnMaturity);
                    excelApp.SetCellValue("Sheet1", i + 2, "S", model.ReturnAnnuity);
                    excelApp.SetCellValue("Sheet1", i + 2, "T", model.ReturnLivBene);
                    excelApp.SetCellValue("Sheet1", i + 2, "U", model.AccountStatus);
                    excelApp.SetCellValue("Sheet1", i + 2, "V", model.PairingStatus);
                    excelApp.SetCellValue("Sheet1", i + 2, "W", model.PairingDate);
                    excelApp.SetCellValue("Sheet1", i + 2, "X", model.CurrentRate);

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

        private void GetLRAccountData(ContractInfoBusiness contractInfoBusiness, string dateyyyymm)
        {
            var listLRInsureContModel = contractInfoBusiness.lstLRInsureContModel;
            var listStatement = contractInfoBusiness.lstInsuranceReinsuranceStatementModel;
            for (int i = 0; i < listStatement.Count; i++)
            {
                var model = listStatement[i];
                try
                {
                    LRInsureContModel lRInsureContModel = GetLRInsureContModel(listLRInsureContModel, model);
                    if (lRInsureContModel == null)
                    {
                        throw new Exception("Get " + model.ToCompanyName + " information error");
                    }
                    LRAccountModel lrAccountModel = new LRAccountModel();
                    var reinsurer = new Reinsurer().GetReinsurerInforByName(model.ToCompanyName);
                    var reinsurerCode = reinsurer == null ? string.Empty : reinsurer.ReinsurerCode;
                    lrAccountModel.TransactionNo = CommFuns.GetTransactionNo(i + 1, dateyyyymm);
                    lrAccountModel.CompanyCode = origanizationCode;
                    lrAccountModel.AccountID = lRInsureContModel.MainReInsuranceContNo + dateyyyymm.Substring(0, 6);//账单编号
                    lrAccountModel.AccountingPeriodfrom = Convert.ToDateTime(dateyyyymm.Substring(0, 4) + "-" + dateyyyymm.Substring(4, 2) + "-01").ToString("yyyy/MM/dd");
                    lrAccountModel.AccountingPeriodto = Convert.ToDateTime(dateyyyymm.Substring(0, 4) + "-" + dateyyyymm.Substring(4, 2) + "-01").AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd");
                    lrAccountModel.ReinsurerCode = reinsurerCode;
                    lrAccountModel.ReinsurerName = reinsurer == null ? string.Empty : reinsurer.ReinsurerChineseName;
                    lrAccountModel.ReInsuranceContNo = lRInsureContModel.MainReInsuranceContNo;//合同号码
                    lrAccountModel.ReInsuranceContName = lRInsureContModel.ReInsuranceContName;//合同名称
                    lrAccountModel.Currency = "CNY";
                    lrAccountModel.ReinsurancePremium = decimal.Round(decimal.Parse(model.Debit.ReinsurancePremiums), 2).ToString();//
                    lrAccountModel.ReinsuranceCommssionRate = "0.5";//分保佣金、分保费50%
                    lrAccountModel.ReinsuranceCommssion = decimal.Round(decimal.Parse(model.Credit.ReinsuranceCommissions), 2).ToString();//
                    lrAccountModel.ReturnReinsurancePremium = "0";
                    lrAccountModel.ReturnReinsuranceCommssion = "0";
                    lrAccountModel.ReturnSurrenderPay = "0";
                    lrAccountModel.ReturnClaimPay = decimal.Round(decimal.Parse(model.Credit.ReinsuranceClaimAmounts), 2).ToString();
                    lrAccountModel.ReturnMaturity = "0";
                    lrAccountModel.ReturnAnnuity = "0";
                    lrAccountModel.ReturnLivBene = "0";
                    lrAccountModel.AccountStatus = "1";
                    lrAccountModel.PairingStatus = "2";
                    lrAccountModel.PairingDate = lrAccountModel.AccountingPeriodto == null ? "" : Convert.ToDateTime(lrAccountModel.AccountingPeriodto).ToString("yyyy/MM/dd");
                    lrAccountModel.CurrentRate = "1";

                    listLRAccount.Add(lrAccountModel);
                }
                catch (Exception ex)
                {
                    throw;
                }
            }
        }

        private static LRInsureContModel GetLRInsureContModel(List<LRInsureContModel> listLRInsureContModel, InsuranceReinsuranceStatement model)
        {
            LRInsureContModel lRInsureContModel = new LRInsureContModel();
            try
            {
                if (model.FilePath.Contains("MR_Health"))
                {
                    lRInsureContModel = listLRInsureContModel.Where(A => A.ReinsurerName == model.ToCompanyName && A.ContOrAmendmentType == "1" && A.ReInsuranceContName.Contains("健康")).ToList().FirstOrDefault();
                }
                else if (model.FilePath.Contains("MR_life"))
                {
                    lRInsureContModel = listLRInsureContModel.Where(A => A.ReinsurerName == model.ToCompanyName && A.ContOrAmendmentType == "1" && A.ReInsuranceContName.Contains("人寿")).ToList().FirstOrDefault();
                }
                else
                {
                    lRInsureContModel = listLRInsureContModel.Where(A => A.ReinsurerName == model.ToCompanyName && A.ContOrAmendmentType == "1").ToList().FirstOrDefault();
                }
            }
            catch (Exception EX)
            {
                throw new Exception("The billing information does not match the contract");
            }

            return lRInsureContModel;
        }
    }
}
