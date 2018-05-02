using HSBC.InsuranceDataAnalysis.ExcelCore;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.BLL
{
    public class LRInsureCont
    {
        private List<LRInsureContModel> listLRInsureContModel = new List<LRInsureContModel>();
        private const string origanizationCode = "000131";
        private string LastDateOfMonth;
        public void WriteLRInsureContSheet(ContractInfoBusiness contractInfoBusiness, string OutPutFolderPath, string LastDateOfMonthyyyyMMdd)
        {
            IExcel excelApp = new ExcelCore.ExcelCore();
            try
            {
                ProcessLogProxy.Normal("Start building LRInsureCont excel");
                LastDateOfMonth = LastDateOfMonthyyyyMMdd;
                var excelPath = OutPutFolderPath + @"\TEMP_" + ExcelTemplateName.LRInsureCont + ".xlsx";
                ExcelTemplate excelTemplate = new ExcelTemplate();
                excelTemplate.CreateTemplate(excelApp, excelPath, ExcelTemplateName.LRInsureCont);//创建模板
                //GetLRInsureContData(contractInfoBusiness);//得到需写入excel的数据
                excelApp.OpenExcel(excelPath, false);
                for (int i = 0; i < contractInfoBusiness.lstLRInsureContModel.Count; i++)
                {
                    var model = listLRInsureContModel[i];
                    excelApp.SetCellValue("Sheet1", i + 2, "A", CommFuns.GetTransactionNo2(i + 1, LastDateOfMonth));
                    excelApp.SetCellValue("Sheet1", i + 2, "B", model.CompanyCode);
                    excelApp.SetCellValue("Sheet1", i + 2, "C", model.ReInsuranceContNo);
                    excelApp.SetCellValue("Sheet1", i + 2, "D", model.ReInsuranceContName);
                    excelApp.SetCellValue("Sheet1", i + 2, "E", model.ReInsuranceContTitle);
                    excelApp.SetCellValue("Sheet1", i + 2, "F", model.MainReInsuranceContNo);
                    excelApp.SetCellValue("Sheet1", i + 2, "G", model.ContOrAmendmentType);
                    excelApp.SetCellValue("Sheet1", i + 2, "H", model.ContAttribute);
                    excelApp.SetCellValue("Sheet1", i + 2, "I", model.ContStatus);
                    excelApp.SetCellValue("Sheet1", i + 2, "J", model.TreatyOrFacultativeFlag);
                    excelApp.SetCellValue("Sheet1", i + 2, "K", model.ContSigndate);
                    excelApp.SetCellValue("Sheet1", i + 2, "L", model.PeriodFrom);
                    excelApp.SetCellValue("Sheet1", i + 2, "M", model.PeriodTo);
                    excelApp.SetCellValue("Sheet1", i + 2, "N", model.ContType);
                    excelApp.SetCellValue("Sheet1", i + 2, "O", model.ReinsurerCode);
                    excelApp.SetCellValue("Sheet1", i + 2, "P", model.ReinsurerName);
                    excelApp.SetCellValue("Sheet1", i + 2, "Q", model.ChargeType);
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

        public void GetLRInsureContData(ContractInfoBusiness contractInfoBusiness)
        {
            GetLRInsureContByRIContractInfo(contractInfoBusiness.lstMuReModel);
            GetLRInsureContByRIContractInfo(contractInfoBusiness.lstHanReModel);
            GetLRInsureContByRIContractInfo(contractInfoBusiness.lstRGAModel);
            GetLRInsureContByRIContractInfo(contractInfoBusiness.lstSwissReModel);
            GetLRInsureContByHugeDisasterModel(contractInfoBusiness.lstHugeDisasterModel);
            contractInfoBusiness.lstLRInsureContModel = listLRInsureContModel;
        }


        private void GetLRInsureContByRIContractInfo(List<RIContractInfo> RIContractInfo, bool isChild = false, int contractNumber = 0, string MainReInsuranceContNo = "")
        {
            int ReInsuranceContNo = 0;
            for (int i = 0; i < RIContractInfo.Count; i++)
            {

                LRInsureContModel lRInsureContModel = new LRInsureContModel();
                var model = RIContractInfo[i];
                var reinsurer = new Reinsurer().GetReinsurerInforByName(model.Reinsurer);
                var reinsurerCode = reinsurer == null ? string.Empty : reinsurer.ReinsurerCode;
                string currentReInsuranceContNo = "";
                if (isChild)
                {
                    if (listLRInsureContModel.Where(A => A.ReInsuranceContName == model.TreatyName
                    && A.ReinsurerName == reinsurer.ReinsurerChineseName).ToList().Count() > 0)
                    {
                        continue;
                    }
                    currentReInsuranceContNo = "RICN" + reinsurerCode + model.ContractTypeSign
                + (contractNumber + 1).ToString().PadLeft(2, '0') + (ReInsuranceContNo++ + 1).ToString().PadLeft(3, '0');

                    //ReInsuranceContNo++;
                }
                else
                {
                    currentReInsuranceContNo = "RICN" + reinsurerCode + model.ContractTypeSign
                                   + (i + 1).ToString().PadLeft(2, '0') + "000";
                    MainReInsuranceContNo = currentReInsuranceContNo;
                }

                lRInsureContModel.TransactionNo = "";
                lRInsureContModel.CompanyCode = origanizationCode;
                lRInsureContModel.ReInsuranceContNo = currentReInsuranceContNo; //
                lRInsureContModel.ReInsuranceContName = model.TreatyName;
                lRInsureContModel.ReInsuranceContTitle = "";
                lRInsureContModel.MainReInsuranceContNo = MainReInsuranceContNo;//主合同号码
                lRInsureContModel.ContOrAmendmentType = "合同".Equals(model.ContOrAmendmentType) ? "1" : "2";
                lRInsureContModel.ContAttribute = "1";
                lRInsureContModel.ContStatus = "1";
                lRInsureContModel.TreatyOrFacultativeFlag = "1";
                lRInsureContModel.ContSigndate = Convert.ToDateTime(model.SignDate_INSH).ToString("yyyy/MM/dd");
                lRInsureContModel.PeriodFrom = Convert.ToDateTime(model.EffectiveDate).ToString("yyyy/MM/dd");
                lRInsureContModel.PeriodTo = "";
                lRInsureContModel.ContType = model.RImethodI == "非比例" ? "2" : "1";
                lRInsureContModel.ReinsurerCode = reinsurer == null ? "" : reinsurer.ReinsurerCode;
                lRInsureContModel.ReinsurerName = reinsurer == null ? "" : reinsurer.ReinsurerChineseName;
                lRInsureContModel.ChargeType = "2";
                listLRInsureContModel.Add(lRInsureContModel);
                if (!isChild)
                {
                    GetLRInsureContByRIContractInfo(model.lstChildRIContractInfo, true, i, MainReInsuranceContNo);
                }
            }
        }


        private void GetLRInsureContByHugeDisasterModel(List<HugeDisasterModel> RIContractInfo)
        {

            int reInsuranceContNo = 0;
            for (int i = 0; i < RIContractInfo.Count; i++)
            {
                LRInsureContModel lRInsureContModel = new LRInsureContModel();
                var model = RIContractInfo[i];
                var reinsurer = new Reinsurer().GetReinsurerInforByName(model.Reinsurer);
                var reinsurerCode = reinsurer == null ? string.Empty : reinsurer.ReinsurerCode;

                if (listLRInsureContModel.Where(A => A.ReInsuranceContName == model.TreatyName
                  && A.ReinsurerName == reinsurer.ReinsurerChineseName).ToList().Count() > 0)
                {
                    continue;
                }
                string currentReInsuranceContNo = "RICN" + reinsurerCode + "M" + (reInsuranceContNo++ + 1).ToString().PadLeft(2, '0') + "000";

                lRInsureContModel.TransactionNo = "";
                lRInsureContModel.CompanyCode = origanizationCode;
                lRInsureContModel.ReInsuranceContNo = currentReInsuranceContNo; //合同号码需赋值
                lRInsureContModel.ReInsuranceContName = model.TreatyName;
                lRInsureContModel.ReInsuranceContTitle = "";
                lRInsureContModel.MainReInsuranceContNo = currentReInsuranceContNo;//主合同号码 
                lRInsureContModel.ContOrAmendmentType = "合同".Equals(model.ContOrAmendmentType) ? "1" : "2";
                lRInsureContModel.ContAttribute = "1";
                lRInsureContModel.ContStatus = "ChinaRe".Equals(model.Reinsurer) && "Terminated".Equals(model.Remark) ? "2" : "1";
                lRInsureContModel.TreatyOrFacultativeFlag = "1";
                lRInsureContModel.ContSigndate = Convert.ToDateTime(model.SignDate_INSH).ToString("yyyy/MM/dd");
                lRInsureContModel.PeriodFrom = Convert.ToDateTime(model.EffectiveDate).ToString("yyyy/MM/dd");
                lRInsureContModel.PeriodTo = "ChinaRe".Equals(model.Reinsurer) && "Terminated".Equals(model.Remark) ? Convert.ToDateTime(Convert.ToDateTime(model.EffectiveDate).AddYears(1).ToString("yyyy") + "/01/01").AddDays(-1).ToString("yyyy/MM/dd") : "";
                lRInsureContModel.ContType = model.RImethodI == "非比例" ? "2" : "1";
                lRInsureContModel.ReinsurerCode = reinsurer == null ? "" : reinsurer.ReinsurerCode;
                lRInsureContModel.ReinsurerName = reinsurer == null ? "" : reinsurer.ReinsurerChineseName;
                lRInsureContModel.ChargeType = "2";
                listLRInsureContModel.Add(lRInsureContModel);
            }
        }

    }
}
