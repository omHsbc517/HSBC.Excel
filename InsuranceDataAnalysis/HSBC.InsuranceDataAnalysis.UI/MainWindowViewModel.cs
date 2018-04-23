////===============================================================================
//
//  Copyright © 2018 中软国际.HSBC业务线.第二事业部.保险与卡交付部 All rights reserved    
//  
//  Filename :MainWindowViewModel
//  Description :
//
//  Created by Tina at  2/2/2018 6:05:12 PM
//
////===============================================================================
using HSBC.InsuranceDataAnalysis.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Threading.Tasks;
using System.Threading;
using System.Xml;
using HSBC.InsuranceDataAnalysis.BLL;
using HSBC.InsuranceDataAnalysis.Model;
using HSBC.InsuranceDataAnalysis.ExcelCore;
using System.Configuration;

namespace HSBC.InsuranceDataAnalysis.UI
{
    public class MainWindowViewModel : NotifyObjects
    {
        public MainWindowViewModel()
        {

            BindingProcess();
            GetConfigInformationByAppConfig();
        }

        string _uploadItem;
        public string UploadItem
        {
            get
            {
                return _uploadItem;
            }
            set
            {
                _uploadItem = value;
                RaisePropertyChanged(nameof(UploadItem));
            }
        }

        string _ReferenceFolder;
        public string ReferenceFolder
        {
            get
            {
                return _ReferenceFolder;
            }
            set
            {
                _ReferenceFolder = value;
                RaisePropertyChanged(nameof(ReferenceFolder));
            }
        }

        string _InputFolderPath;
        public string InputFolderPath
        {
            get
            {
                return _InputFolderPath;
            }
            set
            {
                _InputFolderPath = value;
                RaisePropertyChanged(nameof(InputFolderPath));
            }
        }

        string _OutputFolderPath;
        public string OutputFolderPath
        {
            get
            {
                return _OutputFolderPath;
            }
            set
            {
                _OutputFolderPath = value;
                RaisePropertyChanged(nameof(OutputFolderPath));
            }
        }

        bool _LRProductChecked = true;
        public bool LRProductChecked
        {
            get
            {
                return _LRProductChecked;
            }
            set
            {
                _LRProductChecked = value;
                RaisePropertyChanged(nameof(LRProductChecked));
                CheckCommand();
            }
        }

        bool _LRInsureContChecked = true;
        public bool LRInsureContChecked
        {
            get
            {
                return _LRInsureContChecked;
            }
            set
            {
                _LRInsureContChecked = value;
                RaisePropertyChanged(nameof(LRInsureContChecked));
                CheckCommand();
            }
        }

        bool _LRAccountChecked = true;
        public bool LRAccountChecked
        {
            get
            {
                return _LRAccountChecked;
            }
            set
            {
                _LRAccountChecked = value;
                RaisePropertyChanged(nameof(LRAccountChecked));
                CheckCommand();
            }
        }

        bool _LRContChecked = true;
        public bool LRContChecked
        {
            get
            {
                return _LRContChecked;
            }
            set
            {
                _LRContChecked = value;
                RaisePropertyChanged(nameof(LRContChecked));
                CheckCommand();
            }
        }

        bool _LREdorChecked = true;
        public bool LREdorChecked
        {
            get
            {
                return _LREdorChecked;
            }
            set
            {
                _LREdorChecked = value;
                RaisePropertyChanged(nameof(LREdorChecked));
                CheckCommand();
            }
        }

        bool _LRClaimChecked = true;
        public bool LRClaimChecked
        {
            get
            {
                return _LRClaimChecked;
            }
            set
            {
                _LRClaimChecked = value;
                RaisePropertyChanged(nameof(LRClaimChecked));
                CheckCommand();
            }
        }

        bool _CheckAll = true;
        public bool CheckAll
        {
            get
            {
                return _CheckAll;
            }
            set
            {
                _CheckAll = value;
                RaisePropertyChanged(nameof(CheckAll));
            }
        }



        static ObservableCollection<ProcessMsg> _processList = new ObservableCollection<ProcessMsg>();
        public static ObservableCollection<ProcessMsg> ProcessList
        {
            get
            {
                return _processList;
            }
            set
            {
                _processList = value;
            }
        }

        private ReplyCommand excuteCommand;

        public ReplyCommand ExcuteCommand
        {
            get
            {
                if (excuteCommand == null)
                {
                    excuteCommand = new ReplyCommand(TestCommand);
                }
                return this.excuteCommand;
            }
        }


        private ReplyCommand _CheckAllCommand;

        public ReplyCommand CheckAllCommand
        {
            get
            {
                if (_CheckAllCommand == null)
                {
                    _CheckAllCommand = new ReplyCommand(CheckAllCom);
                }
                return this._CheckAllCommand;
            }
        }

        private void CheckAllCom()
        {

            if (CheckAll)
            {
                LRAccountChecked = true;
                LRClaimChecked = true;
                LRContChecked = true;
                LREdorChecked = true;
                LRInsureContChecked = true;
                LRProductChecked = true;
            }
            else
            {
                LRAccountChecked = false;
                LRClaimChecked = false;
                LRContChecked = false;
                LREdorChecked = false;
                LRInsureContChecked = false;
                LRProductChecked = false;
            }

        }

        private void CheckCommand()
        {
            if (LRAccountChecked && LRClaimChecked && LRContChecked && LREdorChecked && LRInsureContChecked && LRProductChecked)
            {
                CheckAll = true;
            }
            else
            {
                CheckAll = false;
            }
        }

        private void TestCommand()
        {
            if (!CheckExcelPath()) return;
            Task.Factory.StartNew(() =>
            {
                CommFuns.KillExcelProcess();
                ContractInfoBusiness contractInfoBusiness = new ContractInfoBusiness();
                LRProduct lrProduct = new LRProduct();
                LRInsureCont lRInsureCont = new LRInsureCont();
                LRAccount LRAccount = new LRAccount();
                string lastDateOfMonth = string.Empty;
                SetConfigInformationFilePath();
                try
                {
                    string[] lstInputStruct = this.InputFolderPath.Trim('\\').Split('\\');

                    if (lstInputStruct.Length > 1)
                    {
                        bool isValid;
                        string rootInputPathFolderName = lstInputStruct[lstInputStruct.Length - 1];
                        lastDateOfMonth = this.GetCurrentMonthLastDay(rootInputPathFolderName, out isValid);

                        if (!isValid)
                        {
                            ProcessLogProxy.Error("The input path is illegal. The last folder format must be yyyymm or yyyymmdd!");
                            return;
                        }
                    }
                    else
                    {
                        ProcessLogProxy.Error("The input path is illegal. The last folder format must be yyyymm or yyyymmdd!");
                        return;
                    }

                    // get data source from ContractInfo
                    contractInfoBusiness.GetInformationDataFromExcel(ReferenceFolder, InputFolderPath);

                    //得到表二的数据
                    lRInsureCont.GetLRInsureContData(contractInfoBusiness);

                    //1.16.1 input data to LRProduct sheet 再保产品信息表
                    if (LRProductChecked) lrProduct.InputDataToLRProductSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth);

                    //1.16.2 LRInsureCont 生成再保合同信息表
                    if (LRInsureContChecked) lRInsureCont.WriteLRInsureContSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth);

                    //1.16.3 LRAccount 再保账单信息表 
                    if (LRAccountChecked) LRAccount.WriteLRAccountSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth);

                    //1.16.4 LRCont 再保首续期险种明细表
                    if (LRContChecked) { LRCont lRCont = new LRCont(); lRCont.WriteLRContSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth); }

                    //1.16.5 LREdor 再保保全变更信息表
                    if (LREdorChecked) { LREdor lREdor = new LREdor(); lREdor.WriteLREdorSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth); }

                    //1.16.6 LRClaim 再保理赔信息表
                    if (LRClaimChecked) { LRClaim lRClaim = new LRClaim(); lRClaim.WriteLRClaimSheet(contractInfoBusiness, OutputFolderPath, lastDateOfMonth); }

                    ProcessLogProxy.SuccessMessage("Success");

                }
                catch (Exception ex)
                {
                    ProcessLogProxy.Error(ex.Message);
                }
                finally
                {
                    CommFuns.KillExcelProcess();
                }
            });
        }


        private bool CheckExcelPath()
        {
            _processList.Clear();
            ProcessLogProxy.Normal("Start to Check excel Path");

            if (!Directory.Exists(ReferenceFolder))
            {
                ProcessLogProxy.Error("Please input Reference Folder Path");
                return false;
            }

            if (!Directory.Exists(InputFolderPath))
            {
                ProcessLogProxy.Error("Please input Input Folder Path");
                return false;
            }

            if (!Directory.Exists(OutputFolderPath))
            {
                ProcessLogProxy.Error("Please input Output File Path");
                return false;
            }
            ProcessLogProxy.SuccessMessage("Check excel Path Success");
            return true;
        }
        private void GetConfigInformationByAppConfig()
        {
            var confignManager = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConfigInformation.InformationExceFolderPath = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["ReferenceFolder"].Value) ? string.Empty : confignManager.AppSettings.Settings["ReferenceFolder"].Value.Trim();
            ConfigInformation.InputFolderPath = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["InputFolderPath"].Value) ? string.Empty : confignManager.AppSettings.Settings["InputFolderPath"].Value.Trim();
            ConfigInformation.OutputFolderPath = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["OutputFolderPath"].Value) ? string.Empty : confignManager.AppSettings.Settings["OutputFolderPath"].Value.Trim();
            ConfigInformation.NumberValue = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["NumberValue"].Value) ? string.Empty : confignManager.AppSettings.Settings["NumberValue"].Value.Trim();
            ConfigInformation.TextValue = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["TextValue"].Value) ? string.Empty : confignManager.AppSettings.Settings["TextValue"].Value.Trim();
            ConfigInformation.DateValue = string.IsNullOrWhiteSpace(confignManager.AppSettings.Settings["DateValue"].Value) ? string.Empty : confignManager.AppSettings.Settings["DateValue"].Value.Trim();
            ReferenceFolder = ConfigInformation.InformationExceFolderPath;
            InputFolderPath = ConfigInformation.InputFolderPath;
            OutputFolderPath = ConfigInformation.OutputFolderPath;
        }

        private void SetConfigInformationFilePath()
        {
            var confignManager = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            confignManager.AppSettings.Settings["ReferenceFolder"].Value = ReferenceFolder;
            confignManager.AppSettings.Settings["InputFolderPath"].Value = InputFolderPath;
            confignManager.AppSettings.Settings["OutputFolderPath"].Value = OutputFolderPath;
            confignManager.Save(ConfigurationSaveMode.Modified);
            //刷新，否则程序读取的还是之前的值（可能已装入内存）
            ConfigurationManager.RefreshSection("appSettings");
        }


        void BindingProcess()
        {
            try
            {
                var dispatcher = App.Current.Dispatcher;

                ProcessLogProxy.MessageAction = ((x) =>
                {
                    dispatcher.Invoke(() =>
                    {
                        ProcessList.Add(x);
                    });
                });
                ProcessLogProxy.SuccessMessage = ((x) =>
                {
                    dispatcher.Invoke(() =>
                    {
                        ProcessList.Add(new ProcessMsg(x, "Green"));
                    });
                });


                ProcessLogProxy.Error = (
                    (x) =>
                    {
                        dispatcher.Invoke(() =>
                        {
                            ProcessList.Add(new ProcessMsg(x, "Red"));
                        });
                    });

                ProcessLogProxy.Normal = (
                    x =>
                    {
                        dispatcher.Invoke(() =>
                        {
                            ProcessList.Add(new ProcessMsg(x, "Blue"));
                        });
                    });
            }
            catch (Exception e)
            {
                string msg = e.ToString();
                // to do nothing
            }
        }

        private string GetCurrentMonthLastDay(string yearMonthDay, out bool isValid)
        {
            string result = string.Empty;
            isValid = false;

            try
            {
                int year, month;
                year = int.Parse(yearMonthDay.Substring(0, 4));
                month = int.Parse(yearMonthDay.Substring(4, 2));

                DateTime d1 = new DateTime(year, month, 1);

                DateTime d2 = d1.AddMonths(1).AddDays(-1);
                result = d2.ToString("yyyyMMdd");
                isValid = true;
            }
            catch (Exception)
            {
                isValid = false;
            }

            return result;
        }
    }
}
