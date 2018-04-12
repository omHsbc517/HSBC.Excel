using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HSBC.InsuranceDataAnalysis.ExcelCommon.Excel;
using HSBC.InsuranceDataAnalysis.Model;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.DataValidation
{
    public class DataValidationExcelOperation
    {
        public static DataValidationExcelOperation Instance = null;
        private static object obj = new object();
        private static object objLock = new object();
        public static ExcelReflectionHelper excel = null;

        public IEnumerable<TEMP_LCCont> TEMP_LCContList { get; set; }
        public IEnumerable<TEMP_LCInsured> TEMP_LCInsuredList { get; set; }
        public IEnumerable<TEMP_LCProduct> TEMP_LCProductList { get; set; }
        public static DataValidationExcelOperation GetInstance(string path = "")
        {
            lock (obj)
            {
                try
                {
                    if (Instance == null && !string.IsNullOrEmpty(path))
                    {
                        Instance = new DataValidationExcelOperation();
                        GetExcelData(path);
                    }
                    return Instance;
                }
                catch (Exception)
                {
                    Instance = null;
                    return null;
                }
            }
        }

        public static void GetExcelData(string path)
        {
            try
            {
                var _excel = new ExcelHelper();
                excel = new ExcelReflectionHelper(false, path);
                Instance.TEMP_LCInsuredList = _excel.Read<TEMP_LCInsured>(excel);
                Instance.TEMP_LCContList = _excel.Read<TEMP_LCCont>(excel);
                Instance.TEMP_LCProductList = _excel.Read<TEMP_LCProduct>(excel);

            }
            catch (Exception ex)
            {

                throw new Exception("Get excel data error");
            }
            finally
            {
                excel.Close();
            }
        }


        public static void WriteUpdateErrorMessage(string path, WriteCellValues writeCellValues)
        {
            var _excel = new ExcelHelper();
            _excel.WriteCellValue(path, writeCellValues.SheetName, writeCellValues.RowIndex, writeCellValues.ColIndex, writeCellValues.Value);
      
        }
    }
}
