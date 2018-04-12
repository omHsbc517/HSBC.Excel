using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HSBC.InsuranceDataAnalysis.ExcelCommon.Excel;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public interface IExcelHelper
    {

        object[,] Read(string excelFilePath, string sheetName, string startColumn, string endColumn);

        object[,] Read(string excelFilePath, string endColumn);

        object[,] Read(string excelFilePath, string sheetName, string endColumn);

        IEnumerable<object[,]> Read(string excelFilePath, IEnumerable<ReadSheetSetting> settings);

        IEquatable<T> Read<T>(string excelFilePath, string endColumn, bool autoFixHeader, IEnumerable<ExcelMapping> map) where T : new();

        IEquatable<T> Read<T>(string excelFilePath, string endColumn, IEnumerable<ExcelMapping> map) where T : new();

        IEquatable<T> Read<T>(string excelFilePath, string endColumn, string sheetName, IEnumerable<ExcelMapping> map) where T : new();

        IEquatable<T> ReadByKey<T>(string excelFilePath, string endColumn, string key, IEnumerable<ExcelMapping> map) where T : new();

        void Write(SpreadsheetDocument doc, object[,] range, string sheetName);

        void Write(SpreadsheetDocument doc, object[,] range, string sheetName, string reportType = "", int dofultRowIndex = 2);

        void Write<T>(string excelFilePath, IEnumerable<T> entitylist, string sheetName, IEnumerable<ExcelMapping> map);

        void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entitylist, IEnumerable<ExcelMapping> map, string sheetName, string reportType);

        void Write<T>(string excelFilePath, IEnumerable<T> entitylist, IEnumerable<ExcelMapping> map, string sheetName, uint rowIndex);

        void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entitylist, IEnumerable<ExcelMapping> map, string sheetName);

        void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entitylist, IEnumerable<ExcelMapping> map, string sheetName, uint rowIndex);

        void WriteVerticalCellValues(SpreadsheetDocument doc, string sheetName, uint rowIndex, List<List<string>> allValues, string[] cols, string[] formulaCols);

        void WriteCellValue(string excelFilePath, string sheetName, List<CellInfo> cellData);

        void WriteCellValue(string excelFilePath, Dictionary<string, List<CellInfo>> cellData);

        void WriteCellValueRange(string excelFilePath, List<WriteCellValues> values);

        void FastWrite(string excelFilePath, object[,] data, string sheetName, string startCell, string endCell);

        void FastWriteToSheets(string excelFilePath, List<WriteSheetData> sheets);

        void WriteCellValue(string excelFilePath, string sheetName, int rowIndex, int colIndex, string value);

        void Create(string excelFilePath, IEnumerable<CreateSheetDate> data);

        void CalculateFormula(bool IsVisible, string Path);

        void Repair(string excelFilePath);

        Sheet GetSheet(WorkbookPart worksheetPart, string sheetName = "");

        void CalculateFormula(SpreadsheetDocument doc, List<string> sheetNames);

        void CopySheet(string sourceFilePath, string distFilePath, IEnumerable<string> sheets, bool fileFormatIsXLS = false, bool AotuFit = false);

        void CopySheetToMultipleExcel(string sourceFilePath, IEnumerable<string> distFilePaths, IEnumerable<SheetConfiguration> sheets, Dictionary<string,
            List<SpecialValue>> cells = null, Action<string, int, string> notifer = null, bool fileFormatIsXLS = false);

        void CopySheetToExcel(string sourceFilePath, Dictionary<string, List<string>> filePathAndSheets, bool fileFormatIsXLS = false);

        void RemoveColumns(string sourceFilePath, IEnumerable<string> sheets, string startColumn, int columnCount);

        void CreateEmptyExcel(string path, List<string> sheets);

        void SetRangeFormat(string filePath, IEnumerable<string> sheets, IEnumerable<NumberFormat> formats);

        void SaveToExcel97_2003(string sourcePath, string distPath);

    }
}
