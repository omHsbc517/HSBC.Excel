using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCore
{
    public interface IExcel : IDisposable
    {
        void CreateExcel(string file);
        void OpenExcel(string fileName, bool isReadOnly);
        void SelectSheet(string sheetName);
        void SelectSheet(int sheetIndex);
        IList<string> GetSheetNames();
        IList<Row> GetSheetByRow();
        Column GetColumn(string columnName);
        IList<Cell> GetRangeByName(string rangeName);
        IList<Cell> GetRange(Cell start, Cell end);
        void SetCellValue(int rowIndex, string columnName, string value);
        void SetCellValue(string sheetName, int rowIndex, string columnName, string value);
        Cell GetCell(int rowIndex, string columnName);
        void Close();
        void Save();
        void SaveAs(string fileName);
        void DeleteSheet(int sheetIndex);
        void AddNewSheet(string filePath, string sheetName);
        void CloseExcel();
        void SetColumnTextType(string sheetName, int columnIndex);
        void SetSheetAutoFit(string sheetName);
        void SetColumnDateType(string sheetName, int columnIndex);
        void SetColumnDecimalsType(string sheetName, int columnIndex);
        void SetCellBackgroundColor(int rowIndex, int columnIndex, Color color);
    }
}
