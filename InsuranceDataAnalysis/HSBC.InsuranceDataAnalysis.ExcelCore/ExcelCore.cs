using HSBC.InsuranceDataAnalysis.ExcelCore;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCore
{
    public class ExcelCore : IExcel
    {

        internal static dynamic app = null;
        internal dynamic wkb = null;
        public int UsedRowCount = 0;
        public int UsedColumnCount = 0;
        public ExcelCore()
        {
            //设置程序运行语言
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            app.DisplayAlerts = false;
            //设置是否显示Excel
            app.Visible = false;
            app.EnableEvents = false;
            //禁止刷新屏幕
            app.ScreenUpdating = false;

        }

        public void Close()
        {
            if (wkb != null)
            {
                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            app.Quit();
            IntPtr t = new IntPtr(app.Hwnd);
            int k = 0;
            Win32API.GetWindowThreadProcessId(t, out k);
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
            p.Kill();
            wkb = null;
            app = null;
            GC.Collect();
        }

        public void CreateExcel(string file)
        {
            wkb = app.Workbooks.Add(true);
            wkb.SaveAs(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void Dispose()
        {
            Close();
        }

        public Cell GetCell(int rowIndex, string columnName)
        {
            dynamic rng = app.ActiveSheet.Range(columnName + rowIndex);

            object exceldata = rng.Value(XlRangeValueDataType.xlRangeValueDefault);

            Cell cell = new Cell() { Value = exceldata == null ? "" : exceldata.ToString(), ColumnName = columnName, RowIndex = rowIndex };

            return cell;
        }

        public IList<Cell> GetRangeByName(string rangeName)
        {
            List<Cell> cells = new List<Cell>();
            dynamic rng = app.ActiveSheet.Range(rangeName);
            if (rng == null)
            {
                return cells;
            }
            if (rng.Cells.Count == 1)
            {
                Cell cell = new Cell() { Value = rng.Value(XlRangeValueDataType.xlRangeValueDefault), ColumnName = ExcelConvert.ToName((int)rng.Cells.Column - 1), RowIndex = rng.Cells.Row };
                cells.Add(cell);
                return cells;
            }
            object[,] exceldata = (object[,])rng.Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i <= exceldata.GetLongLength(0); i++)
            {
                for (int j = 1; j <= exceldata.GetLongLength(1); j++)
                {
                    Cell cell = new Cell() { Value = exceldata[i, j] == null ? "" : exceldata[i, j].ToString(), ColumnName = ExcelConvert.ToName(j - 1), RowIndex = i };
                    cells.Add(cell);
                }
            }
            return cells;
        }


        public IList<Cell> GetRange(Cell start, Cell end)
        {
            List<Cell> cells = new List<Cell>();
            dynamic rng = app.ActiveSheet.Range(start.ColumnName + start.RowIndex + ":" + end.ColumnName + end.RowIndex);


            if (rng == null)
            {
                return cells;
            }
            if (rng.Cells.Count == 1)
            {
                Cell cell = new Cell() { Value = rng.Value(XlRangeValueDataType.xlRangeValueDefault).ToString(), ColumnName = ExcelConvert.ToName((int)rng.Cells.Column - 1), RowIndex = rng.Cells.Row };
                cells.Add(cell);
                return cells;
            }
            object[,] exceldata = (object[,])rng.Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i <= exceldata.GetLongLength(0); i++)
            {
                for (int j = 1; j <= exceldata.GetLongLength(1); j++)
                {
                    Cell cell = new Cell() { Value = exceldata[i, j] == null ? "" : exceldata[i, j].ToString(), ColumnName = ExcelConvert.ToName(j - 1), RowIndex = i };
                    cells.Add(cell);
                }
            }
            return cells;
        }

        public Column GetColumn(string columnName)
        {
            Column column = new Column();

            dynamic rng = app.ActiveSheet.Range(columnName + 1 + ":" + columnName + UsedRowCount);

            if (rng == null)
            {
                return column;
            }
            if (rng.Cells.Count == 1)
            {
                Cell cell = new Cell() { Value = rng.Value(XlRangeValueDataType.xlRangeValueDefault).ToString(), ColumnName = ExcelConvert.ToName((int)rng.Cells.Column - 1), RowIndex = rng.Cells.Row };
                column.Cells.Add(cell);
                return column;
            }
            object[,] exceldata = (object[,])rng.Value(XlRangeValueDataType.xlRangeValueDefault);
            for (int i = 1; i <= exceldata.GetLongLength(0); i++)
            {
                for (int j = 1; j <= exceldata.GetLongLength(1); j++)
                {
                    Cell cell = new Cell() { Value = exceldata[i, j] == null ? "" : exceldata[i, j].ToString(), ColumnName = ExcelConvert.ToName(j - 1), RowIndex = i };
                    column.Cells.Add(cell);
                }
            }
            return column;
        }

        public IList<Row> GetSheetByRow()
        {
            List<Row> rows = new List<Row>();
            object[,] value = (object[,])app.ActiveSheet.Range["A1", ExcelConvert.ToName(app.ActiveSheet.UsedRange.Columns.Count) + app.ActiveSheet.UsedRange.Rows.Count].Value;
            for (int row = 1; row <= value.GetLongLength(0); row++)
            {
                Row rowInfo = new Row() { Index = row };
                for (int col = 1; col <= value.GetLongLength(1); col++)
                {

                    Cell cell = new Cell() { Value = value[row, col] == null ? "" : value[row, col].ToString(), ColumnName = ExcelConvert.ToName(col - 1), RowIndex = row };
                    rowInfo.Cells.Add(cell);
                }
                rows.Add(rowInfo);
            }
            return rows;
        }

        public IList<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (dynamic sheet in wkb.Worksheets())
            {
                sheetNames.Add(sheet.Name);
            }
            return sheetNames;
        }

        public void OpenExcel(string fileName, bool isReadOnly)
        {
            wkb = app.Workbooks.Open(fileName,
                  Type.Missing,
                  isReadOnly,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing,
                  Type.Missing);
            wkb.Activate();
        }

        public void Save()
        {
            wkb.Save();
        }

        public void SaveAs(string fileName)
        {
            wkb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public void SelectSheet(string sheetName)
        {
            dynamic xlsWorkSheet = wkb.Worksheets[sheetName];
            xlsWorkSheet.Select();
            UsedRowCount = xlsWorkSheet.UsedRange.Rows.Count;
            UsedColumnCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public void SelectSheet(int sheetIndex)
        {
            dynamic xlsWorkSheet = wkb.Worksheets(sheetIndex);
            xlsWorkSheet.Select();
            UsedRowCount = xlsWorkSheet.UsedRange.Rows.Count;
            UsedColumnCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public void DeleteSheet(int sheetIndex)
        {
            dynamic xlsWorkSheet = wkb.Worksheets(sheetIndex);
            xlsWorkSheet.Delete();
        }

        public void SetCellValue(int rowIndex, string columnName, string value)
        {
            app.ActiveSheet.Cells[rowIndex, ExcelConvert.ToIndex(columnName) + 1] = value;
        }
        public void SetCellValue(string sheetName, int rowIndex, string columnName, string value)
        {
            dynamic xlsWorkSheet = wkb.Worksheets[sheetName];
            xlsWorkSheet.Cells[rowIndex, ExcelConvert.ToIndex(columnName) + 1] = value;
        }
        public void AddNewSheet(string excelFilePath, string sheetName)
        {
            var workbooks = app.Workbooks;
            app.DisplayAlerts = false;
            var workbook = workbooks.Open(excelFilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            workbook.Sheets.Add(Missing.Value, workbook.Sheets[workbook.Sheets.Count], 1, Missing.Value);
            workbook.Sheets[workbook.Sheets.Count].Name = sheetName;
            workbook.Save();
            workbook.Close();
        }

        public void CloseExcel()
        {
            app.DisplayAlerts = false;
            var workbooks = app.Workbooks;
            workbooks.Close();
        }

        public void SetColumnTextType(string sheetName, int columnIndex)
        {
            SelectSheet(sheetName);
            var oRang = app.ActiveSheet.Columns(columnIndex);
            //oRang.EntireColumn.AutoFit();
            //mysheet.Cells.EntireColumn.AutoFit();
            oRang.NumberFormatLocal = "@";
        }
        public void SetColumnDateType(string sheetName, int columnIndex)
        {
            SelectSheet(sheetName);
            var oRang = app.ActiveSheet.Columns(columnIndex);
            oRang.NumberFormatLocal = @"yyyy/mm/dd";
        }
        public void SetColumnDecimalsType(string sheetName, int columnIndex)
        {
            SelectSheet(sheetName);
            var oRang = app.ActiveSheet.Columns(columnIndex);
            oRang.NumberFormatLocal = "0.00";
        }
        public void SetSheetAutoFit(string sheetName)
        {
            SelectSheet(sheetName);
            app.ActiveSheet.Cells.EntireColumn.AutoFit();
        }

        public void SetCellBackgroundColor(int rowIndex, int columnIndex, Color color)
        {
            app.ActiveSheet.Cells[rowIndex, columnIndex].Interior.Color =
                ColorTranslator.ToOle(color);
            //titleRange.Interior.Color = Color.FromArgb(224, 224, 224);//设置颜色
        }
    }
}
