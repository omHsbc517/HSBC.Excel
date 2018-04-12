using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HSBC.InsuranceDataAnalysis.ExcelCommon.DataValidation;
using HSBC.InsuranceDataAnalysis.Utils;
using HSBC.InsuranceDataAnalysis.ExcelCommon.Excel;


namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class ExcelHelper : IExcelHelper
    {
        public delegate void SetPropertyValue(string value);

        public object[,] Read(string excelFilePath, string endColumn)
        {
            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                var name = excel.GetActiveSheetName();
                var data = excel.GetData(endColumn);
                object[,] range = (object[,])data;
                return range;
            }
            catch (Exception ex)
            {
                //TODO NEED LOG
                throw new Exception("Excel interop error " + ex.Message);
            }
            finally
            {

                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }

        }

        public object[,] Read(string excelFilePath, string sheetName, string startColumn, string endColumn)
        {

            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                var name = excel.GetActiveSheetName();
                if (!string.IsNullOrEmpty(sheetName) && name != sheetName)
                {
                    excel.SelectSheetByName(sheetName);
                }

                var data = excel.GetData(startColumn, endColumn);

                object[,] range = (object[,])data;
                return range;
            }
            catch (Exception ex)
            {
                //TODO NEED LOG
                throw new Exception("Excel interop error " + ex.Message);
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }

        public IEnumerable<T> Read<T>(ExcelReflectionHelper excel) where T : new()
        {
            var attribute = typeof(T).GetCustomAttributes(true).Where(a => a.GetType() == typeof(SheetAttribute)).
                Select(a => { return a as SheetAttribute; }).FirstOrDefault(); 
            return attribute != null ? ReadSheet<T>(excel, attribute.EndColoumn, DataValidationMapper.GetMapping<T>(),
                sheetName: attribute.SheetName) : null;

        }

        public object[,] Read(string excelFilePath, string sheetName, string endColumn)
        {
            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                var name = excel.GetActiveSheetName();
                if (!string.IsNullOrEmpty(sheetName) && name != sheetName)
                {
                    excel.SelectSheetByName(sheetName);
                }
                var data = excel.GetData(endColumn);

                object[,] range = (object[,])data;
                return range;
            }
            catch (Exception ex)
            {
                //TODO NEED LOG
                throw new Exception("Excel interop error " + ex.Message);
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }

        public IEnumerable<T> ReadByKey<T>(string excelFilePath, string endColumn, string key, IEnumerable<ExcelMapping> map) where T : new()
        {
            return ReadSheet<T>(excelFilePath, endColumn, map, key);
        }

        public IEnumerable<object[,]> Read(string excelFilePath, IEnumerable<ReadSheetSetting> settings)
        {
            ExcelReflectionHelper excel = null;
            var list = new List<object[,]>();
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                var name = excel.GetActiveSheetName();

                foreach (var setting in settings)
                {
                    excel.SelectSheetByName(setting.SheetName);
                    var data = excel.GetData(setting.EndColumn);
                    object[,] range = (object[,])data;
                    list.Add(range);
                }
                return list;
            }
            catch (Exception ex)
            {
                //TODO NEED LOG
                throw new Exception("Excel interop error " + ex.Message);
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }

        public IEnumerable<T> Read<T>(ExcelReflectionHelper excel, string endColumn, string sheetName, IEnumerable<ExcelMapping> map) where T : new()
        {
            return ReadSheet<T>(excel, endColumn, map, sheetName: sheetName);
        }

        public IEnumerable<T> Read<T>(string excelFilePath, string endColumn, IEnumerable<ExcelMapping> map) where T : new()
        {
            return ReadSheet<T>(excelFilePath, endColumn, map);
        }

        public IEnumerable<T> Read<T>(string excelFilePath, bool autoFixHeader, string endColumn, IEnumerable<ExcelMapping> map) where T : new()
        {
            if (autoFixHeader)
            {
                var heaaers = map.Select((x, i) => new ColumnHeader { ColumnName = x.CoumnName, Index = i + 1 }).ToList();
                return ReadSheet<T>(excelFilePath, endColumn, map, specificHeaders: heaaers);
            }
            return ReadSheet<T>(excelFilePath, endColumn, map);
        }

        public IEnumerable<T> Read<T>(string excelFilePath, string endColumn, string sheetName, IEnumerable<ExcelMapping> map) where T : new()
        {
            return ReadSheet<T>(excelFilePath, endColumn, map, sheetName: sheetName);
        }

        public void FastWriteToSheets(string excelFilePath, List<WriteSheetData> sheets)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            foreach (var item in sheets)
            {
                var sheet = excel.SelectSheetByName(item.SheetName);
                var range = excel.GetRange(item.StartCell, item.EndCell);
                excel.SetDataToRange(range, item.Data);
            }
            excel.SaveAndClose();
            excel.CalculateAll();
        }

        public void FastWrite(string excelFilePath, object[,] data, string sheetName, string startCell, string endCell)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            var sheet = excel.SelectSheetByName(sheetName);
            var range = excel.GetRange(startCell, endCell);
            excel.SetDataToRange(range, data);
            excel.SaveAndClose();
            excel.CalculateAll();
        }

        public void Write(SpreadsheetDocument doc, object[,] range, string sheetName)
        {
            var sheet = GetSheet(doc.WorkbookPart, sheetName);
            var workSheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var sheetData = workSheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.RemoveAllChildren();
            var index = 1u;
            for (int i = 1; i <= range.GetLength(0); i++)
            {
                var row = new Row { RowIndex = (uint)i };
                for (int j = 1; j <= range.GetLength(1); j++)
                {
                    row.AppendChild(new Cell
                    {
                        CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(j - 1) + row.RowIndex,
                        DataType = CellValues.String,
                        CellValue = new CellValue(range[i, j] == null ? "" : range[i, j].ToString())
                    });
                }
                sheetData.AppendChild(row);
            }
            workSheetPart.Worksheet.Save();
            index++;
        }

        //607 231

        public void Write(SpreadsheetDocument doc, object[,] range, string sheetName, string reportType, int defaultRowIndex = 2)
        {
            var sheet = GetSheet(doc.WorkbookPart, sheetName);
            var worksheetPart = (WorksheetPart)doc.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.RemoveAllChildren();
            var index = 1u;
            for (int i = 1; i <= range.GetLength(0); i++)
            {
                var row = new Row { RowIndex = (uint)i };
                for (int j = 1; j <= range.GetLength(1); j++)
                {
                    row.AppendChild(new Cell
                    {
                        CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(j - 1) + row.RowIndex,
                        DataType = CellValues.String,
                        CellValue = new CellValue(range[i, j] == null ? "" : range[i, j].ToString())
                    });
                }
                sheetData.AppendChild(row);
            }
            worksheetPart.Worksheet.Save();
            index++;
        }

        public void Write<T>(string excelFilePath, IEnumerable<T> entityList, string sheetName, IEnumerable<ExcelMapping> map)
        {
            WriteToSheet<T>(excelFilePath, entityList, sheetName, map);
        }

        public void Write<T>(string excelFilePath, IEnumerable<T> entityList, IEnumerable<ExcelMapping> map, string sheetName, uint rowIndex)
        {
            WriteToSheet<T>(excelFilePath, entityList, sheetName, map, rowIndex);
        }

        public void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entityList, IEnumerable<ExcelMapping> map, string sheetName)
        {
            WriteToSheet<T>(doc, entityList, sheetName, map);
        }

        public void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entityList, IEnumerable<ExcelMapping> map, string sheetName, string reportType)
        {
            WriteToSheet<T>(doc, entityList, sheetName, map, reportType: reportType);
        }

        public void Write<T>(SpreadsheetDocument doc, IEnumerable<T> entityList, IEnumerable<ExcelMapping> map, string sheetName, uint rowIndex)
        {
            WriteToSheet<T>(doc, entityList, sheetName, map, rowIndex);
        }

        public void WriteVerticalCellValues<T>(SpreadsheetDocument doc, string sheetName, uint rowIndex, List<List<string>> allValues, string[] cols,
            string[] formulaCols)
        {
            var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>()
                .SingleOrDefault(n => n.Name.Value.ToLower() == sheetName.ToLower());
            if (sheet == null) throw new Exception("cannot find sheet: " + sheetName);
            WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var maxRowCount = allValues.Max(x => x.Count());
            var rows = sheetData.Descendants<Row>();
            var flag = 0;
            foreach (var values in allValues)
            {
                foreach (var value in values)
                {
                    Row row = null;
                    if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                    {
                        row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                    }
                    else
                    {
                        row = new Row { RowIndex = rowIndex };
                        sheetData.Append(row);
                    }
                    Cell cell = null;
                    string cellReference = cols[flag] + rowIndex.ToString();
                    if (row.Elements<Cell>().Where(r => r.CellReference.Value == cellReference).Count() != 0)
                    {
                        cell = row.Elements<Cell>().Where(r => r.CellReference.Value == cols[flag] + rowIndex.ToString()).First();
                    }
                    else
                    {
                        Cell refCell = null;
                        foreach (Cell ce in row.Elements<Cell>())
                        {
                            if (string.Compare(ce.CellReference.Value, cellReference, true) > 0)
                            {
                                refCell = ce;
                                break;
                            }
                        }
                        cell = new Cell { CellReference = cellReference };
                        row.InsertBefore(cell, refCell);
                    }
                    //cols[flag]=="C"||COLS[flag]=="J"||cols[flag]=="Q"
                    if (formulaCols.Contains(cols[flag]))
                    {
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString { Text = new Text(value) };
                    }
                    else
                    {
                        cell.CellFormula = new CellFormula(value) { CalculateCell = true };
                    }
                    rowIndex++;
                }
                rowIndex = 2;
                flag++;
            }
            var allRows = worksheetPart.Worksheet.Descendants<Row>();
            if (allRows.Count() > 0)
            {
                allRows.SelectMany(n => n.Elements<Cell>()).Where(n => n.CellFormula != null).ToList().ForEach(n => n.CellFormula.CalculateCell = true);
            }
            worksheetPart.Worksheet.Save();
        }

        public void Create(string excelPath, IEnumerable<CreateSheetDate> data)
        {
            var doc = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook, true);
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            var index = 1u;
            foreach (var item in data)
            {
                var worksheetPart = doc.WorkbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                var sheetData = new SheetData();
                worksheetPart.Worksheet.AppendChild(sheetData);

                var sheet = new Sheet
                {
                    Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                    Name = item.SheetName,
                    SheetId = new UInt32Value(index)
                };
                sheets.AppendChild(sheet);
                var range = item.Data;

                for (int i = 1; i < range.GetLength(0); i++)
                {
                    var row = new Row { RowIndex = (uint)i };
                    for (int j = 1; j < item.ColumnCount; j++)
                    {
                        row.AppendChild(new Cell
                        {
                            CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(j - 1) + row.RowIndex,
                            DataType = CellValues.InlineString,
                            InlineString = new InlineString { Text = new Text(range[i, j] == null ? "" : range[i, j].ToString()) }
                        });
                    }
                    sheetData.AppendChild(row);
                }
                worksheetPart.Worksheet.Save();
                index++;
            }
            doc.Close();
        }
        public void Repair(string excelFilePath)
        {
            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper();
                excel.OpenFile(excelFilePath, fix: true);
            }
            catch (Exception)
            {
                //log
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }
        public void WriteCellValue(string excelFilePath, string sheetName, int rowIndex, int colIndex, string value)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            var sheet = excel.SelectSheetByName(sheetName);
            excel.SetCellValue(rowIndex, colIndex, value);
            var value1 = excel.GetCellValue(rowIndex, colIndex);
            excel.SaveAndClose() ;
        }
        public void WriteCellValue(string excelFilePath, string sheetName, List<CellInfo> cellData)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            var sheet = excel.SelectSheetByName(sheetName);
            foreach (var cell in cellData)
            {
                int colIndex = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(cell.ColumnName) + 1;
                excel.SetCellValue(cell.RowIndex, colIndex, cell.Value);
            }
            excel.SaveAndClose();
        }
        public void WriteCellValue(string excelFilePath, Dictionary<string, List<CellInfo>> cellData)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            foreach (var cells in cellData)
            {
                excel.SelectSheetByName(cells.Key);
                foreach (var cell in cells.Value)
                {
                    int colIndex = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(cell.ColumnName) + 1;
                    excel.SetCellValue(cell.RowIndex, colIndex, cell.Value);
                }
            }
            excel.SaveAndClose();
        }

        public void WriteCellValueRange(string excelFilePath, List<WriteCellValues> values)
        {
            var excel = new ExcelReflectionHelper(false, excelFilePath);
            foreach (var item in values)
            {
                var sheet = excel.SelectSheetByName(item.SheetName);
                excel.SetCellValue(item.RowIndex, item.ColIndex, item.Value);
            }
            excel.CalculateAll();
            excel.SaveAndClose();
        }

        public Sheet GetSheet(WorkbookPart workbookPart, string sheetName = "")
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                return workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            }
            else
            {
                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>().Where(n => n.Name.Value.ToUpper() == sheetName.ToUpper());
                if (sheets.Count() > 0)
                {
                    return sheets.FirstOrDefault();
                }
                else
                {
                    return null;
                }
            }
        }


        public void CalculateFormula(SpreadsheetDocument doc, List<string> sheetNames)
        {
            foreach (var sheetName in sheetNames)
            {
                var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>()
                    .SingleOrDefault(n => n.Name.Value.ToUpper() == sheetName.ToUpper());
                if (sheet == null) throw new Exception("Cannot find sheet: " + sheetName);
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
                var allRows = worksheetPart.Worksheet.Descendants<Row>();
                if (allRows.Count() > 0)
                {
                    allRows.SelectMany(n => n.Elements<Cell>()).Where(n => n.CellFormula != null).ToList().ForEach(n => n.CellFormula.CalculateCell = true);
                }
                worksheetPart.Worksheet.Save();
            }
        }

        public void CalculateFormula(bool IsVisible, string Path)
        {
            var excel = new ExcelReflectionHelper(IsVisible, Path);
            excel.CalculateAll();
            excel.SaveAndClose();
        }
        //可能有错误
        public void CopySheet(string sourceFilePath, string distFilePath, IEnumerable<string> sheets, bool fileFormatIsXLS = false, bool AutoFit = false)
        {
            dynamic sourceExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            sourceExcel.Workbooks.Open(sourceFilePath);

            uint sourcePId;
            GetWindowThreadProcessId((IntPtr)sourceExcel.hwnd, out sourcePId);

            dynamic distExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            distExcel.DisplayAlerts = false;
            var distWorkBook = distExcel.Workbooks.Add(XlWBATemplate.XlWBATWorksheet);

            uint distPId;
            GetWindowThreadProcessId((IntPtr)distExcel.hwnd, out distPId);

            dynamic sourceWorkBook = sourceExcel.ActiveWorkbook;


            var index = 1;
            foreach (var sheetName in sheets)
            {
                var sheet = sourceWorkBook.Worksheets[sheetName];
                sheet.UsedRange.Copy(Missing.Value);

                var newSheet = index == 1 ? distWorkBook.Worksheet[index] : distWorkBook.Worksheets.Add();
                newSheet.Name = sheetName;
                newSheet.UsedRange.PasteSpecial(XlPasteType.XlPasteValues,
                    XlPasteSpecialOperation.XlPasteSpecialOperationNone,
                    Missing.Value, Missing.Value);
                if (AutoFit)
                    newSheet.Columns.AutoFit();
                index++;
            }

            sourceWorkBook.Save();
            sourceWorkBook.Close();
            sourceExcel.Quit();
            if (fileFormatIsXLS)
                distWorkBook.SaveAs(distFilePath, 1);
            else
                distWorkBook.SaveAs(distFilePath);

            distWorkBook.Close();
            distWorkBook.Quit();


            Marshal.ReleaseComObject(sourceExcel);
            Marshal.ReleaseComObject(distExcel);

            GC.Collect();
            GC.WaitForPendingFinalizers();

            KillExcelProcess((int)sourcePId);
            KillExcelProcess((int)distPId);
        }
        public void CopySheetToMultipleExcel(string sourceFilePath, IEnumerable<string> distFilePaths, IEnumerable<SheetConfiguration> sheets, Dictionary<string,
              List<SpecialValue>> cells = null, Action<string, int, string> notifer = null, bool fileFormatIsXLS = false)
        {
            dynamic sourceExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            sourceExcel.Workbooks.Open(sourceFilePath);
            dynamic sourceWorkBook = sourceExcel.ActiveWorkbook;

            uint sourcePId;
            GetWindowThreadProcessId((IntPtr)sourceExcel.hwnd, out sourcePId);

            foreach (var distFilePath in distFilePaths)
            {
                if (cells != null)
                {
                    var values = cells[distFilePath];

                    foreach (var item in values)
                    {
                        var sheet = sourceWorkBook.Worksheets[item.SheetName];
                        sheet.Cells[item.RowIndex, item.ColIndex].value = item.Value;
                    }
                }

                dynamic distExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                distExcel.DisplayAlerts = false;
                var distWorkBook = distExcel.Workbooks.Add(XlWBATemplate.XlWBATWorksheet);

                uint distPId;
                GetWindowThreadProcessId((IntPtr)distExcel.hwnd, out distPId);

                var index = 1;
                foreach (var sheet in sheets)
                {
                    dynamic currentSheet = sourceWorkBook.Worksheet[sheet.SheetName];
                    //可能会错误
                    var range = sheet.UseSpecialRange ?
                        currentSheet.Range(string.Format(sheet.Range, currentSheet.UsedRange.Rows.Count)) :
                        currentSheet.UsedRange;
                    range.Copy(Missing.Value);

                    var newSheet = index == 1 ? distWorkBook.Worksheet[index] : distWorkBook.Worksheets.Add();
                    newSheet.Name = sheet.SheetName;
                    newSheet.UsedRange.PasteSpecial(XlPasteType.XlPasteValues,
                       XlPasteSpecialOperation.XlPasteSpecialOperationNone,
                       Missing.Value, Missing.Value);
                    index++;
                }
                if (fileFormatIsXLS)
                    distWorkBook.SaveAs(distFilePath, 1);
                else
                    distWorkBook.SaveAs(distFilePath);
                distWorkBook.Close();
                distExcel.Quit();
                Marshal.ReleaseComObject(distExcel);
                KillExcelProcess((int)distPId);
            }

            sourceWorkBook.Save();
            sourceWorkBook.Close();
            sourceExcel.Quit();
            Marshal.ReleaseComObject(sourceExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillExcelProcess((int)sourcePId);
        }

        public void CopySheetToExcel(string sourceFilePath, Dictionary<string, List<string>> filePathAndSheets, bool fileFormatIsXLS = false)
        {
            dynamic sourceExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            sourceExcel.Workbooks.Open(sourceFilePath);
            dynamic sourceWorkBook = sourceExcel.ActiveWorkbook;

            uint sourcePId;
            GetWindowThreadProcessId((IntPtr)sourceExcel.hwnd, out sourcePId);

            foreach (var item in filePathAndSheets)
            {
                dynamic distExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                distExcel.DisplayAlerts = false;
                var distWorkBook = distExcel.Workbooks.Add(XlWBATemplate.XlWBATWorksheet);

                uint distPId;
                GetWindowThreadProcessId((IntPtr)distExcel.hwnd, out distPId);

                var index = 1;
                foreach (var sheetName in item.Value)
                {
                    var sheet = sourceWorkBook.Worksheets[sheetName];
                    sheet.UsedRange.Copy(Missing.Value);

                    var newSheet = index == 1 ? distWorkBook.Worksheet[index] : distWorkBook.Worksheets.Add();
                    newSheet.Name = sheetName;
                    newSheet.UsedRange.PasteSpecial(XlPasteType.XlPasteValues,
                        XlPasteSpecialOperation.XlPasteSpecialOperationNone,
                        Missing.Value, Missing.Value);
                    index++;
                }
                if (fileFormatIsXLS)
                    distWorkBook.SaveAs(item.Key, 1);
                else
                    distWorkBook.SaveAs(item.Key);
                distWorkBook.Close();//distWorkBook.Save();//可能会有问题
                distExcel.Quit();
                Marshal.ReleaseComObject(distExcel);
                KillExcelProcess((int)distPId);
            }

            sourceWorkBook.Save();
            sourceWorkBook.Close();
            sourceExcel.Quit();
            Marshal.ReleaseComObject(sourceExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillExcelProcess((int)sourcePId);
        }

        public void RemoveColumns(string sourceFilePath, IEnumerable<string> sheets, string startColumn, int columnCount)
        {
            dynamic sourceExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            sourceExcel.DisplayAlerts = false;
            sourceExcel.Workbooks.Open(sourceFilePath);
            dynamic sourceWorkBook = sourceExcel.ActiveWorkbook;

            sourceExcel.DisplayAlerts = false;
            uint sourcePId;
            GetWindowThreadProcessId((IntPtr)sourceExcel.hwnd, out sourcePId);

            foreach (var sheetName in sheets)
            {
                var sheet = sourceWorkBook.Worksheets[sheetName];
                for (int i = 0; i < columnCount; i++)
                {
                    sheet.Range[startColumn + "1", Missing.Value].EntireColumn.Delete(Missing.Value);
                }
            }

            sourceWorkBook.Save();
            sourceWorkBook.Close();
            sourceExcel.Quit();
            Marshal.ReleaseComObject(sourceExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillExcelProcess((int)sourcePId);
        }

        public void CreateEmptyExcel(string path, List<string> sheets)
        {
            dynamic excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            var workBook = excel.Workbooks.Add(XlWBATemplate.XlWBATWorksheet);
            
            uint pId;
            GetWindowThreadProcessId((IntPtr)excel.hwnd, out pId);

            foreach (var sheetName in sheets)
            {
                var sheet = workBook.Worksheets.Add();
                sheet.Name = sheetName;
            }

            workBook.SaveAs(path);
            workBook.Close();
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillExcelProcess((int)pId);
        }

        public void SetRangeFormat(string filePath, IEnumerable<string> sheets, IEnumerable<NumberFormat> formats)
        {
            var excel = new ExcelReflectionHelper(false, filePath);
            excel.SetRangeFormat(sheets, formats);
            excel.SaveAndClose();
        }

        public void SaveToExcel97_2003(string sourcePath, string distPath)
        {
            dynamic sourceExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            sourceExcel.Workbooks.Open(sourceExcel);
            dynamic sourceWorkBook = sourceExcel.ActiveWorkbook;

            uint sourcePId;
            GetWindowThreadProcessId((IntPtr)sourceExcel.hwnd, out sourcePId);

            sourceWorkBook.SaveAs(distPath, 1);
            sourceExcel.Quit();

            Marshal.ReleaseComObject(sourceExcel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillExcelProcess((int)sourcePId);
        }

        private void WriteToSheet<T>(string excelFilePath, IEnumerable<T> entityList, string sheetName, IEnumerable<ExcelMapping> map, uint rowIndex = 2)
        {
            var doc = SpreadsheetDocument.Open(excelFilePath, true, new OpenSettings { AutoSave = true });
            WriteToSheet(doc, entityList, sheetName, map, rowIndex);
            doc.Close();
        }

        private void WriteToSheet<T>(SpreadsheetDocument doc, IEnumerable<T> entityList, string sheetName,
            IEnumerable<ExcelMapping> map, uint rowIndex = 2, string reportType = "")
        {
            if (entityList.Count() == 0) return;
            var sheet = GetSheet(doc.WorkbookPart, sheetName);
            var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var shareTable = doc.WorkbookPart.SharedStringTablePart.SharedStringTable;
            var rows = sheetData.Descendants<Row>();
            var properties = typeof(T).GetProperties();

            var headers = rows.First().Descendants<Cell>().Where(f => f.CellValue != null).Select((c, i) => new ColumnHeader
            {
                CellRefrence = c.CellReference,
                Index = i,
                ColumnName = (c.DataType != null && c.DataType == CellValues.SharedString) ?
                shareTable.ElementAt(int.Parse(c.CellValue.Text)).InnerText :
                c.CellValue.Text
            });
            var cellCount = headers.Max(h => h.Index);
            int rowCount = rows.Count();
            for (int i = (int)rowIndex; i < rowCount + 1; i++)
            {
                sheetData.RemoveChild(rows.ElementAt((int)rowIndex));
            }
            foreach (var entity in entityList)
            {
                Row newRow = new Row { RowIndex = rowIndex };
                for (int i = 0; i < cellCount + 1; i++)
                {
                    newRow.AppendChild(new Cell
                    {
                        CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(i) + rowIndex
                    });
                }
                sheetData.AppendChild(newRow);
                var cells = newRow.Descendants<Cell>();
                foreach (var property in properties)
                {
                    var matched = map.Single(m => m.PropertyName == property.Name);
                    var colName = matched.CoumnName;

                    var refrence = headers.Single(h => h.ColumnName.ToLower() == colName.ToLower() ||
                        h.ColumnName.Replace("_", " ").ToLower() == colName.Replace("_", " ").ToLower()).CellRefrence;
                    var colIndex = ExcelTools.MapExcelCellRefrenceToColumnIndex(refrence);
                    var cell = cells.ElementAt(colIndex);
                    var value = property.GetValue(entity, null) == null ? "" : property.GetValue(entity, null).ToString();
                    double v;
                    if (double.TryParse(value, out v))
                    {
                        cell.DataType = CellValues.Number;
                        cell.CellValue = new CellValue(value);
                    }
                    else
                    {
                        cell.DataType = CellValues.InlineString;
                        cell.InlineString = new InlineString { Text = new Text(value) };
                    }
                }
                rowIndex++;
            }
            worksheetPart.Worksheet.Save();
        }

        public IEnumerable<T> ReadSheet<T>(ExcelReflectionHelper excel, string endColumn, IEnumerable<ExcelMapping> map, string sheetName = "",
            List<ColumnHeader> specificHeaders = null) where T : new()
        {
            try
            {
                var name = excel.GetActiveSheetName();
                if (!string.IsNullOrEmpty(sheetName) && name != sheetName)
                {
                    excel.SelectSheetByName(sheetName);
                }
                var list = new List<T>();
                var headers = new List<ColumnHeader>();
                if (specificHeaders != null)
                {
                    headers = specificHeaders;
                }

                var data = excel.GetData(endColumn);
                object[,] range = (object[,])data;
                for (int i = 1; i <= range.GetLength(0); i++)
                {
                    var entity = new T();
                    var valueSetter = CreateSetPropertyValueDelegate(entity);
                    var end = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(endColumn) + 1;

                    for (int j = 1; j <= end; j++)
                    {
                        if (i == 1)
                        {
                            if (specificHeaders == null)
                            {
                                var match = map.SingleOrDefault(m => m.CoumnName.ToLower() == range[i, j].ToString().ToLower());
                                if (match != null)
                                {
                                    headers.Add(new ColumnHeader
                                        {
                                            ColumnName = match.CoumnName,
                                            Index = j
                                        });
                                }
                            }
                        }
                        else
                        {
                            var match = headers.SingleOrDefault(h => h.Index == j);
                            if (match != null)
                            {
                                var value = range[i, j];
                                valueSetter[map.Single(m => m.CoumnName.ToLower() == match.ColumnName.ToLower()).PropertyName](value == null ? "" :
                                    value.ToString().Trim());
                            }
                        }
                    }
                    if (i > 1)
                    {
                        list.Add(entity);
                    }
                }
                return list;
            }
            catch (Exception e)
            {

                throw new Exception(string.Format("Read {0} sheet data error", sheetName));
            }
        }

        private IEnumerable<T> ReadSheet<T>(string excelFilePath, string endColumn, IEnumerable<ExcelMapping> map, string sheetName = "",
            List<ColumnHeader> specificHeaders = null) where T : new()
        {
            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                var name = excel.GetActiveSheetName();
                if (!string.IsNullOrEmpty(sheetName) && name != sheetName)
                {
                    excel.SelectSheetByName(sheetName);
                }
                var list = new List<T>();
                var headers = new List<ColumnHeader>();
                if (specificHeaders != null)
                {
                    headers = specificHeaders;
                }
                var data = excel.GetData(endColumn);
                object[,] range = (object[,])data;
                for (int i = 0; i <= range.GetLength(0); i++)
                {
                    var entity = new T();
                    var valueSetter = CreateSetPropertyValueDelegate(entity);
                    var end = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(endColumn) + 1;
                    for (int j = 1; j <= end; j++)
                    {
                        if (i == 1)
                        {
                            if (specificHeaders == null)
                            {
                                var match = map.SingleOrDefault(m => m.CoumnName.ToLower() == range[i, j].ToString().ToLower());
                                if (match != null)
                                {
                                    headers.Add(new ColumnHeader
                                    {
                                        ColumnName = match.CoumnName,
                                        Index = j
                                    });
                                }
                            }
                        }
                        else
                        {
                            var match = headers.SingleOrDefault(h => h.Index == j);
                            if (match != null)
                            {
                                var value = range[i, j];
                                valueSetter[map.Single(m => m.CoumnName.ToLower() == match.ColumnName.ToLower()).PropertyName](value == null ? "" :
                                    value.ToString().Trim());
                            }
                        }
                    }
                    if (i > 1)
                    {
                        list.Add(entity);
                    }
                }

                return list;
            }
            catch (Exception)
            {
                //

                throw new Exception("excel interop error");
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }

        private IEnumerable<T> ReadSheet<T>(string excelFilePath, string endColumn, IEnumerable<ExcelMapping> map, string key) where T : new()
        {
            ExcelReflectionHelper excel = null;
            try
            {
                excel = new ExcelReflectionHelper(false, excelFilePath);
                List<string> sheetNames = excel.SelectSheetNameByKey(key);
                var list = new List<T>();
                foreach (var name in sheetNames)
                {
                    excel.SelectSheetByName(name);
                    var headers = new List<ColumnHeader>();
                    var data = excel.GetData(endColumn);
                    object[,] range = (object[,])data;

                    for (int i = 1; i <= range.GetLength(0); i++)
                    {
                        var entity = new T();
                        var valueSetter = CreateSetPropertyValueDelegate(entity);
                        var end = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(endColumn) + 1;
                        for (int j = 1; j <= end; j++)
                        {
                            if (i == 1)
                            {
                                var match = map.SingleOrDefault(m => m.CoumnName.ToLower() == range[i, j].ToString().ToLower());
                                if (match != null)
                                {
                                    headers.Add(new ColumnHeader
                                    {
                                        ColumnName = match.CoumnName,
                                        Index = j
                                    });
                                }
                            }
                            else
                            {
                                var match = headers.SingleOrDefault(h => h.Index == j);
                                if (match != null)
                                {
                                    var value = range[i, j];
                                    valueSetter[map.Single(m => m.CoumnName.ToLower() == match.ColumnName.ToLower()).PropertyName](value == null ? "" :
                                        value.ToString());
                                }
                            }
                        }
                        if (i > 1)
                        {
                            list.Add(entity);
                        }
                    }
                }
                return list;
            }
            catch (Exception)
            {
                //

                throw new Exception("excel interop error");
            }
            finally
            {
                if (excel != null)
                {
                    excel.SaveAndClose();
                }
            }
        }
        private IEnumerable<T> ReadSheet<T>(ExcelReflectionHelper excel, ReadSheetSetting setting, IEnumerable<ExcelMapping> map, string key) where T : new()
        {
            excel.SelectSheetByName(setting.SheetName);
            var data = excel.GetData(setting.EndColumn);
            object[,] range = (object[,])data;
            var list = new List<T>();
            var headers = new List<ColumnHeader>();
            for (int i = 1; i <= range.GetLength(0); i++)
            {
                var entity = new T();
                var valueSetter = CreateSetPropertyValueDelegate(entity);
                var end = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(setting.EndColumn) + 1;
                for (int j = 1; j <= end; j++)
                {
                    if (i == 1)
                    {
                        var match = map.SingleOrDefault(m => m.CoumnName.ToLower() == range[i, j].ToString().ToLower());
                        if (match != null)
                        {
                            headers.Add(new ColumnHeader
                            {
                                ColumnName = match.CoumnName,
                                Index = j
                            });
                        }
                    }
                    else
                    {
                        var match = headers.SingleOrDefault(h => h.Index == j);
                        if (match != null)
                        {
                            var value = range[i, j];
                            valueSetter[map.Single(m => m.CoumnName.ToLower() == match.ColumnName.ToLower()).PropertyName](value == null ? "" : value.ToString());
                        }
                    }
                }
                if (i > 1)
                {
                    list.Add(entity);
                }
            }
            return list;
        }

        private Dictionary<string, SetPropertyValue> CreateSetPropertyValueDelegate<T>(T t)
        {
            var type = typeof(T);
            var properties = type.GetProperties();
            var dic = new Dictionary<string, SetPropertyValue>();
            foreach (var pi in properties)
            {
                var key = pi.Name;
                var value = (SetPropertyValue)Delegate.CreateDelegate(typeof(SetPropertyValue), t, type.GetProperty(pi.Name).GetSetMethod());
                dic.Add(key, value);
            }
            return dic;
        }

        private void KillExcelProcess(int pid)
        {
            try
            {
                System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid);
                if (process != null)
                {
                    process.Kill();
                }
            }
            catch (Exception)
            {
            }
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint ProcessId);

        public void Write(SpreadsheetDocument doc, object[,] range, string sheetName, string columnName)
        {
            int col = ExcelTools.MapExcelCellRefrenceWithoutNumberToColumnIndex(columnName) + 1;
            var sheet = GetSheet(doc.WorkbookPart, sheetName);
            var worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.RemoveAllChildren();
            var index = 1u;
            try
            {
                for (int i = 1; i <= range.GetLength(0); i++)
                {
                    var row = new Row { RowIndex = (uint)i };
                    for (int j = 1; j <= range.GetLength(1); j++)
                    {
                        if (j != col)
                        {
                            row.AppendChild(new Cell
                            {
                                CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(j - 1) + row.RowIndex,
                                DataType = CellValues.String,
                                CellValue = new CellValue(range[i, j] == null ? "" : range[i, j].ToString())
                            });
                        }
                        else
                        {
                            row.AppendChild(new Cell
                            {
                                CellReference = ExcelTools.MapColumnIndexToExcelCellRefrence(j - 1) + row.RowIndex,
                                DataType = CellValues.String,
                                CellValue = new CellValue(range[i, j].ToString() == null ? "" : (Convert.ToDouble(range[i, j]) * 100).ToString() + "%")
                            });
                        }
                    }
                    sheetData.AppendChild(row);
                }
                worksheetPart.Worksheet.Save();
                index++;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void KillExcelProcess()
        {
            try
            {
                var listProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL").ToList();
                foreach (var item in listProcesses)
                {
                    item.Kill();
                }

            }
            catch (Exception)
            {
            }
        }


        public IEquatable<T> Read<T>(string excelFilePath, string endColumn, bool autoFixHeader, IEnumerable<ExcelMapping> map) where T : new()
        {
            throw new NotImplementedException();
        }

        IEquatable<T> IExcelHelper.Read<T>(string excelFilePath, string endColumn, IEnumerable<ExcelMapping> map)
        {
            throw new NotImplementedException();
        }

        IEquatable<T> IExcelHelper.Read<T>(string excelFilePath, string endColumn, string sheetName, IEnumerable<ExcelMapping> map)
        {
            throw new NotImplementedException();
        }

        IEquatable<T> IExcelHelper.ReadByKey<T>(string excelFilePath, string endColumn, string key, IEnumerable<ExcelMapping> map)
        {
            throw new NotImplementedException();
        }

        public void WriteVerticalCellValues(SpreadsheetDocument doc, string sheetName, uint rowIndex, List<List<string>> allValues, string[] cols, string[] formulaCols)
        {
            throw new NotImplementedException();
        }
    }
}
