using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class ExcelReflectionHelper
    {
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        public uint ProcessID { set; get; }

        public dynamic xlApp { set; get; }

        public string Path { set; get; }

        public ExcelReflectionHelper()
        {
            xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            xlApp.Visible = false;
            xlApp.ScreenUpdating = false;
        }

        public ExcelReflectionHelper(bool IsVisible = true, string path = "", string Password = "")
        {
            xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            xlApp.Visible = IsVisible;
            xlApp.ScreenUpdating = false;
            xlApp.DisplayAlerts = false;

            int xlAppHwnd = (int)xlApp.Hwnd;
            uint processId;
            GetWindowThreadProcessId((IntPtr)xlAppHwnd, out processId);
            ExcelProcess.AddProcessId(processId);
            ProcessID = processId;

            if (!string.IsNullOrEmpty(path))
            {
                if (!string.IsNullOrEmpty(Password))
                {
                    OpenFile(path, Password);
                }
                else
                {
                    OpenFile(path);
                }
                this.Path = path;
            }
            else
            {
                AddSheet();
            }

        }

        public void AddSheet()
        {
            // AddSheet(1);//aaaa
        }

        public void OpenFile(string filePath, string password = "", bool? fix = null)
        {
            dynamic wkBook = xlApp.workbooks;
            object pw = Type.Missing;
            object fixValue = Type.Missing;

            if (!string.IsNullOrEmpty(password))
            {
                pw = password;
            }

            if (fix.HasValue && fix.Value)
            {
                fixValue = fix.Value;
            }

            wkBook.Open(filePath, Type.Missing, false,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, fixValue);

           
        }

        public void OpenFile(string FilePath, string Password)
        {
            dynamic wkBook = xlApp.workbooks;
            wkBook.Open(FilePath,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Password,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               Type.Missing,
               true);
        }

        public void MakeVisible()
        {
            xlApp.Visible = true;
        }

        public void Hide()
        {
            xlApp.Visible = false;
        }

        public void KillThis()
        {
            if (ProcessID > 0)
            {
                try
                {
                    Process process = Process.GetProcessById((int)ProcessID);
                    if (process != null)
                    {
                        process.Kill();
                        ProcessID = 0;
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        public void SaveAs(string FilePath)
        {
            xlApp.ActiveWorkbook.SaveAs(FilePath, 51, false);
        }

        public void CalculateAll()
        {
            xlApp.CalculateFull();
        }

        public void Save()
        {
            xlApp.ActiveWorkbook.acWkBook.Save();
        }

        public void SaveAndClose()
        {
            dynamic acWkBook = xlApp.ActiveWorkbook;

            if (acWkBook != null)
            {
                try
                {
                    acWkBook.CheckCompatibility = false;
                }
                catch (Exception)
                {
                }
                acWkBook.Save();
                acWkBook.Close(true);
                xlApp.Quit();
                Marshal.ReleaseComObject((object)acWkBook);
                Marshal.ReleaseComObject((object)xlApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
            KillThis();
        }

        public void Close()
        {
            dynamic acWkBook = xlApp.ActiveWorkbook;

            if (acWkBook != null)
            {
                try
                {
                    acWkBook.CheckCompatibility = false;
                }
                catch (Exception)
                {
                }
                acWkBook.Close(true);
                xlApp.Quit();
                Marshal.ReleaseComObject((object)acWkBook);
                Marshal.ReleaseComObject((object)xlApp);
            }
            KillThis();
        }

        public void SetCellValue(int RowIndex, int ColIndex, object value)
        {
            dynamic acSheet = xlApp.ActiveSheet;
        
            if (acSheet != null)
            {
                //dynamic cell = acSheet.Cells(RowIndex, ColIndex);
                //cell.value = value;
                acSheet.Cells[RowIndex, ColIndex] = value;
            }

        }

        public object GetCellValue(int RowIndex, int ColIndex)
        {
            dynamic acSheet = xlApp.ActiveSheet;

            if (acSheet != null)
            {
                dynamic cell = acSheet.Cells(RowIndex, ColIndex);
                return cell.value;
            }
            return null;
        }

        public object SelectSheetByName(string SheetName)
        {
            dynamic sheet = xlApp.Sheets[SheetName];
            sheet.Select();
            return sheet;
        }

        public string GetActiveSheetName()
        {

            dynamic sheet = xlApp.ActiveSheet;
            if (sheet != null)
            {
                return Convert.ToString(sheet.Name);
            }
            else
            {
                return "";
            }
        }

        public int GetSheetCount()
        {
            return xlApp.ActiveWorkbook.Worksheets.Count;
        }

        public List<string> SelectSheetNameByKey(string key)
        {
            dynamic sheets = xlApp.ActiveWorkbook.Worksheets;

            List<string> names = new List<string>();

            for (int i = 1; i <= sheets.Count; i++)
            {
                if (sheets[i].Name.ToLower().Contains(key.ToLower()))
                {
                    names.Add(sheets[i].Name);
                }
            }
            return names;
        }

        public void SelectCell(string RangeText)
        {
            dynamic range = xlApp.Range(RangeText);
            range.Select();
        }

        public void UnlockActiveSheet(string Password)
        {
            dynamic acSheet = xlApp.ActiveSheet;
            if (acSheet != null)
            {
                acSheet.Unlock(Password);
            }
        }

        public void SetRangeFormat(IEnumerable<string> sheets, IEnumerable<NumberFormat> format)
        {
            foreach (var sheetName in sheets)
            {
                SelectSheetByName(sheetName);
                foreach (var col in format)
                {
                    dynamic range = xlApp.Range(col.RangeText);
                    range.NumberFormat = col.Format;
                }

            }
        }

        internal enum XlAutoFilterOperator
        {
            xlAnd = 1,
            xlBottom10Items = 4,
            xlBottom10Percent = 6,
            xlFilterCellColor = 8,
            xlFilterDynamic = 11,
            xlFilterFontColor = 9,
            xlFilterIcon = 10,
            xlFilterValues = 7,
            xlOr = 2,
            xlTop10Items = 3,
            xlTop10Percent = 5

        }

        internal enum XlDirection
        {
            xlDown = -4121,
            xlToLeft = -4159,
            xlToRight = -4161,
            xlUp = -4162
        }

        public int GetRowCount()
        {
            dynamic range = xlApp.Range("A:A");
            dynamic func = xlApp.WorksheetFunction;
            dynamic count = func.CountA(range);
            return int.Parse(count.ToString());
        }

        public int GetAllRowCount()
        {
            dynamic sheet = xlApp.ActiveSheet;
            dynamic range = sheet.UsedRange;
            dynamic rows = range.Rows;
            dynamic count = rows.Count;
            return int.Parse(count.ToString());
        }

        public object GetData(string endColumnName)
        {
            var rowCount = GetAllRowCount();
            if (rowCount == 0) return null;
            dynamic range = xlApp.Range(string.Format("A1:{0}{1}", endColumnName, rowCount));
            dynamic data = range.Value2;
            return data;
        }

        public object GetData(string startColumnName, string endColumnName)
        {
            var rowCount = GetAllRowCount();
            if (rowCount == 0) return null;
            dynamic range = xlApp.Range(string.Format(string.Format("{0}:{1}", startColumnName, endColumnName)));
            dynamic data = range.Value2;
            return data;
        }

        public object GetRange(string start, string end)
        {
            dynamic range = xlApp.Range(string.Format("{0}:{1}", start, end));
            return range;
        }

        public void SetDataToRange(object range, object[,] data)
        {
            dynamic Range = range;
            Range.Value2 = data;
        }
    }
    public enum XlWBATemplate
    {
        XlWBATWorksheet = -4167,
        XlWBATChart = -4109,
        XlWBATExcel4MacroSheet = 3,
        XlWBATExcel4IntMacroSheet = 4
    }
    public enum XlPasteType
    {
        XlPasteValues = -4163,
        XlPasteComments = -4144,
        XlPasteFormulas = -4123,
        XlPasteFormats = -4122,
        XlPasteAll = -4104,
        XlPasteValidation = 6,
        XlPasteAllExceptBorders = 7,
        XlPasteColumnWidths = 8,
        XlPasteFormulasAndNumberFormats = 11,
        XlPasteValuesAndNumberFormats = 12,
        XlPasteAllUsingSourceTheme = 13,
        XlPasteAllMergingConditionalFormats = 14
    }

    public enum XlPasteSpecialOperation
    {
        XlPasteSpecialOperationNone = -4142,
        XlPasteSpecialOperationAdd = 2,
        XlPasteSpecialOperationSubtract = 3,
        XlPasteSpecialOperationMultiply = 4,
        XlPasteSpecialOperationDivide = 5
    }

}

