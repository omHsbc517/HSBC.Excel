using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HSBC.InsuranceDataAnalysis.ExcelCommon.Excel
{
    public class ExcelTools
    {
        private static readonly string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static string MapColumnIndexToExcelCellRefrence(int index)
        {
            var x = index % 26;
            var y = index / 26;
            var refrence = alphabet[x].ToString();
            if (index >= 26)
            {
                var first = alphabet[y - 1].ToString();
                return first + refrence;
            }
            return refrence;
        }



        public static int MapExcelCellRefrenceToColumnIndex(string refrence)
        {
            var first = refrence.Select((c, i) => new { C = c, Index = i }).First(x => char.IsDigit(x.C));

            var find = refrence.Substring(0, first.Index);
            if (find.Length == 1)
            {
                return alphabet.IndexOf(find);
            }
            else
            {
                var start = find.First();
                var end = find.Last();
                var sIndex = alphabet.IndexOf(start);
                var eIndex = alphabet.IndexOf(end);

                return 26 * (sIndex + 1) + eIndex;
            }
        }


        public static int MapExcelCellRefrenceWithoutNumberToColumnIndex(string refrence) {
            var find = refrence;
            if (find.Length == 1)
            {
                return alphabet.IndexOf(find);
            }
            else {

                var start = find.First();
                var end = find.Last();
                var sIndex = alphabet.IndexOf(start);
                var eIndex = alphabet.IndexOf(end);
                return 26 * (sIndex + 1) + eIndex;
            }



   
        }
    }
}
