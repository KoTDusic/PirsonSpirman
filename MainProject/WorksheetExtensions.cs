using System;
using System.Globalization;
using Microsoft.Office.Interop.Excel;

namespace MainProject
{
    public static class WorksheetExtensions
    {
        public static string GetCellValue(this Worksheet worksheet, int row, int column)
        {
            return (worksheet.Cells[row, column] as Range)?.Value.ToString();
        }
        public static double GetCellValueDecimal(this Worksheet worksheet, int row, int column)
        {
            var cellValue = (worksheet.Cells[row, column] as Range).Text;
            return double.Parse(cellValue, NumberStyles.Any, CultureInfo.CurrentCulture);
        }
        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;
 
            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }
}