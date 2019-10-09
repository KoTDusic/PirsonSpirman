using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace MainProject
{
    public static class GoogleReader
    {
        private const int DatesCount = 189;
        private const int WordsCount = 11;

        public static List<WordInfo> Read(Worksheet googlePage)
        {
            var results = new List<WordInfo>(DatesCount);
            for (var i = 0; i < WordsCount; i++)
            {
                var skipped = 3 * i;
                var record = new WordInfo();
                record.RusWord = googlePage.GetCellValue(4, i + 2 + skipped);
                record.EngWord = googlePage.GetCellValue(4, i + 3 + skipped);
                for (var j = 0; j < DatesCount; j++)
                {
                    var dateRecord = new WordDateInfo();
                    dateRecord.Date = googlePage.GetCellValue(j + 5, 1);
                    dateRecord.RusPopularity = googlePage.GetCellValueDecimal(j + 5, i + 2 + skipped);
                    dateRecord.EngPopularity = googlePage.GetCellValueDecimal(j + 5, i + 3 + skipped);
                    record.Dates.Add(dateRecord);
                }

                results.Add(record);
            }

            return results;
        }

        public static void FillPage(List<WordInfo> google, Worksheet newGooglePage)
        {
            var chartObjs = (ChartObjects) newGooglePage.ChartObjects();
            newGooglePage.Cells[1, 1] = "Дата";
            var leftOffset = 0.0;
            string formula;
            for (var i = 0; i < google.Count; i++)
            {
                var currentWord = google[i];
                var skipped = 4 * i;
                var rusColumnIndex = 2 + skipped;
                var engColumnIndex = 3 + skipped;
                newGooglePage.Cells[1, rusColumnIndex] = currentWord.RusWord;
                newGooglePage.Cells[1, engColumnIndex] = currentWord.EngWord;
                newGooglePage.Cells[1, rusColumnIndex + 2] = currentWord.RusWord+" (норм.)";
                newGooglePage.Cells[1, engColumnIndex + 2] = currentWord.EngWord+" (норм.)";;
                for (var j = 0; j < currentWord.Dates.Count; j++)
                {
                    var currentRow = j + 2;

                    newGooglePage.Cells[currentRow, 1] = currentWord.Dates[j].Date;
                    newGooglePage.Cells[currentRow, rusColumnIndex] = currentWord.Dates[j].RusPopularity;
                    newGooglePage.Cells[currentRow, engColumnIndex] = currentWord.Dates[j].EngPopularity;
                    var columnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(rusColumnIndex);

                    formula =
                        $"=РАНГ.СР({columnLetter}{currentRow};{columnLetter}$2:{columnLetter}${1 + currentWord.Dates.Count};0)";
                    (newGooglePage.Cells[currentRow, rusColumnIndex + 2] as Range).FormulaLocal = formula;

                    columnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(engColumnIndex);
                    formula =
                        $"=РАНГ.СР({columnLetter}{currentRow};{columnLetter}$2:{columnLetter}${1 + currentWord.Dates.Count};0)";
                    (newGooglePage.Cells[currentRow, rusColumnIndex + 3] as Range).FormulaLocal = formula;
                }

                var rusColumnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(rusColumnIndex);
                var engColumnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(engColumnIndex);

                newGooglePage.Cells[currentWord.Dates.Count + 1 + 1, rusColumnIndex] = "Пирсон";
                formula =
                    $"=КОРРЕЛ({rusColumnLetter}2:{rusColumnLetter}{1 + currentWord.Dates.Count};{engColumnLetter}2:{engColumnLetter}{1 + currentWord.Dates.Count})";
                (newGooglePage.Cells[currentWord.Dates.Count + 1 + 1, engColumnIndex] as Range).FormulaLocal = formula;
                rusColumnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(rusColumnIndex + 2);
                engColumnLetter = WorksheetExtensions.ColumnIndexToColumnLetter(engColumnIndex + 2);

                newGooglePage.Cells[currentWord.Dates.Count + 1 + 2, rusColumnIndex] = "Спримен";
                formula =
                    $"=КОРРЕЛ({rusColumnLetter}2:{rusColumnLetter}{1 + currentWord.Dates.Count};{engColumnLetter}2:{engColumnLetter}{1 + currentWord.Dates.Count})";

                (newGooglePage.Cells[currentWord.Dates.Count + 1 + 2, engColumnIndex] as Range).FormulaLocal = formula;

                newGooglePage.Columns.AutoFit();
                if (i == 0)
                {
                    leftOffset += newGooglePage.Range[newGooglePage.Cells[1, 1],
                        newGooglePage.Cells[1, 1]].Width;
                }

                var dgramWidth = newGooglePage.Range[newGooglePage.Cells[1, rusColumnIndex],
                                     newGooglePage.Cells[1, rusColumnIndex + 3]].Width / 2;
                var dgramTopOffset = newGooglePage.Range[newGooglePage.Cells[1, 1],
                    newGooglePage.Cells[currentWord.Dates.Count + 1 + 2, 1]].Height;

                var chartRusObj = chartObjs.Add(leftOffset, dgramTopOffset, dgramWidth, 150);
                leftOffset += dgramWidth;
                var chartEngObj = chartObjs.Add(leftOffset, dgramTopOffset, dgramWidth, 150);
                leftOffset += dgramWidth;

                Chart xlChart = chartRusObj.Chart;
                var dgramRange = newGooglePage.Range[newGooglePage.Cells[2, rusColumnIndex + 2],
                    newGooglePage.Cells[currentWord.Dates.Count + 1, rusColumnIndex + 2]];
                xlChart.SetSourceData(dgramRange);
                xlChart.ChartType = XlChartType.xlXYScatter;
                xlChart.HasTitle = true;
                xlChart.ChartTitle.Text = currentWord.RusWord;

                xlChart = chartEngObj.Chart;
                dgramRange = newGooglePage.Range[newGooglePage.Cells[2, engColumnIndex + 2],
                    newGooglePage.Cells[currentWord.Dates.Count + 1, engColumnIndex + 2]];
                xlChart.SetSourceData(dgramRange);
                xlChart.ChartType = XlChartType.xlXYScatter;
                xlChart.HasTitle = true;
                xlChart.ChartTitle.Text = currentWord.EngWord;
            }

            newGooglePage.Rows.AutoFit();
        }
    }
}