using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace MainProject
{
    public static class YandexReader
    {
        private const int DatesCount = 24;
        private const int WordsCount = 13;

        public static List<WordInfo> Read(Excel.Worksheet yandexPage)
        {
            var results = new List<WordInfo>(DatesCount);
            for (var i = 0; i < WordsCount; i++)
            {
                var skipped = 3 * i;
                var record = new WordInfo();
                record.RusWord = yandexPage.GetCellValue(4, i + 2 + skipped);
                record.EngWord = yandexPage.GetCellValue(4, i + 3 + skipped);
                for (var j = 0; j < DatesCount; j++)
                {
                    var dataRecord = new WordDateInfo();
                    dataRecord.Date = yandexPage.GetCellValue(j + 5, 1);
                    dataRecord.RusPopularity = yandexPage.GetCellValueDecimal(j + 5, i + 2 + skipped);
                    dataRecord.EngPopularity = yandexPage.GetCellValueDecimal(j + 5, i + 3 + skipped);
                    record.Dates.Add(dataRecord);
                }
                
                results.Add(record);
            }

            return results;
        }
    }
}