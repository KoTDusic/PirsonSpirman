using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
namespace MainProject
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            // Создаём экземпляр нашего приложения
            Excel.Application excelApp = new Excel.Application();
            var dir = Path.Combine(Directory.GetCurrentDirectory(), "base.xlsx");
            var workbook = excelApp.Workbooks.Open(dir);
            var yandexPage = (Excel.Worksheet)workbook.Worksheets.Item[1];
            var googlePage = (Excel.Worksheet)workbook.Worksheets.Item[2];
            
            var yandex = YandexReader.Read(yandexPage);
            var google = GoogleReader.Read(googlePage);

            excelApp = new Excel.Application();
            workbook = excelApp.Workbooks.Add();
            var newYandexPage = (Excel.Worksheet)workbook.Worksheets.Add();
            var newGooglePage = (Excel.Worksheet)workbook.Worksheets.Add();
            newGooglePage.Name = "Google trends";
            newYandexPage.Name = "Yandex wordstat";
            GoogleReader.FillPage(google, newGooglePage);
            GoogleReader.FillPage(yandex, newYandexPage);
            
            

            
 
           
            // Вычисляем сумму этих чисел
//            Excel.Range rng = workSheet.Range["A2"];      
//            rng.Formula = "=SUM(A1:L1)";
//            rng.FormulaHidden = false;
// 
//            // Выделяем границы у этой ячейки
//            Excel.Borders border = rng.Borders;
//            border.LineStyle = Excel.XlLineStyle.xlContinuous;
//       
//            // Строим круговую диаграмму
//            Excel.ChartObjects chartObjs = (Excel.ChartObjects)workSheet.ChartObjects();    
//            Excel.ChartObject chartObj = chartObjs.Add(5, 50, 300, 300);
//            Excel.Chart xlChart = chartObj.Chart;
//            Excel.Range rng2 = workSheet.Range["A1:L1"];
//            // Устанавливаем тип диаграммы
//            xlChart.ChartType = Excel.XlChartType.xlPie;
//            // Устанавливаем источник данных (значения от 1 до 10)
//            xlChart.SetSourceData(rng2);
 
            // Открываем созданный excel-файл
            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
    }
}