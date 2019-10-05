using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace ExcelChart
{
    class Program
    {
        static void Main(string[] args)
        {
            string tableRange = "A1:C10";
            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet1");
            ExcelTable table = sheet.Tables.Add(sheet.Cells[tableRange], "Example");
            table.Columns[0].Name = "Hour";
            table.Columns[1].Name = "Sales October 1st";
            table.Columns[2].Name = "Sales October 2nd";

            //fill hour column
            int firstDataRowIndex = 2; // first row is for the column title
            int startHour = 9;
            int endHour = 17;
            DateTime time = new DateTime(2019, 10, 3, startHour, 0, 0);
            while (time.Hour <= endHour)
            {
                int rowIndex = time.Hour - startHour + firstDataRowIndex;
                sheet.Cells["A" + rowIndex.ToString()].Value = time.ToString("HH:mm:ss");
                time = time.AddHours(1);
            }

            //fill first data column
            foreach (var x in Enumerable.Range(0, endHour-startHour + 1))
            {
                int rowIndex = x + firstDataRowIndex;
                sheet.Cells["B" + rowIndex.ToString()].Value = x % 2 + 1;
            }

            //fill second data column
            foreach (var x in Enumerable.Range(0, endHour - startHour + 1))
            {
                int rowIndex = x + firstDataRowIndex;
                sheet.Cells["C" + rowIndex.ToString()].Value = x % 3 + 2;
            }

            //create chart
            OfficeOpenXml.Drawing.Chart.ExcelChart chart = sheet.Drawings.AddChart("example", eChartType.ColumnClustered);
            chart.XAxis.Title.Text = "Hour";
            chart.XAxis.Title.Font.Size = 10;
            chart.YAxis.Title.Text = "Sales";
            chart.YAxis.Title.Font.Size = 10;
            chart.SetSize(500, 300);
            chart.SetPosition(0, 0, 4, 0);

            //add chart series
            string xValuesRange = "A1:A10";
            chart.Series.Add("B1:B10", xValuesRange); 
            chart.Series.Add("C1:C10", xValuesRange);
            chart.Legend.Position = eLegendPosition.Right;

            //automatically adjust columns width to text
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            string currentDirectoryPath = GetApplicationRoot();
            package.SaveAs(new FileInfo(Path.Combine(currentDirectoryPath, "example.xlsx")));
        }
        private static string GetApplicationRoot()
        {
            var exePath = Path.GetDirectoryName(System.Reflection
                              .Assembly.GetExecutingAssembly().CodeBase);
            Regex appPathMatcher = new Regex(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = appPathMatcher.Match(exePath).Value;
            return appRoot;
        }
    }
}
