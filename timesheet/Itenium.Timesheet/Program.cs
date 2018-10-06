using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Itenium.Timesheet
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            using (var package = new ExcelPackage())
            {
                AddWorkaroundStyles(package);

                for (int month = 1; month <= 12; month++)
                {
                    string monthName = CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(month);
                    var sheet = package.Workbook.Worksheets.Add(monthName);

                    var projectDetails = new ProjectDetails(month);
                    var builder = new ExcelSheetBuilder(sheet, projectDetails);
                    builder.Build();
                }

                package.Workbook.Worksheets[DateTime.Now.Month].Select();

                var currentDllPath = new FileInfo(Environment.GetCommandLineArgs()[0]);
                var saveExcelAs = currentDllPath.DirectoryName + $"\\itenium-timesheet-{DateTime.Now.Year}.xlsx";
                package.SaveAs(new FileInfo(saveExcelAs));

                Console.WriteLine("Timesheet template saved as:");
                Console.WriteLine(saveExcelAs);
            }
        }

        private static void AddWorkaroundStyles(ExcelPackage package)
        {
            var leftStyle = package.Workbook.Styles.CreateNamedStyle("Left");
            leftStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            var centerStyle = package.Workbook.Styles.CreateNamedStyle("Center");
            centerStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var rightStyle = package.Workbook.Styles.CreateNamedStyle("Right");
            rightStyle.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }
    }
}
