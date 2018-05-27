using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using OfficeOpenXml;

namespace Itenium.Timesheet
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            using (var package = new ExcelPackage())
            {
                for (int i = 1; i <= 12; i++)
                {
                    string monthName = CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(i);
                    var sheet = package.Workbook.Worksheets.Add(monthName);

                    string logoPath = Directory.GetCurrentDirectory() + @"\itenium-logo.png";
                    var picture = sheet.Drawings.AddPicture("itenium logo", Image.FromFile(logoPath));
                    picture.SetPosition(0, 0, 2, 0);
                }

                var saveExcelAs = Directory.GetCurrentDirectory() + $"\\..\\itenium-{DateTime.Now.Year}.xlsx";
                package.SaveAs(new FileInfo(saveExcelAs));
            }
        }
    }
}
