using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Itenium.Timesheet
{
    public class ExcelSheetBuilder
    {
        private const string TimesheetEmail = "timesheet@itenium.be";

        private readonly ExcelWorksheet _sheet;
        private readonly ProjectDetails _details;

        public ExcelSheetBuilder(ExcelWorksheet sheet, ProjectDetails details)
        {
            _sheet = sheet;
            _details = details;
        }

        public void Build()
        {
            _sheet.Column(1).Width = 3;
            _sheet.Row(1).Height = 10;

            _sheet.Row(2).Height = 10;
            _sheet.Row(3).Height = 30;
            _sheet.Row(4).Height = 10;

            _sheet.Column(4).Width = 5;
            _sheet.Column(11).Width = 3;

            AddHeader();

            AddMonthTable();

            // TODO: manager signature etc
            // TODO: Add totals
            // TODO: Add holidays
        }

        private void AddMonthTable()
        {
            int startRow = 10;

            int currentRow = startRow;
            _sheet.Row(startRow).Height = 18;

            _sheet.Cells[currentRow, 2].TableHeader("Day");
            _sheet.Cells[currentRow, 3].TableHeader("# Hours");
            _sheet.Cells[currentRow, 3, currentRow, 4].Merge = true;
            _sheet.Cells[currentRow, 2, currentRow, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentRow++;

            foreach (var day in _details.GetDaysInMonth())
            {
                _sheet.Cells[currentRow, 2].Value = day;
                _sheet.Cells[currentRow, 2].StyleName = "Center";

                _sheet.Cells[currentRow, 3, currentRow, 4].Merge = true;
                _sheet.Cells[currentRow, 3].StyleName = "Center";

                if (day.IsHoliday)
                {
                    _sheet.Cells[currentRow, 2, currentRow, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _sheet.Cells[currentRow, 2, currentRow, 4].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }

                currentRow++;
            }

            currentRow--;

            var table = _sheet.Cells[startRow, 2, currentRow, 4];
            table.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            table.Style.Border.BorderAround(ExcelBorderStyle.Medium);
            _sheet.Cells[startRow, 2, startRow, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;


            _sheet.Cells[currentRow + 2, 2].Value = $"Please send back duly signed document by the 3rd working day of the following month to {TimesheetEmail}";
        }

        private void AddHeader()
        {
            string logoPath = Directory.GetCurrentDirectory() + @"\itenium-logo.png";
            var picture = _sheet.Drawings.AddPicture("itenium logo", Image.FromFile(logoPath));
            picture.SetPosition(25, 50);

            _sheet.Cells["F3"].Value = "TIMESHEET";
            _sheet.Cells["F3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            _sheet.Cells["F3"].Style.Font.Bold = true;
            _sheet.Cells["F3"].Style.Font.Size = 18;

            _sheet.Cells["B2:J4"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            _sheet.Cells["B2:J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells["B2:J4"].Style.Fill.BackgroundColor.SetColor(Color.White);

            _sheet.Column(2).Width = 15;

            _sheet.Cells["B6"].HeaderLabel("Consultant");
            _sheet.Cells["C6"].Value = _details.ConsultantName;

            _sheet.Cells["B7"].HeaderLabel("Month");
            _sheet.Cells["C7"].Value = CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(_details.Month) + " " + _details.Year;

            _sheet.Cells["H6"].HeaderLabel("Customer");
            _sheet.Cells["I6"].Value = _details.Customer;

            _sheet.Cells["H7"].HeaderLabel("Reference");
            _sheet.Cells["I7"].Value = _details.CustomerReference;

            _sheet.Row(5).Height = 10;

            _sheet.Row(6).Height = 22;
            _sheet.Row(8).Height = 10;
            _sheet.Cells["B6:J8"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            _sheet.Cells["B6:J8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells["B6:J8"].Style.Fill.BackgroundColor.SetColor(Color.White);

            _sheet.Row(9).Height = 10;
        }
    }
}
