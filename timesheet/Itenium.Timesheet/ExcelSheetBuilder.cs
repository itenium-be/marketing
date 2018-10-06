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
        private const string TimesheetEmail = "invoice@itenium.be";
        private const string TimesheetNotice = "Please send back duly signed document together with your invoice by the 2nd working day of the following month to:";

        private readonly ExcelWorksheet _sheet;
        private readonly ProjectDetails _details;

        private readonly int _startRow = 10;
        private int _endRow;

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
            _sheet.Row(4).Height = 7;

            _sheet.Column(4).Width = 5;
            _sheet.Column(11).Width = 3;

            AddHeader();

            AddMonthTable();

            AddAsideTable();
        }

        private void AddAsideTable()
        {
            _sheet.Column(6).Width = 3;
            _sheet.Column(7).Width = 2;
            _sheet.Column(9).Width = 27;
            _sheet.Column(10).Width = 2;

            _sheet.Cells["H10"].HeaderLabel("Project");
            _sheet.Cells["I10"].Value = _details.ProjectName;

            _sheet.Cells["I11"].StyleName = "Left";
            _sheet.Cells["H11"].HeaderLabel("Total Time");
            _sheet.Cells["I11"].Formula = $"SUM(C{_startRow + 1}:C{_endRow})";
            _sheet.Cells["I11"].Style.Numberformat.Format = "[HH]:MM";

            _sheet.Cells["I12"].StyleName = "Left";
            _sheet.Cells["H12"].HeaderLabel("Days");
            _sheet.Cells["I12"].Formula = "ROUND(I11 * 3, 2)";

            _sheet.Cells["H14"].HeaderLabel("Manager");

            _sheet.Cells["I14"].Style.Fill.PatternType = ExcelFillStyle.Solid;

            _sheet.Cells["H15"].HeaderLabel("Signature");

            _sheet.Cells[10, 7, 20, 10].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            _sheet.Cells[10, 7, 20, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
            _sheet.Cells[10, 7, 20, 10].Style.Fill.BackgroundColor.SetColor(Color.White);

            _sheet.Cells["I14"].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            _sheet.Cells["I14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        private void AddMonthTable()
        {
            int startRow = _startRow;

            int currentRow = startRow;
            _sheet.Row(startRow).Height = 18;

            _sheet.Cells[currentRow, 2].TableHeader("Day");
            _sheet.Cells[currentRow, 3].TableHeader("# Hours");
            _sheet.Cells[currentRow, 3, currentRow, 5].Merge = true;
            _sheet.Cells[currentRow, 2, currentRow, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            currentRow++;

            foreach (var day in _details.GetDaysInMonth())
            {
                _sheet.Cells[currentRow, 2].Value = day;
                _sheet.Cells[currentRow, 2].StyleName = "Center";

                _sheet.Cells[currentRow, 3, currentRow, 5].Merge = true;
                _sheet.Cells[currentRow, 3].StyleName = "Center";

                if (day.IsHoliday)
                {
                    _sheet.Cells[currentRow, 3].StyleName = "";
                    _sheet.Cells[currentRow, 2, currentRow, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    _sheet.Cells[currentRow, 2, currentRow, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                    _sheet.Cells[currentRow, 3].Value = day.GetHolidayDesc();
                }
                else
                {
                    _sheet.Cells[currentRow, 3].Style.Numberformat.Format = "H:MM";
                }

                currentRow++;
            }

            currentRow--;
            _endRow = currentRow;

            var table = _sheet.Cells[startRow, 2, currentRow, 5];
            table.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            table.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            table.Style.Border.BorderAround(ExcelBorderStyle.Medium);
            _sheet.Cells[startRow, 2, startRow, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            _sheet.Cells[currentRow + 2, 2].StyleName = "Center";
            _sheet.Cells[currentRow + 2, 2].Value = TimesheetNotice;
            _sheet.Cells[currentRow + 2, 2].Style.WrapText = true;
            _sheet.Cells[currentRow + 2, 2, currentRow + 2, 9].Merge = true;
            _sheet.Row(currentRow + 2).Height = 28;

            _sheet.Cells[currentRow + 3, 2, currentRow + 3, 9].Merge = true;
            _sheet.Cells[currentRow + 3, 2].StyleName = "Center";

            var ourEmailCell = _sheet.Cells[currentRow + 3, 2];
            ourEmailCell.Hyperlink = new Uri("mailto:" + TimesheetEmail, UriKind.Absolute);
            ourEmailCell.Value = TimesheetEmail;
            ourEmailCell.Style.Font.Color.SetColor(Color.Blue);
            ourEmailCell.Style.Font.UnderLine = true;
        }

        private void AddHeader()
        {
            var currentDllPath = new FileInfo(Environment.GetCommandLineArgs()[0]);
            string logoPath = currentDllPath.DirectoryName + @"\itenium-logo.png";
            var picture = _sheet.Drawings.AddPicture("itenium logo", Image.FromFile(logoPath));
            picture.SetPosition(28, 50);

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
        }
    }
}
