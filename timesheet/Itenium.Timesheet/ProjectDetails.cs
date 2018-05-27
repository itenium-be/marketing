using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Itenium.Timesheet
{
    public class DayInfo
    {
        private readonly DateTime _date;

        public bool IsHoliday => _date.DayOfWeek == DayOfWeek.Saturday || _date.DayOfWeek == DayOfWeek.Sunday;

        public DayInfo(int year, int month, int day)
        {
            _date = new DateTime(year, month, day);
        }

        public override string ToString() => $"{_date.ToString("ddd", CultureInfo.InvariantCulture)}, {_date.Day.ToOrdinal()}";
    }

    public class ProjectDetails
    {
        public string ConsultantName { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }

        public string Customer { get; set; }
        public string CustomerReference { get; set; }

        public string ProjectName { get; set; }

        public IEnumerable<DayInfo> GetDaysInMonth()
        {
            for (int day = 1; day <= DateTime.DaysInMonth(Year, Month); day++)
            {
                yield return new DayInfo(Year, Month, day);
            }
        }

        public ProjectDetails(int month)
        {
            ConsultantName = "Dan Padmore";
            Year = DateTime.Now.Year;
            Month = month;

            Customer = "Digipolis";
            CustomerReference = "IOR1804019";

            ProjectName = "Focus";
            // TODO: Manager?
        }

        public override string ToString() => $"{ConsultantName} ({Year}-{Month}) @{Customer}";
    }
}
