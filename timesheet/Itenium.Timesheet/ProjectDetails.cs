using System;
using System.Collections.Generic;
using System.Text;

namespace Itenium.Timesheet
{
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
            CustomerReference = "";

            ProjectName = "Focus";
        }

        public override string ToString() => $"{ConsultantName} ({Year}-{Month}) @{Customer}";
    }
}
