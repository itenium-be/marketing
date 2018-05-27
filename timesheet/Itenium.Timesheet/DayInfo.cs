using System;
using System.Globalization;
using Nager.Date;

namespace Itenium.Timesheet
{
    public class DayInfo
    {
        private readonly DateTime _date;

        public bool IsHoliday => _date.DayOfWeek == DayOfWeek.Saturday || _date.DayOfWeek == DayOfWeek.Sunday || DateSystem.IsPublicHoliday(_date, CountryCode.BE);

        public string GetHolidayDesc()
        {
            if (!IsHoliday) return "";

            var publicHolidays = DateSystem.GetPublicHoliday("BE", _date, _date);
            foreach (var publicHoliday in publicHolidays)
            {
                if (publicHoliday.LocalName == "Pâques")
                    return "Pasen";

                if (publicHoliday.LocalName == "Weihnachten")
                    return "Kerstdag";

                return publicHoliday.LocalName;
            }
            return "";
        }

        public DayInfo(int year, int month, int day)
        {
            _date = new DateTime(year, month, day);
        }

        public override string ToString() => $"{_date.ToString("ddd", CultureInfo.InvariantCulture)}, {_date.Day.ToOrdinal()}";
    }
}