using System;
using System.Collections.Generic;
using System.Linq;

namespace ohaERP_Library.GENERAL
{
    public class NextWorkdayCalculator
    {
        private readonly IEnumerable<TemplateForm.structHolidays> holidays;
        DateTime m_fromDate;

        public NextWorkdayCalculator() //IEnumerable<DateTime> holidays)
        {
            holidays = TemplateForm.g_holidays;
        }

        /// <summary>
        ///  Returns the first working day (forward or backward) if this the non-working day
        /// </summary>
        /// <param name="fromDate"></param>
        /// <param name="is_day_forward"></param>
        /// <param name="daysToAdd"></param>
        /// <returns></returns>
        public DateTime GetNextWorkDay(DateTime fromDate, int daysToAdd = 1)
        {
            DateTime nextWorkday = fromDate;
            m_fromDate = fromDate;

            if (!IsWorkday(fromDate))
            {
                for (int i = 0; i < daysToAdd; i++)
                {
                    do
                    {
                        nextWorkday = nextWorkday.AddDays(daysToAdd);
                    } while (!IsWorkday(nextWorkday));
                }

                return nextWorkday;
            }
            else
            {
                return fromDate;
            }
        }

        private bool IsWorkday(DateTime day)
        {
            return IsWeekday(day) && !IsHoliday(day);
        }

        private bool IsHoliday(DateTime day)
        {
            bool _ret_val = holidays.Any(d => d.date.Equals(day) && d.date >= m_fromDate);

            return _ret_val;
        }

        private bool IsWeekday(DateTime day)
        {
            return day.DayOfWeek != DayOfWeek.Saturday &&
                   day.DayOfWeek != DayOfWeek.Sunday;
        }
    }

}
