namespace ohaCalendar.Models
{
    public class CalendarItemType
    {
        public CalendarItemType(DateTime day, int week_no, bool active, int countOfTermins)
        {
            Day = day;
            Week_no = week_no;
            Active = active;
            CountOfTermins = countOfTermins;
        }

        public DateTime Day { get; set; }
        public int Week_no { get; set; }
        public bool Active { get; set; }
        public int CountOfTermins { get; set; }
    }
}
