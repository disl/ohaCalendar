namespace ohaCalendar.Models
{
    public class Feiertage
    {
        public string date { get; set; }
        public string fname { get; set; }
        public string all_states { get; set; }
        public string bw { get; set; }
        public string by { get; set; }
        public string be { get; set; }
        public string bb { get; set; }
        public string hb { get; set; }
        public string hh { get; set; }
        public string he { get; set; }
        public string mv { get; set; }
        public string ni { get; set; }
        public string nw { get; set; }
        public string rp { get; set; }
        public string sl { get; set; }
        public string sn { get; set; }
        public string st { get; set; }
        public string sh { get; set; }
        public string th { get; set; }
        public string comment { get; set; }
        public object augsburg { get; set; }
        public string katholisch { get; set; }
    }

    public class HolidayType
    {
        public string status { get; set; }
        public List<Feiertage> feiertage { get; set; }
    }
}
