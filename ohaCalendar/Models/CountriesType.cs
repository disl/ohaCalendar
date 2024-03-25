namespace ohaCalendar.Models
{
    public class CountriesName
    {
        public string language { get; set; }
        public string text { get; set; }
    }

    public class CountriesType
    {
        public string isoCode { get; set; }
        public List<CountriesName> name { get; set; }
        public List<string> officialLanguages { get; set; }
    }
}
