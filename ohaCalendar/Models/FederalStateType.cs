namespace ohaCalendar.Models
{
    public class FederalStateNameCategory
    {
        public string language { get; set; }
        public string text { get; set; }
    }

    public class FederalStateName
    {
        public string language { get; set; }
        public string text { get; set; }
    }

    public class FederalStateType
    {
        public string code { get; set; }
        public string isoCode { get; set; }
        public string shortName { get; set; }
        public List<FederalStateNameCategory> category { get; set; }
        public List<FederalStateName> name { get; set; }
        public List<string> officialLanguages { get; set; }
    }

}
