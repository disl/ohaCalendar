namespace ohaCalendar.Models
{
    // Root myDeserializedClass = JsonConvert.DeserializeObject<List<Root>>(myJsonResponse);
    public class Name
    {
        public string language { get; set; }
        public string text { get; set; }
    }

    public class HolidayType
    {
        public string id { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string type { get; set; }
        public List<Name> name { get; set; }
        public bool nationwide { get; set; }
        public List<Subdivision> subdivisions { get; set; }
    }

    public class Subdivision
    {
        public string code { get; set; }
        public string shortName { get; set; }
    }


}
