// Models/OutlookCalendarItemType.cs
namespace ohaCalendar.Models
{
    public class OutlookCalendarItemType
    {
        // Properties
        public string Id { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string BodyPreview { get; set; }
        public string ItemId { get; set; }
        public DateTime CreationTime { get; set; }
        public DateTime LastModifiedTime { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public int Duration { get; set; }
        public string Location { get; set; }
        public string Telephone { get; set; }
        public string MobilePhone { get; set; }
        public string Email { get; set; }
        public string ContactPerson { get; set; }
        public string CompanyName { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        public string Postcode { get; set; }
        public string Nation { get; set; }
        public List<string> Attachments { get; set; }
        public string Organizer { get; set; }
        public string RequiredAttendees { get; set; }
        public string OptionalAttendees { get; set; }
        public bool AllDayEvent { get; set; }
        public int BusyStatus { get; set; }
        public string CalendarName { get; set; }
        public string OnlineMeetingUrl { get; set; }
        public bool IsCancelled { get; set; }
        public string RecurrencePattern { get; set; }
        public string Sensitivity { get; set; }
        public int Importance { get; set; }
        public List<string> Categories { get; set; }

        // Parameterloser Konstruktor für Serialisierung
        public OutlookCalendarItemType()
        {
            Id = string.Empty;
            Subject = string.Empty;
            Body = string.Empty;
            BodyPreview = string.Empty;
            ItemId = string.Empty;
            Location = string.Empty;
            Telephone = string.Empty;
            MobilePhone = string.Empty;
            Email = string.Empty;
            ContactPerson = string.Empty;
            CompanyName = string.Empty;
            Street = string.Empty;
            City = string.Empty;
            Postcode = string.Empty;
            Nation = string.Empty;
            Attachments = new List<string>();
            Organizer = string.Empty;
            RequiredAttendees = string.Empty;
            OptionalAttendees = string.Empty;
            CalendarName = string.Empty;
            OnlineMeetingUrl = string.Empty;
            RecurrencePattern = string.Empty;
            Sensitivity = "Normal";
            Categories = new List<string>();
        }

        // Originaler Konstruktor für Abwärtskompatibilität
        public OutlookCalendarItemType(
            string subject,
            string body,
            string itemId,
            DateTime creationTime,
            DateTime start,
            DateTime end,
            double duration,
            string location,
            string telephone,
            string mobilePhone,
            string email,
            string contactPerson,
            string companyName,
            string street,
            string city,
            string postcode,
            string nation,
            List<string> attachments,
            string organizer,
            string requiredAttendees,
            string optionalAttendees,
            bool allDayEvent,
            int busyStatus)
        {
            Subject = subject ?? string.Empty;
            Body = body ?? string.Empty;
            ItemId = itemId ?? string.Empty;
            CreationTime = creationTime;
            Start = start;
            End = end;
            Duration = (int)duration;
            Location = location ?? string.Empty;
            Telephone = telephone ?? string.Empty;
            MobilePhone = mobilePhone ?? string.Empty;
            Email = email ?? string.Empty;
            ContactPerson = contactPerson ?? string.Empty;
            CompanyName = companyName ?? string.Empty;
            Street = street ?? string.Empty;
            City = city ?? string.Empty;
            Postcode = postcode ?? string.Empty;
            Nation = nation ?? string.Empty;
            Attachments = attachments ?? new List<string>();
            Organizer = organizer ?? string.Empty;
            RequiredAttendees = requiredAttendees ?? string.Empty;
            OptionalAttendees = optionalAttendees ?? string.Empty;
            AllDayEvent = allDayEvent;
            BusyStatus = busyStatus;

            // Initialize other properties
            Id = string.Empty;
            BodyPreview = string.Empty;
            LastModifiedTime = creationTime;
            CalendarName = string.Empty;
            OnlineMeetingUrl = string.Empty;
            IsCancelled = false;
            RecurrencePattern = string.Empty;
            Sensitivity = "Normal";
            Importance = 1;
            Categories = new List<string>();
        }
    }

    public class CalendarEventFilter
    {
        public string? CalendarId { get; set; }
        public string? CalendarName { get; set; }
        public string? CalendarNameAlternate { get; set; }
        public string? SearchInSubject { get; set; }
        public string? SearchInBody { get; set; }
        public DateTime Start { get; set; }
        public DateTime? Ende { get; set; }
        public bool OnlyUnRead { get; set; }
        public bool IsLightAlgorithmus { get; set; }
        public int? MaxResults { get; set; }
    }

    public class CalendarInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public bool IsDefault { get; set; }
        public string Owner { get; set; } = string.Empty;
    }
}