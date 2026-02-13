using ohaCalendar.Models;

namespace ohaCalendar.Services
{
    public interface IOutlookCalendarService
    {
        Task<List<OutlookCalendarItemType>> GetCalendarItemsAsync(CalendarEventFilter filter, CancellationToken cancellationToken = default);
        Task<List<OutlookCalendarItemType>> GetCalendarItemsByUserAsync(string userId, CalendarEventFilter filter, CancellationToken cancellationToken = default);
        Task<OutlookCalendarItemType?> GetCalendarItemByIdAsync(string eventId, string? calendarId = null, CancellationToken cancellationToken = default);
        Task<List<CalendarInfo>> GetUserCalendarsAsync(CancellationToken cancellationToken = default);
        Task<string?> GetOrCreateItemIdAsync(string eventId, CancellationToken cancellationToken = default);
        Task<bool> AddAttachmentAsync(string eventId, string filePath, string? calendarId = null, CancellationToken cancellationToken = default);
        Task<Stream?> GetAttachmentContentAsync(string eventId, string attachmentId, CancellationToken cancellationToken = default);
    }

    public class CalendarInfo
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public bool IsDefault { get; set; }
        public string Owner { get; set; } = string.Empty;
        public int TotalItems { get; set; }
    }
}
