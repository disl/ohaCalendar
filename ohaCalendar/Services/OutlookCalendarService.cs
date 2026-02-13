using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ohaCalendar.Models;
using Attendee = Microsoft.Graph.Models.Attendee;
using AttendeeType = Microsoft.Graph.Models.AttendeeType;
using DateTimeTimeZone = Microsoft.Graph.Models.DateTimeTimeZone;
using Event = Microsoft.Graph.Models.Event;
using FileAttachment = Microsoft.Graph.Models.FileAttachment;
using FreeBusyStatus = Microsoft.Graph.Models.FreeBusyStatus;
using Importance = Microsoft.Graph.Models.Importance;
using SingleValueLegacyExtendedProperty = Microsoft.Graph.Models.SingleValueLegacyExtendedProperty;

namespace ohaCalendar.Services
{
    public class OutlookCalendarService : IOutlookCalendarService
    {
        #region ENUMS UND KONSTANTEN
        // s. http://msdn.microsoft.com/en-us/library/aa908088.aspx
        public enum OlBusyStatus
        {
            olFree = 0,
            olTentative = 1,
            olBusy = 2,
            olOutOfOffice = 3,
            olWorkingElsewhere = 4
        }

        // s. http://msdn.microsoft.com/en-us/library/aa911624.aspx
        public enum OlImportance
        {
            olImportanceLow = 0,
            olImportanceNormal = 1,
            olImportanceHigh = 2
        }

        // s. http://msdn.microsoft.com/en-us/library/bb208072.aspx
        public enum OlFolderType
        {
            olFolderCalendar = 9 // The Calendar folder. 
    ,
            olFolderConflicts = 19 // The Conflicts folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderContacts = 10 // The Contacts folder. 
    ,
            olFolderDeletedItems = 3 // The Deleted Items folder. 
    ,
            olFolderDrafts = 16 // The Drafts folder. 
    ,
            olFolderInbox = 6 // The Inbox folder. 
    ,
            olFolderJournal = 11 // The Journal folder. 
    ,
            olFolderJunk = 23 // The Junk E-Mail folder. 
    ,
            olFolderLocalFailures = 21 // The Local Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderManagedEmail = 29 // The top-level folder in the Managed Folders group. For more information on Managed Folders, see Help in Microsoft Outlook. Only available for an Exchange account. 
    ,
            olFolderNotes = 12 // The Notes folder. 
    ,
            olFolderOutbox = 4 // The Outbox folder. 
    ,
            olFolderSentMail = 5 // The Sent Mail folder. 
    ,
            olFolderServerFailures = 22 // The Server Failures folder (subfolder of Sync Issues folder). Only available for an Exchange account. 
    ,
            olFolderSyncIssues = 20 // The Sync Issues folder. Only available for an Exchange account. 
    ,
            olFolderTasks = 13 // The Tasks folder. 
    ,
            olFolderToDo = 28 // The To Do folder. 
    ,
            olPublicFoldersAllPublicFolders = 18 // The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account. 
    ,
            olFolderRssFeeds = 25 // The RSS Feeds folder. 
        }

        // s. http://www.online-excel.de/excel/singsel_vba.php?f=85
        public enum OlItemType
        {
            olMailItem = 0,
            olAppointmentItem = 1,
            olContactItem = 2,
            olTaskItem = 3,
            olJournalItem = 4,
            olNoteItem = 5,
            olPostItem = 6,
            olDistributionListItem = 7
        }

        public class OlAttachmentType
        {
            public const Int32 olByReference = 4;
            public const Int32 olByValue = 1;
            public const Int32 olEmbeddedItem = 5;
            public const Int32 olOLE = 6;
        }

        public enum OlUserPropertyType
        {
            olCombination = 19,  // The Property type Is a combination Of other types. It corresponds To the MAPI type PT_STRING8.
            olCurrency = 14, // Represents a Currency Property type. It corresponds To the MAPI type PT_CURRENCY.
            olDateTime = 5,  // Represents a DateTime Property type. It corresponds To the MAPI type PT_SYSTIME.
            olDuration = 7,  // Represents a time duration Property type. It corresponds To the MAPI type PT_LONG.
            olEnumeration = 21,  // Represents an enumeration Property type. It corresponds To the MAPI type PT_LONG.
            olFormula = 18,  // Represents a formula Property type. It corresponds To the MAPI type PT_STRING8. See UserDefinedProperty.Formula Property.
            olInteger = 20,  // Represents an Integer number Property type. It corresponds To the MAPI type PT_LONG.
            olKeywords = 11, // Represents a String array Property type used To store keywords. It corresponds To the MAPI type PT_MV_STRING8.
            olNumber = 3, // Represents a Double number Property type. It corresponds To the MAPI type PT_DOUBLE.
            olOutlookInternal = 0,   // Represents an Outlook internal Property type.
            olPercent = 12,  // Represents a Double number Property type used To store a percentage. It corresponds To the MAPI type PT_LONG.
            olSmartFrom = 22,    // Represents a smart from Property type. This Property indicates that If the From Property Of an Outlook item Is empty, Then the To Property should be used instead.
            olText = 1,  // Represents a String Property type. It corresponds To the MAPI type PT_STRING8.
            olYesNo = 6 // Represents a yes/no (Boolean) property type. It corresponds to the MAPI type PT_BOOLEAN.
        }


        #endregion

        private readonly GraphServiceClient _graphClient;
        private static string? m_current_user_email;

        //private readonly ILogger<OutlookCalendarService> // _logger;
        private const string EXTENDED_PROPERTY_ID = "String {00020329-0000-0000-C000-000000000046} Name GraphItemID";

        public OutlookCalendarService()
        {
            // 1. Zugangsdaten (sollten idealerweise aus appsettings.json kommen)
            string tenantId = "c5f5a5dc-4bcd-48b2-97c4-a108275e87ba";
            string clientId = "9a784777-43ad-4727-8b16-08ceb56735a1";
            string clientSecret = "XEZ8Q~KS9RP70FcSqObaUjsppjdk5KbMH054fbpU";

            // 2. Credential erstellen
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

            // 3. Scopes festlegen (immer .default für ClientSecretFlow)
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // 4. Den readonly GraphServiceClient initialisieren
            _graphClient = new GraphServiceClient(credential, scopes);
        }


        public OutlookCalendarService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;

        }

        #region Hauptmethoden

        /// <summary>
        /// Holt Kalenderitems für den aktuellen Benutzer
        /// </summary>
        public async Task<List<OutlookCalendarItemType>> GetCalendarItemsAsync(
            ohaCalendar.Models.CalendarEventFilter filter,
            CancellationToken cancellationToken = default)
        {
            try
            {
                await SetCurrentUserEmail();

                // Kalenderansicht abrufen
                var events = await GetCalendarViewAsync(
                    filter.CalendarId,
                    filter.Start,
                    (DateTime)filter.Ende,
                    cancellationToken);

                // Filtern und konvertieren
                var result = await FilterAndConvertEvents(
                    events,
                    filter,
                    m_current_user_email,  //calendar?.Name ?? "Default",
                    cancellationToken);

                return result;
            }
            catch (ServiceException ex)
            {
                //// _logger.LogError(ex, "Graph API Fehler: {Message}", ex.Message);
                throw new InvalidOperationException($"Fehler beim Kalenderzugriff: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                //// _logger.LogError(ex, "Graph API Fehler: {Message}", ex.Message);
                throw new InvalidOperationException($"Fehler beim Kalenderzugriff: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Holt Kalenderitems für einen bestimmten Benutzer
        /// </summary>
        public async Task<List<OutlookCalendarItemType>> GetCalendarItemsByUserAsync(
            string userId,
            CalendarEventFilter filter,
            CancellationToken cancellationToken = default)
        {
            try
            {
                var endDate = filter.Ende ?? DateTime.Now;

                // Benutzer-spezifischen Request Builder verwenden
                EventCollectionResponse? calendarView;

                if (string.IsNullOrEmpty(userId))
                    userId = m_current_user_email;

                //if (string.IsNullOrEmpty(userId) || userId == "me")
                //{
                //    calendarView = await _graphClient.Me.CalendarView
                //        .GetAsync(requestConfig =>
                //        {
                //            requestConfig.QueryParameters.StartDateTime = filter.Start.ToString("o");
                //            requestConfig.QueryParameters.EndDateTime = endDate.ToString("o");
                //            requestConfig.QueryParameters.Select = GetEventSelectFields();
                //            requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                //            requestConfig.Headers.Add("Prefer", $"outlook.timezone=\"{GetCurrentTimeZone()}\"");
                //        }, cancellationToken);
                //}
                //else
                //{
                calendarView = await _graphClient.Users[userId].CalendarView
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.StartDateTime = filter.Start.ToString("o");
                        requestConfig.QueryParameters.EndDateTime = endDate.ToString("o");
                        requestConfig.QueryParameters.Select = GetEventSelectFields();
                        requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                        requestConfig.Headers.Add("Prefer", $"outlook.timezone=\"{GetCurrentTimeZone()}\"");
                    }, cancellationToken);
                //}

                var events = calendarView?.Value ?? new List<Event>();

                return await FilterAndConvertEvents(events, filter, $"User_{userId}", cancellationToken);
            }
            catch (ServiceException ex)
            {
                // _logger.LogError(ex, "Fehler beim Abrufen der Kalenderitems für Benutzer {UserId}", userId);
                throw;
            }
        }

        /// <summary>
        /// Holt ein einzelnes Kalenderitem per ID
        /// </summary>
        public async Task<OutlookCalendarItemType?> GetCalendarItemByIdAsync(
            string eventId,
            string? calendarId = null,
            CancellationToken cancellationToken = default)
        {
            try
            {
                Event? eventItem;

                if (!string.IsNullOrEmpty(calendarId))
                {
                    eventItem = await _graphClient.Me.Calendars[calendarId].Events[eventId]
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Select = GetEventSelectFields();
                            requestConfig.QueryParameters.Expand = new[] {
                                "attachments",
                                $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')"
                            };
                        }, cancellationToken);
                }
                else
                {
                    eventItem = await _graphClient.Me.Events[eventId]
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Select = GetEventSelectFields();
                            requestConfig.QueryParameters.Expand = new[] {
                                "attachments",
                                $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')"
                            };
                        }, cancellationToken);
                }

                if (eventItem == null)
                    return null;

                return await ConvertToOutlookCalendarItemType(eventItem, cancellationToken);
            }
            catch (ServiceException ex)
            {
                // _logger.LogError(ex, "Fehler beim Abrufen des Events {EventId}", eventId);
                return null;
            }
        }

        /// <summary>
        /// Listet alle verfügbaren Kalender des Benutzers auf
        /// </summary>
        public async Task<List<CalendarInfo>> GetUserCalendarsAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                var result = new List<CalendarInfo>();

                // Default Kalender abrufen
                var defaultCalendar = await _graphClient.Me.Calendar
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[] { "id", "name", "color", "owner" };
                    }, cancellationToken);

                if (defaultCalendar != null)
                {
                    result.Add(new CalendarInfo
                    {
                        Id = defaultCalendar.Id ?? string.Empty,
                        Name = defaultCalendar.Name ?? "Kalender",
                        Color = defaultCalendar.Color?.ToString() ?? "Auto",
                        IsDefault = true,
                        Owner = defaultCalendar.Owner?.Address ?? "Eigener Kalender"
                    });
                }

                // Alle anderen Kalender abrufen
                var calendars = await _graphClient.Me.Calendars
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Select = new[] { "id", "name", "color", "isDefaultCalendar", "owner" };
                    }, cancellationToken);

                if (calendars?.Value != null)
                {
                    foreach (var cal in calendars.Value.Where(c => c.Id != defaultCalendar?.Id))
                    {
                        result.Add(new CalendarInfo
                        {
                            Id = cal.Id ?? string.Empty,
                            Name = cal.Name ?? "Unbenannt",
                            Color = cal.Color?.ToString() ?? "Auto",
                            IsDefault = cal.IsDefaultCalendar ?? false,
                            Owner = cal.Owner?.Address ?? "Gemeinsamer Kalender"
                        });
                    }
                }

                return result;
            }
            catch (ServiceException ex)
            {
                // _logger.LogError(ex, "Fehler beim Abrufen der Kalenderliste");
                throw;
            }
        }

        /// <summary>
        /// Erstellt oder ruft eine eindeutige ItemID für ein Event ab
        /// </summary>
        public async Task<string?> GetOrCreateItemIdAsync(string eventId, CancellationToken cancellationToken = default)
        {
            try
            {
                // Prüfen ob bereits eine ItemID existiert
                var eventItem = await _graphClient.Users[m_current_user_email].Events[eventId]
                                .GetAsync(requestConfig =>
                                {
                                    requestConfig.QueryParameters.Expand = new[]
                                    {
                                        $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')"
                                    };
                                }, cancellationToken);

                var existingProp = eventItem?.SingleValueExtendedProperties?.FirstOrDefault();

                if (existingProp != null && !string.IsNullOrEmpty(existingProp.Value))
                {
                    return existingProp.Value;
                }

                // Neue ID erstellen und speichern
                var newId = Guid.NewGuid().ToString();

                var patchEvent = new Event
                {
                    SingleValueExtendedProperties = new List<SingleValueLegacyExtendedProperty>
                    {
                        new SingleValueLegacyExtendedProperty
                        {
                            Id = EXTENDED_PROPERTY_ID,
                            Value = newId
                        }
                    }
                };

                //await _graphClient.Me.Events[eventId]
                //    .PatchAsync(patchEvent);

                await _graphClient.Users[m_current_user_email].Events[eventId]
                        .PatchAsync(patchEvent, cancellationToken: cancellationToken);

                return newId;
            }
            catch (ServiceException ex)
            {
                // _logger.LogError(ex, "Fehler beim Erstellen/Abrufen der ItemID für {EventId}", eventId);
                return Guid.NewGuid().ToString();
            }
        }

        /// <summary>
        /// Fügt einen Anhang zu einem Kalenderitem hinzu
        /// </summary>
        public async Task<bool> AddAttachmentAsync(
            string eventId,
            string filePath,
            string? calendarId = null,
            CancellationToken cancellationToken = default)
        {
            try
            {
                var fileBytes = await File.ReadAllBytesAsync(filePath, cancellationToken);
                var fileName = Path.GetFileName(filePath);

                var attachment = new FileAttachment
                {
                    OdataType = "#microsoft.graph.fileAttachment",
                    Name = fileName,
                    ContentType = GetContentType(fileName),
                    ContentBytes = fileBytes
                };

                if (!string.IsNullOrEmpty(calendarId))
                {
                    await _graphClient.Me.Calendars[calendarId].Events[eventId]
                        .Attachments
                        .PostAsync(attachment);
                }
                else
                {
                    await _graphClient.Me.Events[eventId]
                        .Attachments
                        .PostAsync(attachment);
                }

                return true;
            }
            catch (Exception ex)
            {
                // _logger.LogError(ex, "Fehler beim Hinzufügen des Attachments {FilePath}", filePath);
                return false;
            }
        }

        /// <summary>
        /// Holt den Inhalt eines Attachments
        /// </summary>
        public async Task<Stream?> GetAttachmentContentAsync(
            string eventId,
            string attachmentId,
            CancellationToken cancellationToken = default)
        {
            try
            {
                var attachment = await _graphClient.Me.Events[eventId]
                    .Attachments[attachmentId]
                    .GetAsync(cancellationToken: cancellationToken);

                if (attachment is FileAttachment fileAttachment && fileAttachment.ContentBytes != null)
                {
                    return new MemoryStream(fileAttachment.ContentBytes);
                }

                return null;
            }
            catch (Exception ex)
            {
                // _logger.LogError(ex, "Fehler beim Abrufen des Attachments {AttachmentId}", attachmentId);
                return null;
            }
        }

        #endregion

        #region Private Hilfsmethoden

        //private async Task<Microsoft.Graph.Models.Calendar?> GetTargetCalendarAsync(
        //    string? calendarId,
        //    string? calendarName,
        //    string? calendarNameAlternate,
        //    CancellationToken cancellationToken)
        //{
        //    // 1. Direkt nach ID suchen
        //    if (!string.IsNullOrEmpty(calendarId))
        //    {
        //        try
        //        {
        //            return await _graphClient.Me.Calendars[calendarId]
        //                .GetAsync(null, cancellationToken);
        //        }
        //        catch
        //        {
        //            // Ignorieren, weitersuchen
        //        }
        //    }

        //    // 2. Nach Namen suchen (Primär)
        //    if (!string.IsNullOrEmpty(calendarName))
        //    {
        //        var calendar = await FindCalendarByNameAsync(calendarName, cancellationToken);
        //        if (calendar != null)
        //            return calendar;
        //    }

        //    // 3. Nach Alternativnamen suchen
        //    if (!string.IsNullOrEmpty(calendarNameAlternate))
        //    {
        //        var calendar = await FindCalendarByNameAsync(calendarNameAlternate, cancellationToken);
        //        if (calendar != null)
        //            return calendar;
        //    }

        //    // 4. Default-Kalender zurückgeben
        //    try
        //    {
        //        return await _graphClient.Me.Calendar
        //            .GetAsync(cancellationToken: cancellationToken);
        //    }
        //    catch
        //    {
        //        return null;
        //    }
        //}

        private async Task<Microsoft.Graph.Models.Calendar?> FindCalendarByNameAsync(string name, CancellationToken cancellationToken)
        {
            try
            {
                var calendars = await _graphClient.Me.Calendars
                    .GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Filter = $"name eq '{name.Replace("'", "''")}'";
                        requestConfig.QueryParameters.Top = 1;
                        requestConfig.QueryParameters.Select = new[] { "id", "name" };
                    }, cancellationToken);

                return calendars?.Value?.FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        private async Task<List<Event>> GetCalendarViewAsync(
                                            string? calendarId,
                                            DateTime start,
                                            DateTime end,
                                            CancellationToken cancellationToken)
        {
            try
            {
                EventCollectionResponse? response;

                // Verwenden Sie UTC für die Datumsangaben
                var startUtc = start.ToUniversalTime();
                var endUtc = end.ToUniversalTime();

                if (!string.IsNullOrEmpty(calendarId))
                {
                    response = await _graphClient.Me.Calendars[calendarId].CalendarView
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.StartDateTime = startUtc.ToString("o");
                            requestConfig.QueryParameters.EndDateTime = endUtc.ToString("o");
                            requestConfig.QueryParameters.Select = GetEventSelectFields();
                            requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                            requestConfig.QueryParameters.Orderby = new[] { "start/dateTime" };
                            // Wichtig: Kein Prefer Header mit nicht-ASCII Zeichen!
                            // requestConfig.Headers.Add("Prefer", $"outlook.timezone=\"{GetCurrentTimeZone()}\"");
                        }, cancellationToken);
                }
                else
                {
                    //response = await _graphClient.Me.CalendarView
                    //    .GetAsync(requestConfig =>
                    //    {
                    //        requestConfig.QueryParameters.StartDateTime = startUtc.ToString("o");
                    //        requestConfig.QueryParameters.EndDateTime = endUtc.ToString("o");
                    //        requestConfig.QueryParameters.Select = GetEventSelectFields();
                    //        //requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                    //        //requestConfig.QueryParameters.Orderby = new[] { "start/dateTime" };
                    //        // Kein Prefer Header
                    //    }, cancellationToken);

                    //string targetUser = "dimitri.sluzki@haas.de"; // Oder die GUID des Users





                    //var users = await _graphClient.Users.GetAsync(config =>
                    //{
                    //    config.QueryParameters.Filter = $"displayName eq '{Environment.UserName}'";
                    //    config.QueryParameters.Select = new[] { "userPrincipalName" };
                    //});



                    response = await _graphClient.Users[m_current_user_email].CalendarView
                        .GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.StartDateTime = startUtc.ToString("o");
                            requestConfig.QueryParameters.EndDateTime = endUtc.ToString("o");
                            requestConfig.QueryParameters.Select = GetEventSelectFields();
                            requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                            requestConfig.QueryParameters.Orderby = new[] { "start/dateTime" };
                            // Wichtig: Kein Prefer Header mit nicht-ASCII Zeichen!
                            // requestConfig.Headers.Add("Prefer", $"outlook.timezone=\"{GetCurrentTimeZone()}\"");
                        });
                }

                return response?.Value ?? new List<Event>();
            }
            catch (ServiceException ex)
            {
                //_logger.LogWarning(ex, "CalendarView nicht verfügbar, verwende Events-Query");

                // Fallback mit UTC
                try
                {
                    EventCollectionResponse? response;
                    var startUtc = start.ToUniversalTime();
                    var endUtc = end.ToUniversalTime();

                    var filter = $"start/dateTime ge '{startUtc:yyyy-MM-ddTHH:mm:ss}Z' and end/dateTime le '{endUtc:yyyy-MM-ddTHH:mm:ss}Z'";

                    if (!string.IsNullOrEmpty(calendarId))
                    {
                        response = await _graphClient.Me.Calendars[calendarId].Events
                            .GetAsync(requestConfig =>
                            {
                                requestConfig.QueryParameters.Filter = filter;
                                requestConfig.QueryParameters.Select = GetEventSelectFields();
                                requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                                requestConfig.QueryParameters.Orderby = new[] { "start/dateTime" };
                            }, cancellationToken);
                    }
                    else
                    {
                        response = await _graphClient.Me.Events
                            .GetAsync(requestConfig =>
                            {
                                requestConfig.QueryParameters.Filter = filter;
                                requestConfig.QueryParameters.Select = GetEventSelectFields();
                                requestConfig.QueryParameters.Expand = new[] { $"singleValueExtendedProperties($filter=Id eq '{EXTENDED_PROPERTY_ID}')" };
                                requestConfig.QueryParameters.Orderby = new[] { "start/dateTime" };
                            }, cancellationToken);
                    }

                    return response?.Value ?? new List<Event>();
                }
                catch
                {
                    return new List<Event>();
                }
            }
        }

        private async Task SetCurrentUserEmail()
        {
            try
            {
                if(!string.IsNullOrEmpty(m_current_user_email))
                    return;

                // Wir nehmen den Windows Namen (z.B. "DSluzki")
                // Wenn du das erste Zeichen abschneiden willst (D + Username):
                string windowsLoginName = Environment.UserName.Length > 1
                    ? Environment.UserName.Substring(1)
                    : Environment.UserName;

                // GEZIELTE SUCHE: Viel stabiler als die ganze Liste zu laden
                var searchResponse = await _graphClient.Users.GetAsync(config =>
                {
                    // Filtert direkt in der Cloud nach dem Nachnamen oder UPN
                    config.QueryParameters.Filter = $"startsWith(surname, '{windowsLoginName}') or startsWith(userPrincipalName, '{windowsLoginName}')";
                    config.QueryParameters.Select = new[] { "mail", "userPrincipalName", "id" };
                });

                var currentUser = searchResponse?.Value?.FirstOrDefault();

                if (currentUser != null)
                {
                    m_current_user_email = currentUser.Mail ?? currentUser.UserPrincipalName ?? currentUser.Id;
                }
                else
                {
                    // Fallback auf die Standard-Logik deiner Firma, falls nichts gefunden wurde
                    m_current_user_email = $"{windowsLoginName}@haas.de";
                }
            }
            catch (Exception ex)
            {
                // NIEMALS "me" setzen bei ClientSecret! 
                // Stattdessen einen festen Admin-User oder eine Fehlermeldung
                m_current_user_email = "Dimitri.sluzki@haas.de";
                // Hier könntest du ex.Message loggen, um zu sehen warum GetAsync() fehlschlägt
            }
        }

        //private async Task SetCurrentUserEmail()
        //{
        //    try
        //    {
        //        //string userName = Environment.UserName;
        //        var users_1 = (await _graphClient?.Users?.GetAsync())?.Value;


        //        string windowsLoginName = Environment.UserName.Substring(1, Environment.UserName.Length - 1);  // DUsername -> Username

        //        var currentUser = users_1.FirstOrDefault(u =>
        //          u.Surname != null && u.Surname.StartsWith(windowsLoginName, StringComparison.OrdinalIgnoreCase)
        //        );

        //        m_current_user_email = currentUser.Mail ?? currentUser.UserPrincipalName ?? currentUser.Id;
        //    }
        //    catch(Exception ex)
        //    {
        //        m_current_user_email = "me"; // Fallback auf "me"
        //    }
        //}

        private async Task<List<OutlookCalendarItemType>> FilterAndConvertEvents(
            List<Event> events,
            CalendarEventFilter filter,
            string calendarName,
            CancellationToken cancellationToken)
        {
            var result = new List<OutlookCalendarItemType>();

            foreach (var ev in events)
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // Text-Filter anwenden
                    //if (!ApplySearchFilter(ev, filter.SearchInSubject, filter.SearchInBody))
                    //    continue;

                    // Konvertieren und hinzufügen
                    var converted = await ConvertToOutlookCalendarItemType(ev, cancellationToken);
                    converted.CalendarName = calendarName;

                    result.Add(converted);

                    // Max-Results Limit
                    if (filter.MaxResults.HasValue && result.Count >= filter.MaxResults.Value)
                        break;
                }
                catch (Exception ex)
                {
                    // _logger.LogWarning(ex, "Fehler bei Konvertierung von Event {Id}", ev.Id);
                }
            }

            return result;
        }

        private bool ApplySearchFilter(Event ev, string? searchSubject, string? searchBody)
        {
            if (string.IsNullOrWhiteSpace(searchSubject) && string.IsNullOrWhiteSpace(searchBody))
                return true;

            var subjectMatch = !string.IsNullOrWhiteSpace(searchSubject) &&
                ev.Subject?.Contains(searchSubject, StringComparison.OrdinalIgnoreCase) == true;

            var bodyMatch = !string.IsNullOrWhiteSpace(searchBody) && (
                ev.Body?.Content?.Contains(searchBody, StringComparison.OrdinalIgnoreCase) == true ||
                ev.BodyPreview?.Contains(searchBody, StringComparison.OrdinalIgnoreCase) == true
            );

            return subjectMatch || bodyMatch;
        }

        private async Task<OutlookCalendarItemType> ConvertToOutlookCalendarItemType(
            Event ev,
            CancellationToken cancellationToken)
        {
            var startDateTime = ParseGraphDateTime(ev.Start);
            var endDateTime = ParseGraphDateTime(ev.End);
            var creationTime = ev.CreatedDateTime?.DateTime ?? DateTime.Now;

            // ItemID aus Extended Properties holen
            var itemId = GetItemIdFromExtendedProperties(ev);
            if (string.IsNullOrEmpty(itemId) && ev.Id != null)
            {
                itemId = await GetOrCreateItemIdAsync(ev.Id, cancellationToken) ?? string.Empty;
            }

            // Anhänge verarbeiten
            var attachments = await ProcessAttachments(ev, cancellationToken);

            // Teilnehmer parsen
            var (required, optional) = ParseAttendees(ev.Attendees);

            // Location parsen
            var location = string.Empty;
            if (ev.Location != null)
            {
                location = ev.Location.DisplayName ??
                          ev.Location.Address?.Street ??
                          string.Empty;
            }

            var item = new OutlookCalendarItemType
            {
                Subject = ev.Subject ?? string.Empty,
                Body = ev.Body?.Content ?? string.Empty,
                BodyPreview = ev.BodyPreview ?? string.Empty,
                ItemId = itemId,
                CreationTime = creationTime,
                LastModifiedTime = ev.LastModifiedDateTime?.DateTime ?? creationTime,
                Start = startDateTime,
                End = endDateTime,
                Duration = (int)(endDateTime - startDateTime).TotalMinutes,
                Location = location,
                Attachments = attachments,
                Organizer = ev.Organizer?.EmailAddress?.Name ??
                           ev.Organizer?.EmailAddress?.Address ??
                           string.Empty,
                RequiredAttendees = required,
                OptionalAttendees = optional,
                AllDayEvent = ev.IsAllDay ?? false,
                BusyStatus = ConvertFreeBusyStatus(ev.ShowAs),
                OnlineMeetingUrl = ev.OnlineMeeting?.JoinUrl ?? string.Empty,
                IsCancelled = ev.IsCancelled ?? false,
                RecurrencePattern = ev.Recurrence?.Pattern?.Type?.ToString() ?? string.Empty,
                Sensitivity = ev.Sensitivity?.ToString() ?? "Normal",
                Importance = (int)(ev.Importance ?? Importance.Normal),
                Categories = ev.Categories?.ToList() ?? new List<string>(),
            };

            //var item = new OutlookCalendarItemType
            //{
            //    //Id = ev.Id ?? string.Empty,
            //    Subject = ev.Subject ?? string.Empty,
            //    Body = ev.Body?.Content ?? string.Empty,
            //    BodyPreview = ev.BodyPreview ?? string.Empty,
            //    ItemId = itemId,
            //    CreationTime = creationTime,
            //    LastModifiedTime = ev.LastModifiedDateTime?.DateTime ?? creationTime,
            //    Start = startDateTime,
            //    End = endDateTime,
            //    Duration = (int)(endDateTime - startDateTime).TotalMinutes,
            //    Location = location,
            //    Attachments = attachments,
            //    Organizer = ev.Organizer?.EmailAddress?.Name ??
            //               ev.Organizer?.EmailAddress?.Address ??
            //               string.Empty,
            //    RequiredAttendees = required,
            //    OptionalAttendees = optional,
            //    AllDayEvent = ev.IsAllDay ?? false,
            //    BusyStatus = ConvertFreeBusyStatus(ev.ShowAs),
            //    OnlineMeetingUrl = ev.OnlineMeeting?.JoinUrl ?? string.Empty,
            //    IsCancelled = ev.IsCancelled ?? false,
            //    RecurrencePattern = ev.Recurrence?.Pattern?.Type?.ToString() ?? string.Empty,
            //    Sensitivity = ev.Sensitivity?.ToString() ?? "Normal",
            //    Importance = (int)(ev.Importance ?? Importance.Normal),
            //    Categories = ev.Categories?.ToList() ?? new List<string>(),
            //    CalendarName = ev.Calendar.Name
            //};

            return item;
        }

        private DateTime ParseGraphDateTime(DateTimeTimeZone? dateTimeZone)
        {
            if (dateTimeZone?.DateTime == null)
                return DateTime.Now;

            if (DateTime.TryParse(dateTimeZone.DateTime, out var result))
                return result;

            return DateTime.Now;
        }

        private string GetItemIdFromExtendedProperties(Event ev)
        {
            var prop = ev.SingleValueExtendedProperties?.FirstOrDefault(p =>
                p.Id == EXTENDED_PROPERTY_ID);

            return prop?.Value ?? string.Empty;
        }

        private async Task<List<string>> ProcessAttachments(Event ev, CancellationToken cancellationToken)
        {
            var attachmentPaths = new List<string>();

            if (ev.Id != null)
            {
                try
                {
                    // Attachments nur abrufen wenn nicht bereits geladen
                    if (ev.Attachments == null)
                    {
                        //var attachments = await _graphClient.Me.Events[ev.Id]
                        //    .Attachments
                        //    .GetAsync(cancellationToken: cancellationToken);

                        var attachments = await _graphClient.Users[m_current_user_email].Events[ev.Id]
                                                .Attachments
                                                .GetAsync(cancellationToken: cancellationToken);
                        ev.Attachments = attachments?.Value;
                    }

                    if (ev.Attachments != null)
                    {
                        foreach (var attachment in ev.Attachments.OfType<FileAttachment>())
                        {
                            try
                            {
                                if (attachment.Name?.Equals("item_info.xml", StringComparison.OrdinalIgnoreCase) == true)
                                {
                                    var tempFile = Path.Combine(
                                        Path.GetTempPath(),
                                        $"item_info_{Guid.NewGuid():N}.xml");

                                    if (attachment.ContentBytes != null)
                                    {
                                        await File.WriteAllBytesAsync(tempFile, attachment.ContentBytes, cancellationToken);
                                        attachmentPaths.Add(tempFile);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // _logger.LogWarning(ex, "Fehler beim Speichern von Attachment {Name}", attachment.Name);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // _logger.LogWarning(ex, "Fehler beim Abrufen der Attachments für Event {Id}", ev.Id);
                }
            }

            return attachmentPaths;
        }

        private (string required, string optional) ParseAttendees(List<Attendee>? attendees)
        {
            if (attendees == null || !attendees.Any())
                return (string.Empty, string.Empty);

            var required = attendees
                .Where(a => a.Type == AttendeeType.Required)
                .Select(a => a.EmailAddress?.Name ?? a.EmailAddress?.Address ?? "Unbekannt");

            var optional = attendees
                .Where(a => a.Type == AttendeeType.Optional)
                .Select(a => a.EmailAddress?.Name ?? a.EmailAddress?.Address ?? "Unbekannt");

            return (
                string.Join("; ", required),
                string.Join("; ", optional)
            );
        }

        private int ConvertFreeBusyStatus(FreeBusyStatus? status)
        {
            return status switch
            {
                FreeBusyStatus.Free => 0,
                FreeBusyStatus.Tentative => 1,
                FreeBusyStatus.Busy => 2,
                FreeBusyStatus.Oof => 3,
                FreeBusyStatus.WorkingElsewhere => 4,
                _ => 2
            };
        }

        private string[] GetEventSelectFields()
        {
            return new[]
            {
                "id",
                "subject",
                "body",
                "bodyPreview",
                "start",
                "end",
                "location",
                "attendees",
                "organizer",
                "createdDateTime",
                "lastModifiedDateTime",
                "isAllDay",
                "showAs",
                "categories",
                "importance",
                "sensitivity",
                "isCancelled",
                "onlineMeeting",
                "recurrence",
                "singleValueExtendedProperties"
            };
        }

        private string GetCurrentTimeZone()
        {
            try
            {
                //return TimeZoneInfo.Local.StandardName;
                return TimeZoneInfo.Local.Id;
            }
            catch
            {
                return "W. Europe Standard Time";
            }
        }

        private string GetContentType(string fileName)
        {
            var extension = Path.GetExtension(fileName).ToLowerInvariant();
            return extension switch
            {
                ".xml" => "application/xml",
                ".txt" => "text/plain",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".png" => "image/png",
                ".pdf" => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream"
            };
        }

        #endregion
    }

    public class OutlookCalendarItemType
    {
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
        public List<string> Attachments { get; set; }
        public string Organizer { get; set; }
        public string RequiredAttendees { get; set; }
        public string OptionalAttendees { get; set; }
        public bool AllDayEvent { get; set; }
        public int BusyStatus { get; set; }
        public string OnlineMeetingUrl { get; set; }
        public bool IsCancelled { get; set; }
        public string RecurrencePattern { get; set; }
        public string Sensitivity { get; set; }
        public int Importance { get; set; }
        public List<string> Categories { get; set; }
        public string CalendarName { get; set; }
        public string Id { get; set; }

        public OutlookCalendarItemType()
        {
            Subject = string.Empty;
            Body = string.Empty;
            BodyPreview = string.Empty;
            CreationTime = DateTime.MinValue;
            Start = DateTime.MinValue;
            End = DateTime.MinValue;
            Duration = 0;
            Location = string.Empty;
            Organizer = string.Empty;
            RequiredAttendees = string.Empty;
            OptionalAttendees = string.Empty;
            Attachments = new List<string>();
            Categories = new List<string>();
            OnlineMeetingUrl = string.Empty;
            RecurrencePattern = string.Empty;
            Sensitivity = "Normal";
            CalendarName = string.Empty;
            ItemId = string.Empty;
            Id = string.Empty;
            // weitere Initialisierungen falls nötig
        }
    }
}