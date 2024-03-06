namespace ohaCalendar
{
    public class OutlookCalendarItemType
    {
        public string Subject;
        public string Body;
        public string EntryID;
        public DateTime CreationTime;
        public DateTime Start;
        public DateTime End_;
        public double Duration;
        public string Location;
        public string Telephone;
        public string Mobile_phone;
        public string Email;
        public string Contact_person;
        public string Companyname1;
        public string Street;
        public string City;
        public string Postcode;
        public string Nation;
        public List<string> Attachments;
        public string Organizer;
        public string RequiredAttendees;
        public string Availability;

        public OutlookCalendarItemType(string subject,
            string body,
            string entryID,
            DateTime creationTime,
            DateTime start,
            DateTime end_,
            double duration,
            string location,
            string telephone,
            string mobile_phone,
            string email,
            string contact_person,
            string companyname1,
            string street,
            string city,
            string postcode,
            string nation,
            List<string> Attachments,
            string Organizer,
            string RequiredAttendees,
            string Availability
            )
        {
            {
                var withBlock = this;
                withBlock.Subject = subject;
                withBlock.Body = body;
                withBlock.EntryID = entryID;
                withBlock.CreationTime = creationTime;
                withBlock.Start = start;
                withBlock.End_ = end_;
                withBlock.Duration = duration;
                withBlock.Location = location;
                withBlock.Telephone = telephone;
                withBlock.Mobile_phone = mobile_phone;
                withBlock.Email = email;
                withBlock.Contact_person = contact_person;
                withBlock.Companyname1 = companyname1;
                withBlock.Street = street;
                withBlock.City = city;
                withBlock.Postcode = postcode;
                withBlock.Nation = nation;
                withBlock.Attachments = Attachments;
                withBlock.Organizer = Organizer;
                withBlock.RequiredAttendees = RequiredAttendees;
                withBlock.Availability = Availability;
            }
        }

        public string DisplayMember
        {
            get
            {
                return Companyname1 + " [" + Email + "]";
            }
        }

        public string ValueMember
        {
            get
            {
                string ret_val;
                Int32 _start;
                Int32 _end;

                if (string.IsNullOrEmpty(Email))
                    return null;

                _start = Email.IndexOf("(") + 1;
                _end = Email.LastIndexOf(")");
                if ((_end > _start))
                    ret_val = Email.Substring(_start, _end - _start);
                else
                    ret_val = null;

                return ret_val;
            }
        }

        public string? Address
        {
            get
            {
                string? ret_val = null;
                ret_val = Street + ", " + Postcode + " " + City + ", " + Nation;
                ret_val = ret_val == ",  , " ? "" : ret_val;
                return ret_val;
            }
        }

        public string? Companyname1_Contact_person
        {
            get
            {
                string? ret_val = null;
                ret_val = !string.IsNullOrEmpty(Companyname1) ? Companyname1 : Contact_person;
                ret_val = ret_val == ",  , " ? "" : ret_val;
                return ret_val;
            }
        }
    }

}
