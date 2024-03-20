using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ohaERP_Library.DateTimePicker
{
    public partial class ohaDateTimePickerMS : System.Windows.Forms.DateTimePicker 
    {
        DateTimeFormatInfo dtfi = CultureInfo.CurrentUICulture.DateTimeFormat;

        public ohaDateTimePickerMS()
        {
            InitializeComponent();

           
        }

        private string ReplaceWith24HourClock(string fmt)
        {
            string pattern = @"^(?<openAMPM>\s*t+\s*)? " +
                             @"(?(openAMPM) h+(?<nonHours>[^ht]+)$ " +
                             @"| \s*h+(?<nonHours>[^ht]+)\s*t+)";
            return Regex.Replace(fmt, pattern, "HH${nonHours}", RegexOptions.IgnorePatternWhitespace);
        }

        private void ohaDateTimePickerMS_Layout(object sender, LayoutEventArgs e)
        {
            dtfi.ShortTimePattern = ReplaceWith24HourClock(dtfi.ShortTimePattern);

            Format = DateTimePickerFormat.Custom;
            dtfi.ShortTimePattern = ReplaceWith24HourClock(dtfi.ShortTimePattern);
            CustomFormat = CultureInfo.CurrentUICulture.DateTimeFormat.ShortDatePattern + "  " +
                           dtfi.ShortTimePattern;
        }
    }
}
