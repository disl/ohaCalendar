using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using static ohaCalendar.CalendarDataSet;

//using static ohaCalendar.CalendarDataSet;
using Label = System.Windows.Forms.Label;

namespace ohaCalendar
{
    public partial class calendar : Form
    {
        int g_current_culturesysid = 9;
        string g_current_culture = "de-DE";
        private string g_client_name = "Otto Haas KG";

        int m_count_of_months = 4;
        List<CalendarItemType> m_dateTimes_1 = new();
        List<CalendarItemType> m_dateTimes_2 = new();
        List<CalendarItemType> m_dateTimes_3 = new();
        List<CalendarItemType> m_dateTimes_4 = new();
        private DayOfWeek m_first_day_of_week;
        private List<DateTime> m_list_of_days;
        public DateTime? SelectedDay;
        private bool ShowOutlook = true;
        private List<OutlookCalendarItemType> m_calendar_items;
        //BindingSource m_bindingSource = new BindingSource();
        int m_active_month_no = 0;
        private CalendarRow? m_selected_calendar;
        private DateTime m_basis_date = DateTime.Today;
        public struct structHolidays
        {
            public structHolidays(DateTime Date, string Description)
            {
                date = Date;
                description = Description;
            }

            public DateTime date;
            public string description;
        }
        static List<structHolidays> g_holidays = new List<structHolidays>();
        static List<string> g_weekends = new List<string>();

        public calendar()
        {
            InitializeComponent();
        }

        private void calendar_Load(object sender, EventArgs e)
        {
            try
            {
                getCalendarItemsToolStripMenuItem.Image = Properties.Resources.refresh_icon;
                moveCalendarToolStripMenuItem.Image = Properties.Resources.dynamic_feed;

                ShowHideInfo();

                GetData();

                bodyTextBox.SetSelectionLink(true);

                SetStartup();
            }
            catch (Exception ex)
            {
                DisplayError(this, null, ex,
                    Properties.Settings.Default.LoggingPerEmail,
                    Properties.Resources.msgCaptionError + " in: " + ToString(), g_client_name);
            }
        }

        private static void SetStartup()
        {
            string StartupKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
            string StartupValue = "ohaCalendar";

            //Set the application to run at startup
            RegistryKey key = Registry.CurrentUser.OpenSubKey(StartupKey, true);
            key.SetValue(StartupValue, Application.ExecutablePath.ToString());
        }

        private void GetData(DateTime? actDate = null)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                System.Windows.Forms.Application.DoEvents();

                DateTime active_date = m_basis_date.AddMonths(m_active_month_no);
                if (actDate != null)
                {
                    m_active_month_no = 0;
                    m_basis_date = (DateTime)actDate;
                    active_date = (DateTime)actDate;
                }

                DateTime active_date_minus_1 = active_date.AddMonths(-1);
                DateTime active_date_plus_1 = active_date.AddMonths(1);
                DateTime active_date_plus_2 = active_date.AddMonths(2);

                tableLayoutPanel1.Controls.Clear();
                tableLayoutPanel2.Controls.Clear();
                tableLayoutPanel3.Controls.Clear();
                tableLayoutPanel4.Controls.Clear();

                pls_waitLabel.Text = Properties.Resources.please_wait;
                pls_waitLabel.Visible = true;
                splitContainer1.Visible = false;
                generalSplitContainer.Visible = false;

                var info = new System.Globalization.CultureInfo(g_current_culturesysid);
                m_first_day_of_week = info.DateTimeFormat.FirstDayOfWeek;
                m_list_of_days = GetDaysByWeek(m_basis_date);

                FillCalendarItems();

                int year_1 = active_date_minus_1.Year;
                int month_1 = active_date_minus_1.Month;
                string month_1_str = new DateTime(year_1, month_1, 1).ToString("MMMM", CultureInfo.CurrentCulture);
                groupBox1.Text = (month_1_str + " " + year_1).ToUpper();
                m_dateTimes_1 = getAllDates(year_1, month_1);
                FillDataTable(tableLayoutPanel1, m_dateTimes_1, holidays_1Label);

                int year_2 = active_date.Year;
                int month_2 = active_date.Month;
                string month_2_str = new DateTime(year_2, month_2, 1).ToString("MMMM", CultureInfo.CurrentCulture);
                groupBox2.Text = (month_2_str + " " + year_2).ToUpper();
                m_dateTimes_2 = getAllDates(year_2, month_2);
                FillDataTable(tableLayoutPanel2, m_dateTimes_2, holidays_2Label);

                int year_3 = active_date_plus_1.Year;
                int month_3 = active_date_plus_1.Month;
                string month_3_str = new DateTime(year_3, month_3, 1).ToString("MMMM", CultureInfo.CurrentCulture);
                groupBox3.Text = (month_3_str + " " + year_3).ToUpper();
                m_dateTimes_3 = getAllDates(year_3, month_3);
                FillDataTable(tableLayoutPanel3, m_dateTimes_3, holidays_3Label);

                int year_4 = active_date_plus_2.Year;
                int month_4 = active_date_plus_2.Month;
                string month_4_str = new DateTime(year_4, month_4, 1).ToString("MMMM", CultureInfo.CurrentCulture);
                groupBox4.Text = (month_4_str + " " + year_4).ToUpper();
                m_dateTimes_4 = getAllDates(year_4, month_4);
                FillDataTable(tableLayoutPanel4, m_dateTimes_4, holidays_4Label);

                pls_waitLabel.Visible = false;
                splitContainer1.Visible = true;
                generalSplitContainer.Visible = true;

            }
            catch (Exception ex)
            {
                DisplayError(this, null, ex,
                     Properties.Settings.Default.LoggingPerEmail,
                     Properties.Resources.msgCaptionError + " in: " + ToString(), g_client_name);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void DisplayError(calendar calendar, object value, Exception ex, bool loggingPerEmail, string v, string g_client_name)
        {
            MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
        }

        private void FillCalendarItems()
        {
            var calendar = cOutlook.GetCurrentCalendar();
            m_calendar_items = cOutlook.Outlook_GetCalendarItems(
                calendar,
                null, //m_calendar_name,
                null,
                null,
                m_basis_date.AddMonths(m_active_month_no).AddMonths(-1).AddDays(-1),
                m_basis_date.AddMonths(m_active_month_no).AddMonths(2).AddDays(1),
                false,
                IsLightAlgorithmus: true);
        }

        private void FillDataTable(TableLayoutPanel panel, List<CalendarItemType> date_list, Label holidaysLabel)
        {
            int row = 0, col = 0;
            holidaysLabel.Text = string.Empty;
            panel.ControlAdded += Panel_ControlAdded;
            structHolidays? holiday = null;

            try
            {
                var first_day = date_list.First().Day;
                // Days before first day
                var missing_days_before_counter = 7 - date_list.Count(x => x.Week_no == date_list.Min(y => y.Week_no));
                List<CalendarItemType> day_list_missing_days_before = new List<CalendarItemType>();
                for (int i = missing_days_before_counter; i > 0; i--)
                {
                    day_list_missing_days_before.Add(new CalendarItemType(first_day.AddDays((-1) * i), date_list[0].Week_no, false, 0));
                }
                // Days after last day
                var last_day = date_list.Last().Day;
                var missing_days_after_counter = 7 - date_list.Count(x => x.Week_no == date_list.Max(y => y.Week_no));
                List<CalendarItemType> day_list_missing_days_after = new List<CalendarItemType>();
                for (int i = 1; i <= missing_days_after_counter; i++)
                {
                    day_list_missing_days_after.Add(
                        new CalendarItemType(last_day.AddDays(i), date_list.First(x => x.Week_no == date_list.Max(y => y.Week_no)).Week_no, false, 0));
                }
                // Full list
                List<CalendarItemType> day_list_full = new List<CalendarItemType>();
                day_list_full.AddRange(day_list_missing_days_before);
                day_list_full.AddRange(date_list);
                day_list_full.AddRange(day_list_missing_days_after);

                var weeks_list = date_list.Select(x => x.Week_no).Distinct().ToList();

                List<CalendarItemType> week_list_tmp = new List<CalendarItemType>();

                bool is_today = false;
                for (row = 0; row < 6; row++)
                {
                    for (col = 0; col < 8; col++)
                    {
                        System.Windows.Forms.Control control = new System.Windows.Forms.Control();

                        control = row == 0 || col == 0 ? new Label() : new LinkLabel();
                        if (row == 0)
                        {
                            if (col == 0)
                            {
                                // Wort "Week" ("Woche")
                                control.Text = Properties.Resources.week;
                            }
                            // Week days
                            else if (col > 0 && col <= 7)
                            {
                                control.Text = m_list_of_days[col - 1].ToString("dddd");
                            }
                        }
                        else if (col == 0 && row > 0)
                        {
                            // Weeks-no.
                            control.Text = weeks_list[row - 1].ToString();
                        }
                        else
                        {
                            // Days
                            if (row < 6)
                                week_list_tmp = day_list_full.Take(new Range((row - 1) * 7, ((row - 1) * 7 + 7))).ToList();
                            control = week_list_tmp[col - 1].Active && col <= 6 ? new LinkLabel() : new Label();
                            holiday = IsHoliday(week_list_tmp[col - 1].Day);
                            is_today = week_list_tmp[col - 1].Day == m_basis_date;

                            // Calendar items
                            var calendar_items = m_calendar_items.Count(
                                x => x.Start > week_list_tmp[col - 1].Day.AddDays(-1).AddHours(23).AddMinutes(59) && x.Start < week_list_tmp[col - 1].Day.AddDays(1));
                            if (calendar_items > 0)
                            {
                                week_list_tmp[col - 1].CountOfTermins = calendar_items;

                                //if (control is LinkLabel)
                                control.Text = week_list_tmp[col - 1].Day.Day.ToString() + GetPotenz(calendar_items); // + calendar_items ;
                                //else
                                //    control.Text = week_list_tmp[col - 1].Day.Day.ToString();
                            }
                            else
                                control.Text = week_list_tmp[col - 1].Day.Day.ToString();
                            control.Tag = week_list_tmp[col - 1];
                            if (control is LinkLabel)
                            {
                                ((LinkLabel)control).LinkClicked += Calendar_LinkClicked;
                            }
                            // Holidays & Weekends
                            if (col >= 6 || (holiday != null && !string.IsNullOrEmpty(((structHolidays)holiday).description)))
                            {
                                if (control is LinkLabel)
                                {
                                    if (holiday != null)
                                    {
                                        var _holiday = (structHolidays)holiday;
                                        toolTip1.SetToolTip(control, _holiday.description);
                                        holidaysLabel.Text += !string.IsNullOrEmpty(holidaysLabel.Text) ?
                                            ";  " + _holiday.date.ToShortDateString() + " " + _holiday.description :
                                            _holiday.date.ToShortDateString() + " " + _holiday.description;
                                    }
                                }
                                else
                                    control.ForeColor = week_list_tmp[col - 1].Active ? Color.Red : Color.Gray;
                            }
                            else if (col > 0 && col <= 5)
                            {
                                if (control is LinkLabel)
                                    ((LinkLabel)control).LinkColor = week_list_tmp[col - 1].Active ? Color.Red : Color.Gray;
                                else
                                    control.ForeColor = Color.Gray;
                            }
                        }
                        control.Font = row == 0 || col == 0 ? new Font("Arial", 9) : new Font("Arial", 14);
                        // Today
                        if (is_today)
                        {
                            if (control is LinkLabel)
                                control.BackColor = Color.Gold;
                            toolTip1.SetToolTip(control, Properties.Resources.today);
                        }
                        // Holiday (because of colours)
                        if (holiday != null)
                        {
                            if (control is LinkLabel)
                            {
                                var label = new Label();
                                label.Text = control.Text;
                                if (week_list_tmp[col - 1].Active)
                                {
                                    label.ForeColor = Color.Red;
                                    label.Tag = week_list_tmp[col - 1];
                                    label.Click += Label_Click;
                                    label.Cursor = Cursors.Hand;
                                }
                                control = label;
                            }
                            else
                            {
                                if (week_list_tmp[col - 1].Active)
                                    control.ForeColor = Color.Red;
                            }
                        }

                        panel.Controls.Add(control, col, row);
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayError(this, "Row=" + row + ", Col=" + col, ex,
                    Properties.Settings.Default.LoggingPerEmail,
                    Properties.Resources.msgCaptionError + " in: " + ToString(), g_client_name);
            }
        }

        private void Label_Click(object? sender, EventArgs e)
        {
            ForLinkClicked(((Label)sender).Tag as CalendarItemType);
        }

        private string? GetPotenz(int CountOfTermins)
        {
            //const string SuperscriptDigits = "\u2070\u00b9\u00b2\u00b3\u2074\u2075\u2076\u2077\u2078\u2079";
            //string superscript = new string(text.Select(x => SuperscriptDigits[x - '0']).ToArray());

            switch (CountOfTermins)
            {
                case 0: return "\u2070";
                case 1: return "\u00b9";
                case 2: return "\u00b2";
                case 3: return "\u00b3";
                case 4: return "\u2074";
                case 5: return "\u2075";
                case 6: return "\u2076";
                case 7: return "\u2077";
                case 8: return "\u2078";
                case 9: return "\u2079";
            }
            return null;
        }

        private void Panel_ControlAdded(object? sender, ControlEventArgs e)
        {
            var control = e.Control is Label ? e.Control as Label : e.Control as LinkLabel;
            if (control != null)
            {
                if (control is LinkLabel)
                {
                    ((LinkLabel)control).LinkColor = Color.Black;
                    ((LinkLabel)control).LinkBehavior = control.Text.Contains("*") ? LinkBehavior.AlwaysUnderline : LinkBehavior.NeverUnderline;
                    if (control.Text.Contains("*"))
                    {
                        control.Text = control.Text.Replace("*", " ");
                    }
                }
                else
                {
                    if (control.Text.Contains("*"))
                    {
                        control.Text = control.Text.Replace("*", " ");
                    }
                }
                control.TextAlign = ContentAlignment.MiddleLeft;
                control.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
                control.TextAlign = ContentAlignment.MiddleCenter;
                control.AutoSize = false;
            }
        }

        private structHolidays? IsHoliday(DateTime date)
        {
            structHolidays? ret_val = null;
            ret_val = g_holidays.FirstOrDefault(x => x.date == date && !string.IsNullOrEmpty(x.description) && x.date != DateTime.MinValue);
            return ((structHolidays)ret_val).date != DateTime.MinValue ? (structHolidays)ret_val : null;
        }

        private void Calendar_LinkClicked(object? sender, LinkLabelLinkClickedEventArgs e)
        {
            if (sender != null)
                ForLinkClicked(((LinkLabel)sender).Tag as CalendarItemType);
        }

        private void ForLinkClicked(CalendarItemType SelectedItem)
        {
            try
            {
                calendarDataSet.Calendar.Rows.Clear();
                var sel_obj = SelectedItem; //((LinkLabel)sender).Tag as CalendarItemType;
                SelectedDay = sel_obj.Day;

                dateTextBox.Text = ((DateTime)SelectedDay).ToLongDateString();

                if (ShowOutlook)
                {
                    var calendar = cOutlook.GetCurrentCalendar();
                    m_calendar_items = cOutlook.Outlook_GetCalendarItems(
                        calendar,
                        null, //m_calendar_name,
                        null,
                        null,
                        ((DateTime)SelectedDay).AddDays(-1).AddHours(23),
                        ((DateTime)SelectedDay).AddDays(1).AddMinutes(1),
                        false,
                        IsLightAlgorithmus: true
                        );

                    Cursor = Cursors.WaitCursor;

                    if (m_calendar_items == null || m_calendar_items.Count == 0)
                    {
                        MessageBox.Show(Properties.Resources.No_dates_found_in_the_calendar);
                        return;
                    }
                    else
                    {
                        foreach (var item in m_calendar_items)
                        {
                            calendarDataSet.Calendar.AddCalendarRow(
                                item.Subject,
                                item.Start,
                                Convert.ToInt32(item.Duration),
                                item.Location,
                                item.Body,
                                item.End_,
                                item.Organizer,
                                item.RequiredAttendees,
                                item.EntryID
                                );
                        }
                    }
                    if (generalSplitContainer.Panel2Collapsed)
                        ShowHideInfo();
                }
            }
            catch (Exception ex)
            {
                DisplayError(this, null, ex,
                    Properties.Settings.Default.LoggingPerEmail,
                    Properties.Resources.msgCaptionError + " in: " + ToString(), g_client_name);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public int GetWeekNumber(DateTime dt)
        {
            CultureInfo curr = CultureInfo.CurrentCulture;
            var info = new CultureInfo(g_current_culture);
            var first_day_of_week = info.DateTimeFormat.FirstDayOfWeek;
            int week = curr.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstFourDayWeek, first_day_of_week);
            return week;
        }

        private List<DateTime> GetDaysByWeek(DateTime act_day)
        {
            int currentDayOfWeek = (int)act_day.DayOfWeek;
            DateTime sunday = act_day.AddDays(-currentDayOfWeek);
            DateTime monday = sunday.AddDays(1);
            // If we started on Sunday, we should actually have gone *back*
            // 6 days instead of forward 1...
            if (currentDayOfWeek == 0)
            {
                monday = monday.AddDays(-7);
            }
            var dates = Enumerable.Range(0, 7).Select(d => monday.AddDays(d)).ToList();
            return dates;
        }

        public List<CalendarItemType> getAllDates(int year, int month)
        {
            var ret = new List<CalendarItemType>();
            for (int i = 1; i <= DateTime.DaysInMonth(year, month); i++)
            {
                ret.Add(new CalendarItemType(new DateTime(year, month, i), GetWeekNumber(new DateTime(year, month, i)), true, 0));
            }
            return ret;
        }

        private void getCalendarItemsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GetData();
        }

        private void close_openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowHideInfo();
        }

        private void ShowHideInfo()
        {
            generalSplitContainer.Panel2Collapsed = !generalSplitContainer.Panel2Collapsed;
            if (generalSplitContainer.Panel2Collapsed)
            {
                this.Size = new Size(620, 930);  //(697, 930);
                close_openToolStripMenuItem.Text = Properties.Resources.calendar_info_show;
                close_openToolStripMenuItem.Image = Properties.Resources.expand;
            }
            else
            {
                this.Size = new Size(1394, 930);
                close_openToolStripMenuItem.Text = Properties.Resources.calendar_info_hidden;
                close_openToolStripMenuItem.Image = Properties.Resources.collapse;
            }
        }

        private void move_calendar_backToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_active_month_no--;
            GetData();
        }

        private void move_calendar_forwardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_active_month_no++;
            GetData();
        }

        private void move_calendar_currentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_basis_date = DateTime.Today;
            m_active_month_no = 0;
            GetData();
        }

        private void bodyTextBox_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(e.LinkText))
                {

                    var ps = new ProcessStartInfo(e.LinkText)
                    {
                        UseShellExecute = true,
                        Verb = "open"
                    };
                    Process.Start(ps);
                }
            }
            catch (Win32Exception)
            {
                Process.Start("IExplore.exe", "http://myurl");
            }
        }

        private void bodyTextBox_TextChanged(object sender, EventArgs e)
        {
            //RichTextBox rtb = (RichTextBox)sender;
        }

        private void checklist_for_release_answersBindingSource_CurrentChanged(object sender, EventArgs e)
        {
            m_selected_calendar = null;
            if (checklist_for_release_answersBindingSource.Current != null)
                m_selected_calendar = ((DataRowView)checklist_for_release_answersBindingSource.Current).Row as CalendarRow;

            // ((DataRowView)bindingSource.Current
            //   ((DataRowView)bindingSource.Current).Row.Table.Columns.Contains(FieldName) &&
        }

        private void dataGridView1_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ShowItemInOutlook();
        }

        private void ShowItemInOutlook()
        {
            if (m_selected_calendar != null && !m_selected_calendar.IsEntryIDNull())
            {
                var item = cOutlook.GetAndOpenAppointment(m_selected_calendar.EntryID, null, null);
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex == outlookColumn.Index)
            {
                ShowItemInOutlook();
            }
        }

        private void byDateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InputForm_date inputForm_Date = new InputForm_date();
            if (inputForm_Date.ShowDialog() == DialogResult.OK)
            {
                GetData(inputForm_Date.Value);
            }
        }
    }

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
