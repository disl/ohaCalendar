using Microsoft.Win32;
using ohaCalendar.Models;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using static ohaCalendar.cOutlook;
using static ohaCalendar.DataSet1;
using Label = System.Windows.Forms.Label;

namespace ohaCalendar
{
    public partial class Calendar : Form
    {
        int g_clientsysid = 1;
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
        int m_active_month_no = 0;
        private CalendarRow? m_selected_calendar;
        private DateTime m_basis_date = DateTime.Today;
        public struct structHolidays
        {
            public structHolidays(DateTime Date, string Description, bool IsPublicHoliday, bool IsHoliday)
            {
                date = Date;
                description = Description;
                is_public_holiday = IsPublicHoliday;
                is_holiday = IsHoliday;
            }
            public DateTime date;
            public string description;
            public bool is_public_holiday;
            public bool is_holiday;
        }
        static List<structHolidays> g_holidays = new List<structHolidays>();
        static List<string> g_weekends = new List<string>();
        private int year_1;
        private int year_4;
        private int m_old_year;
        private string m_holidays_file;
        private string m_current_culture_str;
        private CultureInfo m_current_culture;
        private CultureInfo m_current_culture_default;
        private string m_current_school_holiday;
        private bool m_is_started = true;
        private bool m_is_school_holidays = false;
        private string m_subdivisionCode;
        private int maxRowHeight = 100;
        private rp_staff_jubileeDataTable? m_all_birthdays = null;

        public Calendar()
        {
            InitializeComponent();
        }

        private void calendar_Load(object sender, EventArgs e)
        {
            //rp_staff_jubileeTableAdapter.Connection = new Microsoft.Data.SqlClient.SqlConnection("...");

            try
            {
                int? res_row_count = null;
                m_all_birthdays = rp_staff_jubileeTableAdapter.GetDataByAll(g_clientsysid, g_current_culturesysid);

                ShowProgress(true);

                m_current_culture_default = CultureInfo.CurrentCulture;
                pls_waitLabel.Text = Properties.Resources.please_wait;

                getCalendarItemsToolStripMenuItem.Image = Properties.Resources.refresh_icon;
                moveCalendarToolStripMenuItem.Image = Properties.Resources.dynamic_feed;

                countriesToolStripComboBox.ComboBox.SelectedValueChanged += CountriesToolStripComboBox_SelectedValueChanged;
                stateToolStripComboBox.ComboBox.SelectedValueChanged += StateToolStripComboBox_SelectedValueChanged;
                is_school_holidaysToolStripComboBox.ComboBox.SelectedValueChanged += Is_school_holidaysToolStripComboBox_SelectedValueChanged;

                m_current_culture_str = CultureInfo.CurrentCulture.TwoLetterISOLanguageName.ToUpper();

                FillIs_school_holidays_CMB();
                is_school_holidaysToolStripComboBox.ComboBox.SelectedIndex = 0;
                m_current_school_holiday = is_school_holidaysToolStripComboBox.ComboBox.SelectedValue.ToString().ToUpper();

                FillCountriesCMB();
                ForCalendar_load();
                GetData();

                m_is_started = false;

                countriesToolStripComboBox.ComboBox.SelectedValue = m_current_culture_str;

                rp_staff_jubileeTableAdapter.Fill(calendarDataSet.rp_staff_jubilee, g_clientsysid, g_current_culturesysid, DateTime.Today);
                if (calendarDataSet.rp_staff_jubilee.Rows.Count > 0)
                {
                    SetImageThambnails();
                    ShowHideInfo();
                    tabControl1.SelectedTab = birthdaysTabPage;
                }

                ShowProgress(false);

                //WindowState = FormWindowState.Normal;
            }
            catch (Exception ex)
            {
                DisplayError(this, null, ex,
                    Properties.Settings.Default.LoggingPerEmail,
                    Properties.Resources.msgCaptionError + " in: " + ToString(), g_client_name);
            }
        }

        private void ShowProgress(bool InProgress)
        {
            pls_waitLabel.Visible = InProgress;
            splitContainer1.Visible = !InProgress;
            generalSplitContainer.Visible = !InProgress;
        }

        private void FillIs_school_holidays_CMB()
        {
            DataTable table = new DataTable();
            table.Columns.Add("DisplayMember", typeof(string));
            table.Columns.Add("ValueMember", typeof(string));
            var new_row = table.NewRow();
            new_row["DisplayMember"] = "Public holidays";
            new_row["ValueMember"] = bool.FalseString;
            table.Rows.Add(new_row);
            new_row = table.NewRow();
            new_row["DisplayMember"] = "School holidays";
            new_row["ValueMember"] = bool.TrueString;
            table.Rows.Add(new_row);

            is_school_holidaysToolStripComboBox.ComboBox.DisplayMember = "DisplayMember";
            is_school_holidaysToolStripComboBox.ComboBox.ValueMember = "ValueMember";
            is_school_holidaysToolStripComboBox.ComboBox.DataSource = table;
        }

        private async void Is_school_holidaysToolStripComboBox_SelectedValueChanged(object? sender, EventArgs e)
        {
            m_is_school_holidays = false;
            if (is_school_holidaysToolStripComboBox.ComboBox.SelectedValue != null)
                m_is_school_holidays = is_school_holidaysToolStripComboBox.ComboBox.SelectedValue.ToString() == bool.TrueString;

            if (!m_is_started)
            {
                if (countriesToolStripComboBox.ComboBox.SelectedValue == null
                    || string.IsNullOrEmpty(countriesToolStripComboBox.ComboBox.SelectedValue.ToString())
                    || countriesToolStripComboBox.ComboBox.SelectedValue is DataRowView
                    || stateToolStripComboBox.ComboBox.SelectedValue == null
                    || is_school_holidaysToolStripComboBox.ComboBox.SelectedValue == null
                )
                    return;

                ShowProgress(true);

                var _current_school_holiday = is_school_holidaysToolStripComboBox.ComboBox.SelectedValue.ToString().ToUpper();
                if (_current_school_holiday == null || string.IsNullOrEmpty(_current_school_holiday))
                    return;

                if (_current_school_holiday != m_current_school_holiday)
                    g_holidays = new();

                m_current_school_holiday = _current_school_holiday;

                RefreshHolidays(true);

                CollapseInfo();

                ShowProgress(false);
            }
        }

        private void ForCalendar_load()
        {
            m_holidays_file = Path.Combine(Path.GetTempPath(), "ohaCalendar_holidays.xml");

            ShowHideInfo();
            GetData();

            bodyTextBox.SetSelectionLink(true);
            SetStartup();
        }

        private async void CountriesToolStripComboBox_SelectedValueChanged(object? sender, EventArgs e)
        {
            if (!m_is_started)
            {
                if (countriesToolStripComboBox.ComboBox.SelectedValue != null &&
                    !string.IsNullOrEmpty(countriesToolStripComboBox.ComboBox.SelectedValue.ToString()) &&
                    countriesToolStripComboBox.ComboBox.SelectedValue is string
                   )
                {
                    ShowProgress(true);

                    var _current_culture = countriesToolStripComboBox.ComboBox.SelectedValue.ToString().ToUpper();
                    if (_current_culture == null || string.IsNullOrEmpty(_current_culture))
                        return;

                    if (_current_culture != m_current_culture_str)
                        g_holidays = new();

                    m_current_culture_str = _current_culture;
                    m_current_culture = new CultureInfo(m_current_culture_str);

                    FillStateCMB();
                    RefreshHolidays(true);
                    ForCalendar_load();
                    CollapseInfo();
                    ShowProgress(false);
                }
            }
        }

        private void CollapseInfo()
        {
            generalSplitContainer.Panel2Collapsed = true;
            this.Size = new Size(725, 930);
            close_openToolStripMenuItem.Text = Properties.Resources.calendar_info_show;
            close_openToolStripMenuItem.Image = Properties.Resources.expand;
        }

        private static void SetStartup()
        {
            string StartupKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
            string StartupValue = "ohaCalendar";

            //Set the application to run at startup
            RegistryKey _key = Registry.CurrentUser.OpenSubKey(StartupKey, true);
            if (_key != null)
                _key.SetValue(StartupValue, Application.ExecutablePath.ToString());
        }

        private void GetData(DateTime? actDate = null)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

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

                year_1 = active_date_minus_1.Year;
                year_4 = active_date_plus_2.Year;

                // Refresh holidays
                //await RefreshHolidays(true);
                //StateToolStripComboBox_SelectedValueChanged(null, null);

                tableLayoutPanel1.Controls.Clear();
                tableLayoutPanel2.Controls.Clear();
                tableLayoutPanel3.Controls.Clear();
                tableLayoutPanel4.Controls.Clear();

                var info = new System.Globalization.CultureInfo(g_current_culturesysid);
                m_first_day_of_week = info.DateTimeFormat.FirstDayOfWeek;
                m_list_of_days = GetDaysByWeek(m_basis_date);

                FillCalendarItems();

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

                //year_4 = active_date_plus_2.Year;
                int month_4 = active_date_plus_2.Month;
                string month_4_str = new DateTime(year_4, month_4, 1).ToString("MMMM", CultureInfo.CurrentCulture);
                groupBox4.Text = (month_4_str + " " + year_4).ToUpper();
                m_dateTimes_4 = getAllDates(year_4, month_4);
                FillDataTable(tableLayoutPanel4, m_dateTimes_4, holidays_4Label);
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

        private void DisplayError(Calendar calendar, object value, Exception ex, bool loggingPerEmail, string v, string g_client_name)
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
            List<CalendarItemType> week_list_tmp = new List<CalendarItemType>();
            bool is_my_holiday = false;
            List<DateTime> myholiday_list = new List<DateTime>();

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
                            control = week_list_tmp[col - 1].Active && col <= 6 ? new LinkLabel() : new LinkLabel();
                            holiday = IsHoliday(week_list_tmp[col - 1].Day);
                            is_today = week_list_tmp[col - 1].Day == m_basis_date;


                            var calendar_item = m_calendar_items.FirstOrDefault(
                                     x => x.Start.ToShortDateString() == week_list_tmp[col - 1].Day.ToShortDateString());

                            if (calendar_item != null)
                            {
                                is_my_holiday = calendar_item.AllDayEvent && calendar_item.BusyStatus == (int)OlBusyStatus.olOutOfOffice;

                                if (is_my_holiday)
                                {
                                    var start_day = m_calendar_items.FirstOrDefault(
                                        x => x.Start.ToShortDateString() == week_list_tmp[col - 1].Day.ToShortDateString()).Start;
                                    var end_day = m_calendar_items.FirstOrDefault(
                                        x => x.Start.ToShortDateString() == week_list_tmp[col - 1].Day.ToShortDateString()).End_;

                                    myholiday_list = new List<DateTime>();
                                    for (DateTime d = start_day; d < end_day; d = d.AddDays(1))
                                    {
                                        myholiday_list.Add(d);
                                    }
                                }
                            }
                            // Calendar items
                            var countOfTermins = m_calendar_items.Count(
                                x => x.Start > week_list_tmp[col - 1].Day.AddDays(-1).AddHours(23).AddMinutes(59) && x.Start < week_list_tmp[col - 1].Day.AddDays(1));
                            if (countOfTermins > 0)
                            {
                                week_list_tmp[col - 1].CountOfTermins = countOfTermins;
                                control.Text = week_list_tmp[col - 1].Day.Day.ToString() + GetPotenz(countOfTermins);
                            }
                            else
                                control.Text = week_list_tmp[col - 1].Day.Day.ToString();

                            var has_birthdays = IsContainsDayInBirthdaysArray(week_list_tmp[col - 1].Day);
                            control.Text += has_birthdays ? "\U0001F382" : "   ";

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

                                        char[] trimmer = DateTimeFormatInfo.CurrentInfo.DateSeparator.ToCharArray();
                                        string dateStr = _holiday.date.ToString("d").Replace(_holiday.date.ToString("yyyy"), string.Empty).Trim(trimmer);

                                        holidaysLabel.Text += !string.IsNullOrEmpty(holidaysLabel.Text) ?
                                           ";  " + dateStr + " " + _holiday.description :
                                           dateStr + " " + _holiday.description;
                                    }
                                }
                                else
                                    control.ForeColor = is_my_holiday ? Color.Green : week_list_tmp[col - 1].Active ? Color.Red : Color.Gray;
                            }
                            else if (col > 0 && col <= 5)
                            {
                                if (control is LinkLabel)
                                {
                                    if (Is_my_holiday(control, myholiday_list))
                                    {
                                        var label = new Label();
                                        label.Text = control.Text;
                                        if (week_list_tmp[col - 1].Active)
                                        {
                                            label.ForeColor = Color.LimeGreen;
                                            label.Font = new Font(label.Font, FontStyle.Bold);
                                            label.Tag = week_list_tmp[col - 1];
                                            label.Click += Label_Click;
                                            label.Cursor = Cursors.Hand;
                                        }
                                        control = label;
                                    }
                                }
                                else
                                    control.ForeColor = Is_my_holiday(control, myholiday_list) ? Color.Green : week_list_tmp[col - 1].Active ? Color.Red : Color.Gray;
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
                                    label.ForeColor = is_my_holiday ? Color.Green : Color.Red;
                                    label.Tag = week_list_tmp[col - 1];
                                    label.Click += Label_Click;
                                    label.Cursor = Cursors.Hand;
                                }
                                control = label;
                            }
                            else
                            {
                                if (col > 0 && week_list_tmp[col - 1].Active)
                                    control.ForeColor = is_my_holiday ? Color.Green : Color.Red;
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

        private bool Is_my_holiday(Control control, List<DateTime> myholiday_list)
        {
            bool ret_val = false;

            if (control.Tag == null)
                ret_val = false;

            var tag = control.Tag as CalendarItemType;

            ret_val = myholiday_list.Contains(tag.Day);

            return ret_val;
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

                rp_staff_jubileeTableAdapter.Fill(calendarDataSet.rp_staff_jubilee, g_clientsysid, g_current_culturesysid, (DateTime)SelectedDay);  //, "G", ref res_row_count);

                startTextBox.Text = ((DateTime)SelectedDay).ToLongDateString();

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

                    if (m_calendar_items != null && m_calendar_items.Count > 0)
                    {
                        foreach (OutlookCalendarItemType item in m_calendar_items)
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
                                item.EntryID,
                                item.AllDayEvent,
                                item.BusyStatus
                                );
                        }
                    }
                    if (generalSplitContainer.Panel2Collapsed)
                        ShowHideInfo();

                    SetImageThambnails();
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

        private void SetImageThambnails()
        {

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                Image img = null;

                if (row.Cells[staff_image.Name].Value == DBNull.Value)
                {
                    continue;
                }

                try
                {
                    using (var ms = new MemoryStream((byte[])row.Cells[staff_image.Name].Value))
                    {
                        img = Image.FromStream(ms);
                    }
                }
                catch (Exception ex)
                {

                }
                if (img != null)
                {
                    var image_height = 200;
                    var image_width = image_height * img.Width / img.Height;
                    var m_thumbnail = (Bitmap)img.GetThumbnailImage(image_width, image_height, null, IntPtr.Zero);
                    if (m_thumbnail != null)
                        row.Cells[staff_thumbnail.Name].Value = m_thumbnail;
                    else
                        row.Cells[staff_thumbnail.Name].Value = img;
                }
            }
            staff_image.Visible = false;
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
                this.Size = new Size(725, 930);  //(697, 930);
                close_openToolStripMenuItem.Text = Properties.Resources.calendar_info_show;
                close_openToolStripMenuItem.Image = Properties.Resources.expand;
            }
            else
            {
                this.Size = new Size(1400, 930);
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

        private void checklist_for_release_answersBindingSource_CurrentChanged(object sender, EventArgs e)
        {
            m_selected_calendar = null;
            if (checklist_for_release_answersBindingSource.Current != null)
                m_selected_calendar = ((DataRowView)checklist_for_release_answersBindingSource.Current).Row as CalendarRow;
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

        private void richTextBoxEx1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBoxEx1_LinkClicked(object sender, LinkClickedEventArgs e)
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

        private void FillHolidayArray(int year, bool DeleteOld = true, bool DoAlways = false)
        {
            List<HolidayType?> m_holidays = null;

            if (!DoAlways || (g_holidays.Count > 0 && year == m_old_year))
                return;

            Cursor = Cursors.WaitCursor;

            try
            {
                if (DeleteOld)
                    g_holidays = new List<structHolidays>();

                using (HttpClient wc = new())
                {
                    string? url = null;
                    if (m_is_school_holidays)
                    {
                        if (stateToolStripComboBox.Visible && stateToolStripComboBox.ComboBox.SelectedValue != null)
                            url = "https://openholidaysapi.org/SchoolHolidays?countryIsoCode=" + m_current_culture_str +
                               "&languageIsoCode=" + m_current_culture_str +
                               "&subdivisionCode=" + m_current_culture_str + "-" + stateToolStripComboBox.ComboBox.SelectedValue.ToString().ToUpper() +
                               "&validFrom=" + year + "-01-01" +
                               "&validTo=" + year + "-12-31";
                        else
                            url = "https://openholidaysapi.org/SchoolHolidays?countryIsoCode=" + m_current_culture_str +
                               "&languageIsoCode=" + m_current_culture_str +
                               "&validFrom=" + year + "-01-01" +
                               "&validTo=" + year + "-12-31";
                    }
                    else
                        url = "https://openholidaysapi.org/PublicHolidays?countryIsoCode=" + m_current_culture_str +
                            "&languageIsoCode=" + m_current_culture_str +
                            "&validFrom=" + year +
                            "-01-01&validTo=" + year + "-12-31";

                    var json = wc.GetStringAsync(url).Result;
                    if (json != null)
                        m_holidays = JsonSerializer.Deserialize<List<HolidayType>>(json);
                }
                if (m_holidays != null)
                {
                    foreach (var holiday in m_holidays)
                    {
                        structHolidays new_item = default;
                        new_item.date = Convert.ToDateTime(holiday.startDate);

                        Name? name_obj = holiday.name.FirstOrDefault(x => x.language == m_current_culture_str);
                        if (name_obj == null)
                        {
                            name_obj = holiday.name.FirstOrDefault(x => x.language == "EN");
                        }
                        if (name_obj == null)
                        {
                            return;
                        }
                        var holiday_name = name_obj.text;
                        new_item.description = holiday_name != null ? holiday_name : "";
                        new_item.is_public_holiday = true;
                        new_item.is_holiday = false;

                        g_holidays.Add(new_item);
                    }

                    GetData();
                }
                m_old_year = year;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void FillCountriesCMB()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DisplayMember", typeof(string));
            dt.Columns.Add("ValueMember", typeof(string));

            List<CountriesType> countriesList = new List<CountriesType>();

            using (HttpClient wc = new HttpClient())
            {
                var url = "https://openholidaysapi.org/Countries";
                var json = wc.GetStringAsync(url).Result;
                if (json != null)
                    countriesList = JsonSerializer.Deserialize<List<CountriesType>>(json);
            }

            AddRow(dt, "<countries>", "");
            if (countriesList != null)
            {
                var countriesList_by_isocode = countriesList.Where(x => x.isoCode != null && !string.IsNullOrEmpty(x.isoCode)).ToList();

                if (countriesList_by_isocode == null || countriesList_by_isocode.Count() == 0)
                {
                    countriesToolStripComboBox.Visible = false;
                    return;
                }
                else
                    countriesToolStripComboBox.Visible = true;

                foreach (var item in countriesList_by_isocode)
                {
                    var name_obj = item.name.FirstOrDefault(x => x.language == m_current_culture_str);
                    if (name_obj == null)
                    {
                        name_obj = item.name.FirstOrDefault(x => x.language == "EN");
                    }
                    if (name_obj == null)
                    {
                        countriesToolStripComboBox.Visible = false;
                        return;
                    }
                    var displayMember = name_obj.text;
                    AddRow(dt, displayMember, item.isoCode);
                }
                countriesToolStripComboBox.ComboBox.DataSource = dt;
                countriesToolStripComboBox.ComboBox.DisplayMember = "DisplayMember";
                countriesToolStripComboBox.ComboBox.ValueMember = "ValueMember";
            }
        }

        private void FillStateCMB()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DisplayMember", typeof(string));
            dt.Columns.Add("ValueMember", typeof(string));

            List<FederalStateType> federalStateList = new List<FederalStateType>();

            stateToolStripComboBox.Visible = false;

            using (HttpClient wc = new HttpClient())
            {
                var url = "https://openholidaysapi.org/Subdivisions?countryIsoCode=" + m_current_culture_str;
                var json = wc.GetStringAsync(url).Result;
                if (json != null)
                {
                    federalStateList = JsonSerializer.Deserialize<List<FederalStateType>>(json);
                }
            }

            AddRow(dt, "<federal states>", "");
            if (federalStateList != null)
            {
                var federalStateList_by_isocode = federalStateList.Where(x => !string.IsNullOrEmpty(x.shortName)).ToList(); // !string.IsNullOrEmpty(x.isoCode));

                if (federalStateList_by_isocode == null || federalStateList_by_isocode.Count() == 0)
                {
                    stateToolStripComboBox.Visible = false;
                    return;
                }
                else
                {
                    stateToolStripComboBox.Visible = true;
                }

                foreach (var item in federalStateList_by_isocode)
                {
                    var name_obj = item.name.FirstOrDefault(x => x.language == m_current_culture_str);
                    if (name_obj == null)
                    {
                        name_obj = item.name.FirstOrDefault(x => x.language == "EN");
                    }
                    if (name_obj == null)
                    {
                        stateToolStripComboBox.Visible = false;
                        return;
                    }
                    var displayMember = name_obj.text;
                    AddRow(dt, displayMember, item.shortName);
                }
                stateToolStripComboBox.ComboBox.DataSource = dt;
                stateToolStripComboBox.ComboBox.DisplayMember = "DisplayMember";
                stateToolStripComboBox.ComboBox.ValueMember = "ValueMember";
                stateToolStripComboBox.ComboBox.SelectedIndex = 1;

                stateToolStripComboBox.Visible = true;
                //is_school_holidaysToolStripComboBox.Visible = true;

                if (!string.IsNullOrEmpty(Properties.Settings.Default.state))
                    stateToolStripComboBox.ComboBox.SelectedValue = Properties.Settings.Default.state;
            }

        }

        private static void AddRow(DataTable dt, string displayMember, string valueMember)
        {
            var new_row = dt.NewRow();
            new_row["DisplayMember"] = displayMember;
            new_row["ValueMember"] = valueMember;
            dt.Rows.Add(new_row);
        }

        private void calendar_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (stateToolStripComboBox.ComboBox.SelectedValue != null && !string.IsNullOrEmpty(stateToolStripComboBox.ComboBox.SelectedValue.ToString()))
            {
                Properties.Settings.Default.state = stateToolStripComboBox.ComboBox.SelectedValue.ToString();
                Properties.Settings.Default.Save();
            }
        }

        private async void StateToolStripComboBox_SelectedValueChanged(object? sender, EventArgs e)
        {
            m_subdivisionCode = null;

            if (!m_is_started)
            {
                if (countriesToolStripComboBox.ComboBox.SelectedValue == null
                    || string.IsNullOrEmpty(countriesToolStripComboBox.ComboBox.SelectedValue.ToString())
                    || countriesToolStripComboBox.ComboBox.SelectedValue is DataRowView
                    || stateToolStripComboBox.ComboBox.SelectedValue == null
                )
                    return;

                ShowProgress(true);

                var _subdivisionCode = stateToolStripComboBox.ComboBox.SelectedValue.ToString().ToUpper();

                if (_subdivisionCode != m_subdivisionCode)
                    g_holidays = new();

                m_subdivisionCode = _subdivisionCode;

                RefreshHolidays(true);

                CollapseInfo();

                ShowProgress(false);
            }
        }

        private void RefreshHolidays(bool doAlways)
        {
            if (year_1 > 0)
            {
                FillHolidayArray(year_1, DoAlways: doAlways);
                if (year_4 > year_1)
                    FillHolidayArray(year_4, false, DoAlways: doAlways);
            }
        }

        private bool IsContainsDayInBirthdaysArray(DateTime testDay)
        {
            if (m_all_birthdays == null || m_all_birthdays.Count == 0)
                return false;

            return m_all_birthdays.Count(x => x.birthday.Day == testDay.Day && x.birthday.Month == testDay.Month) > 0;
        }

    }


}
