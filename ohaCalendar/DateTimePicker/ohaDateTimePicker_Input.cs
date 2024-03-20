using System.Globalization;

namespace ohaERP_Library.DateTimePicker
{
    public partial class ohaDateTimePicker_Input : Form
    {
        #region LOAD's

        CultureInfo m_CultureInfo = CultureInfo.CurrentCulture;

        string m_string;
        string String { get { return m_string; } }

        private string m_value = "";
        public string Value { get { return m_value; } }

        private string m_DateSeparator;
        private string m_ShortDatePattern;
        private string m_today = "t";
        private string m_month = "m";
        private string m_week = "w";
        private string m_year = "y";
        private Control m_control;
        int m_location_x;
        int m_location_y;
        DateTime? m_date;

        public ohaDateTimePicker_Input()
        {
            InitializeComponent();

            SetFormat();
        }

        public ohaDateTimePicker_Input(
           string String,
           Control control,
           DateTime? Date)
        {
            try
            {
                InitializeComponent();

                m_string = String;

                m_location_x = control.PointToScreen(new Point(0, 0)).X - 3;
                m_location_y = control.PointToScreen(new Point(0, 0)).Y - Size.Height - 3;
                m_control = control;
                m_date = Date;
                dateTimePicker1.Value =  m_date;

                SetFormat();
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private void ohaDateTimePicker_Input_Load(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(m_string))
                {
                    Location = new Point(m_location_x, m_location_y);
                    manualTextBox.Focus();
                    manualTextBox.Text = m_string;
                    if (!string.IsNullOrEmpty(m_string))
                        manualTextBox.Select(1, 1);
                }
                else
                {
                    DialogResult = DialogResult.Cancel;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
                Close();
            }
        }

        #endregion

        #region manualTextBox

        private void manualTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(e.KeyChar.ToString() == "0" || e.KeyChar.ToString() == "1" ||
                e.KeyChar.ToString() == "2" || e.KeyChar.ToString() == "3" ||
                e.KeyChar.ToString() == "4" || e.KeyChar.ToString() == "5" ||
                e.KeyChar.ToString() == "6" || e.KeyChar.ToString() == "7" ||
                e.KeyChar.ToString() == "8" || e.KeyChar.ToString() == "9" ||
                e.KeyChar.ToString() == "." || e.KeyChar.ToString() == "\b" ||
                e.KeyChar.ToString().ToLower() == m_today || e.KeyChar.ToString() == "-" ||
                e.KeyChar.ToString() == "+" || e.KeyChar.ToString() == "0" || e.KeyChar.ToString() == m_today ||
                e.KeyChar.ToString().ToLower() == m_month || e.KeyChar.ToString().ToLower() == m_week || e.KeyChar.ToString() == "/"))

                e.Handled = true;
        }

        private void manualTextBox_TextChanged(object sender, EventArgs e)
        {
            //if (m_date != null)
            //    dateTimePicker1.Value = m_date;

            if (!string.IsNullOrEmpty(manualTextBox.Text))
                DateBilder();
            else
                resultLabel.Text = string.Empty;
        }

        private void DateBilder()
        {
            DateTime oDateTime;
            string strDate = "";
            int anz = 0;

            stateTextBox.Text = string.Empty;

            if (manualTextBox.Text == "")
            {
                SetDate("");

            }
            else if (manualTextBox.Text.ToLower() == m_today)
            {
                SetDate(DateTime.Today.ToShortDateString());
            }
            else if (manualTextBox.Text.Length == 2)
            {
                switch (m_ShortDatePattern)
                {
                    case "dd.MM.yyyy":        // de-DE                       
                    case "dd/MM/yyyy":        // en-GB 
                        strDate = manualTextBox.Text.Substring(0, 1) + m_DateSeparator +
                                  manualTextBox.Text.Substring(1, 1) + m_DateSeparator +
                                  DateTime.Today.Year;
                        break;
                    case "yyyy. MM. dd.":     // hu-HU       
                        strDate = manualTextBox.Text.Substring(0, 4) + m_DateSeparator + " " +
                                  manualTextBox.Text.Substring(4, 2) + m_DateSeparator + " " +
                                  manualTextBox.Text.Substring(6, 2) + m_DateSeparator;
                        break;
                    default:
                        break;
                }

                try
                {
                    oDateTime = DateTime.Parse(strDate);
                    SetDate(oDateTime);
                }
                catch (FormatException)
                {
                    SetDate("");
                    manualTextBox.Focus();
                }
            }
            else if (manualTextBox.Text.Length >= 3)
            {
                // 
                if (manualTextBox.Text.Substring(0, 2).ToLower() == m_today + "-" ||
                         manualTextBox.Text.Substring(0, 2).ToLower() == m_today + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(manualTextBox.Text.Substring(2, manualTextBox.Text.Length - 2));

                        if (manualTextBox.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddDays(-anz) : ((DateTime)dateTimePicker1.Value).AddDays(-anz);

                        }
                        else
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddDays(anz) : ((DateTime)dateTimePicker1.Value).AddDays(anz);
                        }

                        //dateTimePicker1.Value = new NextWorkdayCalculator().GetNextWorkDay((DateTime)dateTimePicker1.Value, 1);

                        SetDate(dateTimePicker1.Value != null ? ((DateTime)dateTimePicker1.Value).ToString() : "");
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Substring(0, 2).ToLower() == m_month + "-" ||
                         manualTextBox.Text.Substring(0, 2).ToLower() == m_month + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(manualTextBox.Text.Substring(2, manualTextBox.Text.Length - 2));

                        if (manualTextBox.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddMonths(-anz) : ((DateTime)dateTimePicker1.Value).AddMonths(-anz);
                        }
                        else
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddMonths(anz) : ((DateTime)dateTimePicker1.Value).AddMonths(anz);
                        }

                        //dateTimePicker1.Value = new NextWorkdayCalculator().GetNextWorkDay((DateTime)dateTimePicker1.Value, 1);

                        SetDate(dateTimePicker1.Value != null ? ((DateTime)dateTimePicker1.Value).ToString() : "");
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Substring(0, 2).ToLower() == m_week + "-" ||
                         manualTextBox.Text.Substring(0, 2).ToLower() == m_week + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(manualTextBox.Text.Substring(2, manualTextBox.Text.Length - 2));

                        if (manualTextBox.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddDays(-(7 * anz)) : ((DateTime)dateTimePicker1.Value).AddDays(-(7 * anz));
                        }
                        else
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddDays((7 * anz)) : ((DateTime)dateTimePicker1.Value).AddDays((7 * anz));
                        }

                        //dateTimePicker1.Value = new NextWorkdayCalculator().GetNextWorkDay((DateTime)dateTimePicker1.Value, 1);

                        SetDate(dateTimePicker1.Value != null ? ((DateTime)dateTimePicker1.Value).ToString() : "");
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Substring(0, 2).ToLower() == m_year + "-" ||
                         manualTextBox.Text.Substring(0, 2).ToLower() == m_year + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(manualTextBox.Text.Substring(2, manualTextBox.Text.Length - 2));

                        if (manualTextBox.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddYears(-anz) : ((DateTime)dateTimePicker1.Value).AddYears(-anz);
                        }
                        else
                        {
                            dateTimePicker1.Value = dateTimePicker1.Value == null ? DateTime.Today.AddYears(anz) : ((DateTime)dateTimePicker1.Value).AddYears(anz);
                        }

                        //dateTimePicker1.Value = new NextWorkdayCalculator().GetNextWorkDay((DateTime)dateTimePicker1.Value, 1);

                        SetDate(dateTimePicker1.Value != null ? ((DateTime)dateTimePicker1.Value).ToString() : "");
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Length == 6)    // z.B. ddmmyy oder yymmdd 
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate =
                                manualTextBox.Text.Substring(0, 2) + m_DateSeparator +
                                manualTextBox.Text.Substring(2, 2) + m_DateSeparator +
                                DateTime.Today.Year.ToString().Substring(0, 2) +
                                   manualTextBox.Text.Substring(4, 2);
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate =
                                DateTime.Today.Year.ToString().Substring(0, 2) +
                                manualTextBox.Text.Substring(0, 2) + m_DateSeparator + " " +
                                manualTextBox.Text.Substring(2, 2) + m_DateSeparator + " " +
                                manualTextBox.Text.Substring(4, 2) + m_DateSeparator;
                            break;
                        default:
                            break;
                    }

                    try
                    {
                        oDateTime = DateTime.Parse(strDate);
                        SetDate(oDateTime);
                    }
                    catch (FormatException)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Length == 8)    // z.B. ddmmyyyy oder yyyymmdd 
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate = manualTextBox.Text.Substring(0, 2) + m_DateSeparator +
                                      manualTextBox.Text.Substring(2, 2) + m_DateSeparator +
                                      manualTextBox.Text.Substring(4, 4);
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate = manualTextBox.Text.Substring(0, 4) + m_DateSeparator + " " +
                                      manualTextBox.Text.Substring(4, 2) + m_DateSeparator + " " +
                                      manualTextBox.Text.Substring(6, 2) + m_DateSeparator;
                            break;
                        default:
                            break;
                    }

                    try
                    {
                        oDateTime = DateTime.Parse(strDate);
                        SetDate(oDateTime);
                    }
                    catch (FormatException)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
                else if (manualTextBox.Text.Length == 4)
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate = manualTextBox.Text.Substring(0, 2) + m_DateSeparator +
                                            manualTextBox.Text.Substring(2, 2) + m_DateSeparator +
                                            DateTime.Today.Year;
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate = DateTime.Today.Year + m_DateSeparator + " " +
                                            manualTextBox.Text.Substring(2, 2) + m_DateSeparator + " " +
                                            manualTextBox.Text.Substring(0, 2) + m_DateSeparator;
                            break;
                        default:
                            break;
                    }
                    try
                    {
                        oDateTime = DateTime.Parse(strDate);
                        SetDate(oDateTime);
                    }
                    catch (FormatException)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }

                else
                {
                    try
                    {
                        oDateTime = DateTime.Parse(manualTextBox.Text);
                        SetDate(oDateTime);
                    }
                    catch (FormatException)
                    {
                        SetDate("");
                        manualTextBox.Focus();
                    }
                }
            }
            else
            {
                if (manualTextBox.Text.Contains(m_today))
                {

                }
                else
                {
                    SetDate("");
                    manualTextBox.Focus();
                }
            }
        }

        private void SetDate(string pString)
        {
            m_value = pString;
            DateTime tmp_datetime = DateTime.Today;

            if (DateTime.TryParse(m_value, out tmp_datetime))
            {
                resultLabel.Text = tmp_datetime.ToShortDateString() +
                    " (" + m_CultureInfo.DateTimeFormat.GetDayName(tmp_datetime.DayOfWeek) +
                    ", " + GetCalendarWeek(tmp_datetime).Week + ")";
            }
            else
            {
                resultLabel.Text = string.Empty;
            }
        }

        private void SetDate(DateTime pDateTime)
        {
            //if (dateTimePicker1.Value != null)
            //{
                if (pDateTime >= dateTimePicker1.MinDate && pDateTime <= dateTimePicker1.MaxDate)
                {
                    dateTimePicker1.Value = pDateTime;
                    m_value = Convert.ToDateTime(dateTimePicker1.Value).ToShortDateString();
                    resultLabel.Text = m_value +
                        " (" + m_CultureInfo.DateTimeFormat.GetDayName(Convert.ToDateTime(dateTimePicker1.Value).DayOfWeek) +
                        ", " + GetCalendarWeek(Convert.ToDateTime(dateTimePicker1.Value)).Week + ")";
                }
                else
                {
                    resultLabel.Text = string.Empty;
                }
            //}
            //else
            //{
            //    resultLabel.Text = string.Empty;
            //}
        }

        public class CalendarWeek
        {
            /// <summary>
            /// Das Jahr
            /// </summary>
            public int Year;

            /// <summary>
            /// Die Kalenderwoche
            /// </summary>
            public int Week;

            /// <summary>
            /// Konstruktor
            /// </summary>
            /// <param name="year">Das Jahr</param>
            /// <param name="week">Die Kalenderwoche</param>
            public CalendarWeek(int year, int week)
            {
                Year = year;
                Week = week;
            }
        }

        public CalendarWeek GetCalendarWeek(DateTime date)
        {
            // Aktuelle Kultur ermitteln
            CultureInfo currentCulture = CultureInfo.CurrentCulture;

            // Aktuellen Kalender ermitteln
            Calendar calendar = currentCulture.Calendar;

            // Kalenderwoche über das Calendar-Objekt ermitteln
            int calendarWeek = calendar.GetWeekOfYear(date,
               currentCulture.DateTimeFormat.CalendarWeekRule,
               currentCulture.DateTimeFormat.FirstDayOfWeek);

            // Überprüfen, ob eine Kalenderwoche größer als 52
            // ermittelt wurde und ob die Kalenderwoche des Datums
            // in einer Woche 2 ergibt: In diesem Fall hat
            // GetWeekOfYear die Kalenderwoche nicht nach ISO 8601 
            // berechnet (Montag, der 31.12.2007 wird z. B.
            // fälschlicherweise als KW 53 berechnet). 
            // Die Kalenderwoche wird dann auf 1 gesetzt
            if (calendarWeek > 52)
            {
                date = date.AddDays(7);
                int testCalendarWeek = calendar.GetWeekOfYear(date,
                   currentCulture.DateTimeFormat.CalendarWeekRule,
                   currentCulture.DateTimeFormat.FirstDayOfWeek);
                if (testCalendarWeek == 2)
                    calendarWeek = 1;
            }

            // Das Jahr der Kalenderwoche ermitteln
            int year = date.Year;
            if (calendarWeek == 1 && date.Month == 12)
                year++;
            if (calendarWeek >= 52 && date.Month == 1)
                year--;

            // Die ermittelte Kalenderwoche zurückgeben
            return new CalendarWeek(year, calendarWeek);
        }

        #endregion


        private void manualTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
            else if (e.KeyCode == Keys.Enter)
                DialogResult = DialogResult.OK;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //IsHoliday();

            IsWeekend();
        }

        private bool IsWeekend()
        {
            if (dateTimePicker1.Value != null)
            {
                DateTime _date = Convert.ToDateTime(dateTimePicker1.Value);


                // Weekends
                CultureInfo oCultureInfo = CultureInfo.CurrentCulture;
                //oCultureInfo.  .DateTimeFormat  ee = new DateTimeFormatInfo (); //.fiFirstDayOfWeek     // _date.DayOfWeek 

                return true;
            }
            else
            {
                return false;
            }
        }

        //private bool IsHoliday()
        //{
        //    if (dateTimePicker1.Value != null)
        //    {
        //        DateTime _date = Convert.ToDateTime(dateTimePicker1.Value);

        //        foreach (structHolidays oStructHolidays in g_holidays)
        //        {
        //            if (oStructHolidays.date == _date)
        //            {
        //                stateTextBox.Text = oStructHolidays.description;
        //                return true;
        //            }
        //        }
        //        return false;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}

        private void SetFormat()
        {
            CultureInfo oCultureInfo = CultureInfo.CurrentCulture;
            m_DateSeparator = oCultureInfo.DateTimeFormat.DateSeparator;
            m_ShortDatePattern = oCultureInfo.DateTimeFormat.ShortDatePattern;

            switch (m_ShortDatePattern)
            {
                case "dd.MM.yyyy":        // de-DE                       
                    m_today = "h";
                    m_month = "m";
                    m_year = "j";
                    break;
                case "dd/MM/yyyy":        // en-GB 
                    m_today = "t";
                    m_month = "m";
                    m_year = "y";
                    break;
                case "yyyy.MM.dd.":     // hu-HU 
                case "yyyy. MM. dd.":     // hu-HU 
                    m_today = "m";
                    m_month = "h";
                    m_year = "é";
                    break;
            }
        }

        private void stateTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(stateTextBox.Text))
                System.Media.SystemSounds.Beep.Play();
        }

        private void dateTimePicker1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                DialogResult = DialogResult.Cancel;
            else if (e.KeyCode == Keys.Enter)
                DialogResult = DialogResult.OK;
        }

        private void ohaDateTimePicker_Input_Deactivate(object sender, EventArgs e)
        {
            Close();
        }

        private void ohaDateTimePicker_Input_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Properties.Settings.Default.Save();
        }

        private void forward_backCheckBox_CheckStateChanged(object sender, EventArgs e)
        {
            //if (m_date != null)
            //    dateTimePicker1.Value = m_date;

            manualTextBox_TextChanged(null, null);
        }






    }
}

