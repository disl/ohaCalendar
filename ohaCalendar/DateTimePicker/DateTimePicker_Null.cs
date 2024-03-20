using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Globalization;
using System.Runtime.InteropServices;


namespace ohaERP_Library
{
    //[System.ComponentModel.DefaultBindingProperty("Value")]
    [ComVisible(false)]
    public partial class DateTimePicker_Null : UserControl
    {
        private string m_value = "";
        private string m_DateSeparator;
        private string m_ShortDatePattern;
        private string m_today = "t";
        private string m_month = "m";
        private string m_year = "y";

        public DateTimePicker_Null()
        {
            InitializeComponent();
            m_value = "";
            SetFormat();



        }


#region Properties

        [Bindable(true), Browsable(true)]
        public Object Value
        {
            get
            {
                if (m_value == "")
                    return null;
                else
                    return dateTimePicker1.Value;
            }
            set
            {
                if (value == null || value == DBNull.Value)
                    SetDate("");
                else
                    SetDate((DateTime)value);
            }
        }

        [Browsable(true)]
        public string CustomFormat
        {
            get
            {
                return dateTimePicker1.CustomFormat;
            }
            set
            {
                dateTimePicker1.CustomFormat = value;
            }
        }

        [Browsable(true)]
        public DateTimePickerFormat Format
        {
            get
            {
                return dateTimePicker1.Format;
            }
            set
            {
                dateTimePicker1.Format = value;
            }
        }
        [EditorBrowsable(EditorBrowsableState.Advanced )]
        [Browsable(true)]
        public override Color BackColor
        {
            get
            {
                return textBox1.BackColor;                
            }
            set
            {
                textBox1.BackColor = value;
            }
        }

        //[Browsable(true)]
        //public override string Text
        //{
        //    get
        //    {
        //        return textBox1.Text;
        //    }
        //    set
        //    {
        //        textBox1.Text = value;
        //    }
        //}
      
#endregion

#region Functions

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

        private void DateTimePicker_Null_Resize(Object sender, EventArgs e)
        {
            //textBox1.Location = new Point(0, 0);
            //textBox1.Width = this.Width - dateTimePicker1.Width - 30;
            //dateTimePicker1.Location = new Point(textBox1.Width, dateTimePicker1.Location.Y);             
        }

        private void DateTimePicker_Null_SizeChanged(Object sender, EventArgs e)
        {
            //textBox1.Location = new Point(0, 0);
            //textBox1.Width = this.Width - dateTimePicker1.Width - 30;  //this.Width - dateTimePicker1.Width - 20;
            //dateTimePicker1.Location = new Point(textBox1.Width, dateTimePicker1.Location.Y); 
        }

        private void textBox1_KeyDown(Object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                SetDate("");
            }
        }

        private void textBox1_KeyPress(Object sender, KeyPressEventArgs e)
        {
            e.Handled = false;

            if (!(e.KeyChar.ToString() == "0" || e.KeyChar.ToString() == "1" || 
                e.KeyChar.ToString() == "2" || e.KeyChar.ToString() == "3" || 
                e.KeyChar.ToString() == "4" || e.KeyChar.ToString() == "5" ||
                e.KeyChar.ToString() == "6" || e.KeyChar.ToString() == "7" || 
                e.KeyChar.ToString() == "8" || e.KeyChar.ToString() == "9" || 
                e.KeyChar.ToString() == "." || e.KeyChar.ToString() == "\b"||
                e.KeyChar.ToString().ToLower () == m_today || e.KeyChar.ToString() == "-" || 
                e.KeyChar.ToString() == "+" || e.KeyChar.ToString() == "0" ||
                e.KeyChar.ToString().ToLower() == m_month || e.KeyChar.ToString().ToLower() == m_year ||
                e.KeyChar.ToString() == "/"
               ))

                e.Handled = true;
            
        }

        private void textBox1_Leave(Object sender, EventArgs e)
        {
            DateTime oDateTime;
            string strDate = "";
            int anz = 0;


            if (textBox1.Text == "")
            {
                SetDate("");
            }
            else if (textBox1.Text.ToLower() == m_today)
            {
                SetDate(DateTime.Today.ToShortDateString());
            }
            else if (textBox1.Text.Length > 2)
            {

                if (textBox1.Text.Substring(0, 2).ToLower() == m_today + "-" ||
                         textBox1.Text.Substring(0, 2).ToLower() == m_today + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(textBox1.Text.Substring(2, textBox1.Text.Length - 2));

                        if (textBox1.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = DateTime.Today.AddDays(-anz);
                        }
                        else
                        {
                            dateTimePicker1.Value = DateTime.Today.AddDays(anz);
                        }
                        SetDate(dateTimePicker1.Text);
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        textBox1.Focus();
                    }
                }
                else if (textBox1.Text.Substring(0, 2).ToLower() == m_month + "-" ||
                         textBox1.Text.Substring(0, 2).ToLower() == m_month + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(textBox1.Text.Substring(2, textBox1.Text.Length - 2));

                        if (textBox1.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = DateTime.Today.AddMonths(-anz);
                        }
                        else
                        {
                            dateTimePicker1.Value = DateTime.Today.AddMonths(anz);
                        }
                        SetDate(dateTimePicker1.Text);
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        textBox1.Focus();
                    }
                }
                else if (textBox1.Text.Substring(0, 2).ToLower() == m_year + "-" ||
                         textBox1.Text.Substring(0, 2).ToLower() == m_year + "+")
                {
                    try
                    {
                        anz = Convert.ToInt32(textBox1.Text.Substring(2, textBox1.Text.Length - 2));

                        if (textBox1.Text.Substring(1, 1) == "-")
                        {
                            dateTimePicker1.Value = DateTime.Today.AddYears(-anz);
                        }
                        else
                        {
                            dateTimePicker1.Value = DateTime.Today.AddYears(anz);
                        }
                        SetDate(dateTimePicker1.Text);
                    }
                    catch (Exception)
                    {
                        SetDate("");
                        textBox1.Focus();
                    }
                }
                else if (textBox1.Text.Length == 6)    // z.B. ddmmyy oder yymmdd 
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate =
                                textBox1.Text.Substring(0, 2) + m_DateSeparator +
                                textBox1.Text.Substring(2, 2) + m_DateSeparator +
                                DateTime.Today.Year.ToString().Substring(0, 2) +
                                   textBox1.Text.Substring(4, 2);
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate =
                                DateTime.Today.Year.ToString().Substring(0, 2) +
                                textBox1.Text.Substring(0, 2) + m_DateSeparator + " " +
                                textBox1.Text.Substring(2, 2) + m_DateSeparator + " " +
                                textBox1.Text.Substring(4, 2) + m_DateSeparator;
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
                        textBox1.Focus();
                    }
                }
                else if (textBox1.Text.Length == 8)    // z.B. ddmmyyyy oder yyyymmdd 
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate = textBox1.Text.Substring(0, 2) + m_DateSeparator +
                                      textBox1.Text.Substring(2, 2) + m_DateSeparator +
                                      textBox1.Text.Substring(4, 4);
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate = textBox1.Text.Substring(0, 4) + m_DateSeparator + " " +
                                      textBox1.Text.Substring(4, 2) + m_DateSeparator + " " +
                                      textBox1.Text.Substring(6, 2) + m_DateSeparator;
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
                        textBox1.Focus();
                    }
                }
                else if (textBox1.Text.Length == 4)
                {
                    switch (m_ShortDatePattern)
                    {
                        case "dd.MM.yyyy":        // de-DE                       
                        case "dd/MM/yyyy":        // en-GB 
                            strDate = textBox1.Text.Substring(0, 2) + m_DateSeparator +
                                            textBox1.Text.Substring(2, 2) + m_DateSeparator +
                                            DateTime.Today.Year;
                            break;
                        case "yyyy. MM. dd.":     // hu-HU       
                            strDate = DateTime.Today.Year + m_DateSeparator + " " +
                                            textBox1.Text.Substring(2, 2) + m_DateSeparator + " " +
                                            textBox1.Text.Substring(0, 2) + m_DateSeparator;
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
                        textBox1.Focus();
                    }
                }

                else
                {
                    try
                    {
                        oDateTime = DateTime.Parse(textBox1.Text);
                        SetDate(oDateTime);
                    }
                    catch (FormatException)
                    {
                        SetDate("");
                        textBox1.Focus();
                    }
                }
            }
            else
            {
                SetDate("");
                textBox1.Focus();
            }
        }

        private void SetDate(DateTime pDateTime)
        {
            dateTimePicker1.Value = pDateTime;
            m_value = dateTimePicker1.Text;
            textBox1.Text = m_value;
        }

        private void SetDate(string pString)
        {
            m_value = pString;
            textBox1.Text = m_value;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            SetDate(dateTimePicker1.Text);
        }

#endregion

        

        



      

  



   



    }
}
