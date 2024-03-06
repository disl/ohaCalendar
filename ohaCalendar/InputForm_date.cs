namespace ohaCalendar
{
    public partial class InputForm_date : Form
    {
        private DateTime? m_min_date;
        //public DateTime? Min_date { set { m_min_date = value; } }

        private string m_labeltext;
        public string Labeltext { set { m_labeltext = value; } }

        private DateTime m_value;
        public DateTime Value { get { return m_value; } set { m_value = value; } }      

        
        public InputForm_date()
        {
            InitializeComponent();
        }

        public InputForm_date(DateTime? Min_date)
        {
            InitializeComponent();

            m_min_date = Min_date;
        }    

        private void InputForm_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(m_labeltext))
                label1.Text = m_labeltext;

            if (m_min_date != null)
                ohaDateTimePicker.MinDate = Convert.ToDateTime (m_min_date);
            
            if (m_value != null)
                ohaDateTimePicker.Value = m_value;
            else
                ohaDateTimePicker.Value = new DateTime();

            //ohaDateTimePicker.fo;
        }

        private void InputForm_FormClosing(object sender, FormClosingEventArgs e)
        {
          
        }

        private void ohaDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            m_value = ohaDateTimePicker.Value != null ? Convert.ToDateTime(ohaDateTimePicker.Value) : new DateTime();
        }
    }



}

