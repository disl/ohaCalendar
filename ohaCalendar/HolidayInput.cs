using System.Data;

namespace ohaCalendar
{
    public partial class HolidayInput : Form
    {
        public string State { get; set; }
        public int Year { get; set; }

        public HolidayInput()
        {
            InitializeComponent();
        }

        private void HolidayInput_Load(object sender, EventArgs e)
        {
            yearNumericUpDown.Value = DateTime.Now.Year;

            FillStateCMB();
        }

        private void FillStateCMB()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("DisplayMember", typeof(string));
            dt.Columns.Add("ValueMember", typeof(string));
            AddRow(dt, "Baden-Württemberg", "BW");
            AddRow(dt, "Bayern", "BY");
            AddRow(dt, "Berlin", "BW");
            AddRow(dt, "Brandenburg", "BB");
            AddRow(dt, "Bremen", "HB");
            AddRow(dt, "Hamburg", "HH");
            AddRow(dt, "Hessen", "HE");
            AddRow(dt, "Mecklenburg-Vorpommern", "MV");
            AddRow(dt, "Niedersachsen", "NI");
            AddRow(dt, "Nordrhein-Westfalen", "NW");
            AddRow(dt, "Rheinland-Pfalz", "RP");
            AddRow(dt, "Saarland", "SL");
            AddRow(dt, "Sachsen", "SN");
            AddRow(dt, "Sachsen-Anhalt", "ST");
            AddRow(dt, "Schleswig-Holstein", "SH");
            AddRow(dt, "Thüringen", "TH");

            stateComboBox.DataSource = dt;
            stateComboBox.DisplayMember = "DisplayMember";
            stateComboBox.ValueMember = "ValueMember";
        }

        private static void AddRow(DataTable dt, string displayMember, string valueMember)
        {
            var new_row = dt.NewRow();
            new_row["DisplayMember"] = displayMember;
            new_row["ValueMember"] = valueMember;
            dt.Rows.Add(new_row);
        }

        private void stateComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            State = stateComboBox.SelectedValue.ToString();
        }

        private void yearNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            Year = (int)yearNumericUpDown.Value;
        }
    }
}
