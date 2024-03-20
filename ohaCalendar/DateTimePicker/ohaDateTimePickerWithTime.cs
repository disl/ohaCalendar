using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace ohaERP_Library.DateTimePicker
{
    public partial class ohaDateTimePickerWithTime : UserControl
    {
        public event EventHandler ValueChanged;

        Font m_old_font;

        private Label m_Label;
        [Browsable(true)]
        [Category("Users properties")]
        public Label Label
        {
            get { return m_Label; }
            set { m_Label = value; }
        }

        [Browsable(false)]
        [Bindable(true)]
        [DesignerSerializationVisibility(0)]
        [RefreshProperties(RefreshProperties.All)]
        public object Value
        {


            get
            {
                if (DateDateTimePicker.Value == null || TimeDateTimePicker.Value == null)
                    return null;
                else
                    return new DateTime(
                        ((DateTime)DateDateTimePicker.Value).Year,
                        ((DateTime)DateDateTimePicker.Value).Month,
                        ((DateTime)DateDateTimePicker.Value).Day,
                        ((DateTime)TimeDateTimePicker.Value).Hour,
                        ((DateTime)TimeDateTimePicker.Value).Minute,
                        0);
            }
            set
            {
                if (value != null && value != DBNull.Value)
                {
                    DateDateTimePicker.Value = (DateTime)value;
                    TimeDateTimePicker.Value = (DateTime)value;
                }
                else
                {
                    DateDateTimePicker.Value = null;
                    TimeDateTimePicker.Value = null;
                }

                ////DateTime _tmp_date_time = DBNull.Value ;  

                ////if (Value != null &&
                ////   Convert.ToDateTime  (DateTime)value > DateTime.MinValue && (DateTime)value < DateTime.MaxValue)
                ////    DateDateTimePicker.Value = value;
                ////else
                //    DateDateTimePicker.Value = value;
            }
        }

        public DateTime? MinDate
        {
            get
            {
                if (DateDateTimePicker.Value == null || TimeDateTimePicker.Value == null)
                    return null;
                else
                {
                    return new DateTime(
                        DateDateTimePicker.MinDate.Year,
                        DateDateTimePicker.MinDate.Month,
                        DateDateTimePicker.MinDate.Day,
                        ((DateTime)TimeDateTimePicker.Value).Hour,
                        ((DateTime)TimeDateTimePicker.Value).Minute,
                        0);
                }
            }
            set
            {
                if (value != null)
                {
                    DateDateTimePicker.MinDate = (DateTime)value;
                    TimeDateTimePicker.MinDate = (DateTime)value;
                    DateDateTimePicker.Value = (DateTime)value;
                    TimeDateTimePicker.Value = (DateTime)value;
                }
            }
        }

        public ohaDateTimePickerWithTime()
        {
            InitializeComponent();

            TimeDateTimePicker.NormalMode = true;
        }

        public void SetValue(DateTime? datetime)
        {
            DateDateTimePicker.Value = datetime;
            TimeDateTimePicker.Value = datetime;
        }

        private void DateDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (ValueChanged != null)
            {
                ValueChanged(this, EventArgs.Empty);
            }
            TimeDateTimePicker.Value = DateDateTimePicker.Value;
        }




        private void TimeDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            // 
        }

        private void TimeDateTimePicker_Validated(object sender, EventArgs e)
        {
            ChangeTime();
        }

        DateTime? m_old_time = null;

        private void ChangeTime()
        {
            DateTime? oDateTime = null;

            if (DateDateTimePicker.Value != null &&
                TimeDateTimePicker.Value != null)
            {
                oDateTime = new DateTime(
                    Convert.ToDateTime(DateDateTimePicker.Value).Year,
                    Convert.ToDateTime(DateDateTimePicker.Value).Month,
                    Convert.ToDateTime(DateDateTimePicker.Value).Day,
                    Convert.ToDateTime(TimeDateTimePicker.Value).Hour,
                    Convert.ToDateTime(TimeDateTimePicker.Value).Minute, 0);

                //if (
                //               Convert.ToDateTime (m_old_time). != oDateTime)
                //           {
                //               DateDateTimePicker.Value = oDateTime;
                //               m_old_time = oDateTime;
                //           }
            }
            DateDateTimePicker.Value = oDateTime;

        }

        //private void ohaDateTimePickerWithTime_BindingContextChanged(object sender, EventArgs e)
        //{
        //    SetBindingContextChanged();
        //}

        //protected void SetBindingContextChanged()
        //{
        //    Font UnderlineTrue, UnderlineFalse;
        //    FontFamily oFontFamily = new FontFamily("Arial");
        //    UnderlineTrue = new Font(oFontFamily, 9, FontStyle.Underline);
        //    UnderlineFalse = new Font(oFontFamily, 9);

        //    if (Label != null && DataBindings.Count > 0)
        //    {
        //        Label.Font = DB.dbUtilities.is_field_nullable(this) ? UnderlineFalse : UnderlineTrue;
        //    }
        //}

        private void ohaDateTimePickerWithTime_Enter(object sender, EventArgs e)
        {
            //Font = new Font(Font, FontStyle.Bold);

            if (Label != null)
            {
                m_old_font = Label.Font;
                Label.Font = new Font(Label.Font, m_old_font.Style | FontStyle.Bold);
            }
        }

        private void ohaDateTimePickerWithTime_Leave(object sender, EventArgs e)
        {
            //Font = new Font(Font, FontStyle.Regular);

            if (Label != null)
            {
                Label.Font = m_old_font;
            }
        }






    }
}
