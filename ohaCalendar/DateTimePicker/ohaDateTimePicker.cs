using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ohaERP_Library
{
    /// <summary>
    /// Represents a Windows date time picker control. It enhances the .NET standard <b>DateTimePicker</b>
    /// control with a ReadOnly mode as well as with the possibility to show empty values (null values).
    /// </summary>
    [ComVisible(true)]
    public partial class ohaDateTimePicker : System.Windows.Forms.DateTimePicker
    {
        #region Member variables                

        Font m_old_font;

        // true, when no date shall be displayed (empty DateTimePicker)
        private bool _isNull;

        // If _isNull = true, this value is shown in the DTP
        private string _nullValue;

        // The format of the DateTimePicker control
        private DateTimePickerFormat _format = DateTimePickerFormat.Short;

        // The custom format of the DateTimePicker control
        private string _customFormat;

        // The format of the DateTimePicker control as string
        private string _formatAsString;

        private ContextMenuStrip contextMenuStrip1;
        private IContainer components;
        private ToolStripMenuItem toolStripMenuItem_today;
        private ToolStripMenuItem toolStripMenuItem_normal_mode;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem toolStripMenuItem_table;
        private ToolTip toolTip1;


        //public DateTime MaxDate
        //{
        //    get { return this.; }
        //}

        // public DateTime MinDate

        [Browsable(true)]
        [DefaultValue(true)]
        [Category("Users properties")]
        public bool NormalMode { get; set; }


        [Browsable(true)]
        [DefaultValue(true)]
        [Category("Users properties")]
        public bool AllowEdit
        {
            get { return m_allow_edit; }
            set { m_allow_edit = value; }
        }
        private bool m_allow_edit = true;




        [Browsable(true)]
        [Category("Users properties")]
        public void SetAllowEdit(bool AllowEdit)
        {
            m_allow_edit = AllowEdit;

            if (AllowEdit)
            {
                Enabled = true;
                base.BackColor = SystemColors.Window;

                //this.ShowUpDown = false;
            }
            else
            {
                Enabled = false;
                BackColor = SystemColors.Control;

                //this.ShowUpDown = true;

            }
            base.ForeColor = SystemColors.WindowText;
        }

        [Browsable(false)]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public override Color ForeColor
        {
            set
            {
                base.ForeColor = SystemColors.WindowText;
            }
            get
            {
                return base.ForeColor;
            }
        }


        ErrorProvider m_ErrorProvider = new ErrorProvider();

        private Label m_Label;
        [Browsable(true)]
        [Category("Users properties")]
        public Label Label
        {
            get { return m_Label; }
            set { m_Label = value; }
        }

        private bool m_interner_validating = true;
        private bool m_drop_down_on;

        [Browsable(true)]
        [Category("Users properties")]
        [DefaultValue(true)]
        public bool InternerValidating
        {
            get { return m_interner_validating; }
            set { m_interner_validating = value; }
        }

        #endregion

        #region Constructor
        /// <summary>
        /// Default Constructor
        /// </summary>
        public ohaDateTimePicker()
            : base()
        {
            InitializeComponent();

            base.Format = DateTimePickerFormat.Custom;
            NullValue = " ";
            Format = DateTimePickerFormat.Short;
        }

        protected void SetBindingContextChanged()
        {
            //Font UnderlineTrue, UnderlineFalse;
            //FontFamily oFontFamily = new FontFamily("Arial");
            //UnderlineTrue = new Font(oFontFamily, 9, FontStyle.Underline);
            //UnderlineFalse = new Font(oFontFamily, 9);

            //if (Label != null && DataBindings.Count > 0)
            //{
            //    Label.Font = DB.dbUtilities.is_field_nullable(this) ? UnderlineFalse : UnderlineTrue;
            //}
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Gets or sets the date/time value assigned to the control.
        /// </summary>
        /// <value>The DateTime value assigned to the control
        /// </value>
        /// <remarks>
        /// <p>If the <b>Value</b> property has not been changed in code or by the user, it is set
        /// to the current date and time (<see cref="DateTime.Now"/>).</p>
        /// <p>If <b>Value</b> is <b>null</b>, the DateTimePicker shows 
        /// <see cref="NullValue"/>.</p>
        /// </remarks>
        public new Object? Value
        {
            get
            {
                if (_isNull)
                    return null;
                else
                {
                    return base.Value;
                }
            }
            set
            {
                if (value == null || value == DBNull.Value)
                {
                    SetToNullValue();
                    OnValueChanged(EventArgs.Empty);
                }
                else
                {
                    if (value.ToString() != "")
                    {
                        SetToDateTimeValue();

                        DateTime tmp_date = DateTime.Parse(value.ToString());

                        if (tmp_date >= MinDate && tmp_date <= MaxDate)
                            base.Value = DateTime.Parse(value.ToString());
                        else
                        {
                            if (tmp_date < MinDate)
                                base.Value = MinDate;
                            else
                                base.Value = MaxDate;
                        }

                    }
                    else
                    {
                        SetToNullValue();
                        OnValueChanged(EventArgs.Empty);
                    }
                }
            }
        }

        public string ToString(string CustomFormat)
        {
            this.CustomFormat = CustomFormat;

            return base.Value.ToString();
        }


        /// <summary>
        /// Gets or sets the format of the date and time displayed in the control.
        /// </summary>
        /// <value>One of the <see cref="DateTimePickerFormat"/> values. The default is 
        /// <see cref="DateTimePickerFormat.Short"/>.</value>
        [Browsable(true)]
        [DefaultValue(DateTimePickerFormat.Short), TypeConverter(typeof(Enum))]
        public new DateTimePickerFormat Format
        {
            get { return _format; }
            set
            {
                _format = value;
                SetFormat();
                OnFormatChanged(EventArgs.Empty);
            }
        }

        /// <summary>
        /// Gets or sets the custom date/time format string.
        /// <value>A string that represents the custom date/time format. The default is a null
        /// reference (<b>Nothing</b> in Visual Basic).</value>
        /// </summary>
        public new String CustomFormat
        {
            get { return _customFormat; }
            set
            {
                _customFormat = value;
            }
        }

        /// <summary>
        /// Gets or sets the string value that is assigned to the control as null value. 
        /// </summary>
        /// <value>The string value assigned to the control as null value.</value>
        /// <remarks>
        /// If the <see cref="Value"/> is <b>null</b>, <b>NullValue</b> is
        /// shown in the <b>DateTimePicker</b> control.
        /// </remarks>
        [Browsable(true)]
        [Category("Behavior")]
        [Description("The string used to display null values in the control")]
        [DefaultValue(" ")]
        public String NullValue
        {
            get { return _nullValue; }
            set { _nullValue = value; }
        }
        #endregion

        #region Private methods/properties
        /// <summary>
        /// Stores the current format of the DateTimePicker as string. 
        /// </summary>
        private string FormatAsString
        {
            get { return _formatAsString; }
            set
            {
                _formatAsString = value;
                base.CustomFormat = value;
            }
        }

        /// <summary>
        /// Sets the format according to the current DateTimePickerFormat.
        /// </summary>
        private void SetFormat()
        {
            CultureInfo ci = Thread.CurrentThread.CurrentCulture;
            DateTimeFormatInfo dtf = ci.DateTimeFormat;
            switch (_format)
            {
                case DateTimePickerFormat.Long:
                    FormatAsString = dtf.LongDatePattern;
                    break;
                case DateTimePickerFormat.Short:
                    FormatAsString = dtf.ShortDatePattern;
                    break;
                case DateTimePickerFormat.Time:
                    FormatAsString = dtf.ShortTimePattern;
                    break;
                case DateTimePickerFormat.Custom:
                    FormatAsString = CustomFormat;
                    break;
            }
        }

        /// <summary>
        /// Sets the <b>DateTimePicker</b> to the value of the <see cref="NullValue"/> property.
        /// </summary>
        private void SetToNullValue()
        {
            _isNull = true;
            base.CustomFormat = (_nullValue == null || _nullValue == String.Empty) ? " " : "'" + _nullValue + "'";

        }

        /// <summary>
        /// Sets the <b>DateTimePicker</b> back to a non null value.
        /// </summary>
        private void SetToDateTimeValue()
        {
            if (_isNull)
            {
                SetFormat();
                _isNull = false;
                base.OnValueChanged(new EventArgs());
            }
        }
        #endregion

        #region OnXXXX()

        /// <summary>
        /// This member overrides <see cref="DateTimePicker.OnCloseUp"/>.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCloseUp(EventArgs e)
        {
            if (MouseButtons == MouseButtons.None)
            {
                if (_isNull)
                {
                    SetToDateTimeValue();
                    _isNull = false;
                }
            }
            base.OnCloseUp(e);
        }

        /// <summary>
        /// This member overrides <see cref="Control.OnKeyDown"/>.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnKeyUp(KeyEventArgs e)
        {
            //DateTime oDateTime = DateTime.Today;

            if (e.KeyCode == Keys.Delete)
            {
                Value = null;
                OnValueChanged(EventArgs.Empty);
            }
            base.OnKeyUp(e);
        }



        protected override void OnValueChanged(EventArgs eventargs)
        {
            base.OnValueChanged(eventargs);
        }

        #endregion

        private void InitializeComponent()
        {
            components = new Container();
            ComponentResourceManager resources = new ComponentResourceManager(typeof(ohaDateTimePicker));
            contextMenuStrip1 = new ContextMenuStrip(components);
            toolStripMenuItem_today = new ToolStripMenuItem();
            toolStripMenuItem_normal_mode = new ToolStripMenuItem();
            toolStripSeparator1 = new ToolStripSeparator();
            toolStripMenuItem_table = new ToolStripMenuItem();
            toolTip1 = new ToolTip(components);
            contextMenuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            resources.ApplyResources(contextMenuStrip1, "contextMenuStrip1");
            contextMenuStrip1.Items.AddRange(new ToolStripItem[] { toolStripMenuItem_today, toolStripMenuItem_normal_mode, toolStripSeparator1, toolStripMenuItem_table });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.ItemClicked += contextMenuStrip1_ItemClicked;
            // 
            // toolStripMenuItem_today
            // 
            toolStripMenuItem_today.Name = "toolStripMenuItem_today";
            resources.ApplyResources(toolStripMenuItem_today, "toolStripMenuItem_today");
            // 
            // toolStripMenuItem_normal_mode
            // 
            toolStripMenuItem_normal_mode.Name = "toolStripMenuItem_normal_mode";
            resources.ApplyResources(toolStripMenuItem_normal_mode, "toolStripMenuItem_normal_mode");
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(toolStripSeparator1, "toolStripSeparator1");
            // 
            // toolStripMenuItem_table
            // 
            toolStripMenuItem_table.Name = "toolStripMenuItem_table";
            resources.ApplyResources(toolStripMenuItem_table, "toolStripMenuItem_table");
            // 
            // ohaDateTimePicker
            // 
            ContextMenuStrip = contextMenuStrip1;
            ValueChanged += ohaDateTimePicker_ValueChanged;
            DropDown += ohaDateTimePicker_DropDown;
            BindingContextChanged += ohaDateTimePicker_BindingContextChanged;
            Enter += ohaDateTimePicker_Enter;
            KeyUp += ohaDateTimePicker_KeyUp;
            Leave += ohaDateTimePicker_Leave;
            Validating += ohaDateTimePicker_Validating;
            contextMenuStrip1.ResumeLayout(false);
            ResumeLayout(false);
        }

        private void ohaDateTimePicker_BindingContextChanged(object sender, EventArgs e)
        {
            SetBindingContextChanged();
        }

        protected virtual void ohaDateTimePicker_Validating(object sender, CancelEventArgs e)
        {
            //e.Cancel = false;
            //if (m_interner_validating &&
            //    !DB.dbUtilities.is_field_nullable(this) &&
            //    Value == null)
            //{
            //    if (Label != null)
            //        m_ErrorProvider.SetError(this, Properties.Resources.ex0002_01.Replace("%1", Label.Text));
            //    //e.Cancel = true;
            //}
            //else
            //    m_ErrorProvider.Clear();
        }

        private void ohaDateTimePicker_Enter(object sender, EventArgs e)
        {
            Font = new Font(Font, FontStyle.Bold);

            if (Label != null)
            {
                m_old_font = Label.Font;
                Label.Font = new Font(Label.Font, m_old_font.Style | FontStyle.Bold);
            }

            if (Value != null)
            {
                SetToDateTimeValue();
            }
            else
            {
                //this.Value = DateTime.Now;
                SetToNullValue();
            }

            // Is Tablet?
            //if (m_allow_edit &&
            //    TabletPCSupport.IsTabletMode && File.Exists(@"C:\Windows\System32\OSK.exe"))
            //{
            //    TemplateForm.ProcessStartStatic(@"C:\Windows\System32\OSK.exe");
            //}
        }

        private void ohaDateTimePicker_Leave(object sender, EventArgs e)
        {
            Font = new Font(Font, FontStyle.Regular);

            if (Label != null)
            {
                //OnDropDown(e);
                Label.Font = m_old_font;
            }
        }


        protected override void OnKeyDown(KeyEventArgs e)
        {
            DateTimePicker.ohaDateTimePicker_Input frm = null;

            m_drop_down_on = false;

            if (!NormalMode)
            {
                //if (e.Control == false)
                //{
                if ((e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9)
                     || (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9)
                     || e.KeyCode == Keys.H
                     || e.KeyCode == Keys.T
                     || e.KeyCode == Keys.M
                     || e.KeyCode == Keys.W
                     || e.KeyCode == Keys.Y
                    )
                {
                    string tmp = ((Char)e.KeyCode).ToString();

                    if (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9)
                    {
                        if (e.KeyCode == Keys.NumPad0) tmp = ((Char)Keys.D0).ToString();
                        else if (e.KeyCode == Keys.NumPad1) tmp = ((Char)Keys.D1).ToString();
                        else if (e.KeyCode == Keys.NumPad2) tmp = ((Char)Keys.D2).ToString();
                        else if (e.KeyCode == Keys.NumPad3) tmp = ((Char)Keys.D3).ToString();
                        else if (e.KeyCode == Keys.NumPad4) tmp = ((Char)Keys.D4).ToString();
                        else if (e.KeyCode == Keys.NumPad5) tmp = ((Char)Keys.D5).ToString();
                        else if (e.KeyCode == Keys.NumPad6) tmp = ((Char)Keys.D6).ToString();
                        else if (e.KeyCode == Keys.NumPad7) tmp = ((Char)Keys.D7).ToString();
                        else if (e.KeyCode == Keys.NumPad8) tmp = ((Char)Keys.D8).ToString();
                        else if (e.KeyCode == Keys.NumPad9) tmp = ((Char)Keys.D9).ToString();
                    }
                    //this.Value = DBNull.Value;
                    frm = new DateTimePicker.ohaDateTimePicker_Input(tmp, this,
                        Value != null ? (DateTime?)Value : new DateTime?());

                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        Value = frm.Value;
                    }
                }
            }
            else
            {
                base.OnKeyDown(e);
            }

        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (e.ClickedItem == toolStripMenuItem_today)
            {
                Value = DateTime.Today;
            }
            else if (e.ClickedItem == toolStripMenuItem_normal_mode)
            {
                NormalMode = !NormalMode;
                toolStripMenuItem_normal_mode.Checked = !toolStripMenuItem_normal_mode.Checked;
            }
            //else if (e.ClickedItem == toolStripMenuItem_table)
            //{
            //    new mssqlinfo_tables(this).ShowDialog();
            //}
        }

        private void ohaDateTimePicker_DropDown(object sender, EventArgs e)
        {
            DateTime tmp_datetime = DateTime.Now;

            m_drop_down_on = true;

            if (!DateTime.TryParse(base.Text, out tmp_datetime))
            {
                SuspendLayout();


                SendKeys.Send("%{DOWN}");
                base.Value = DateTime.Today;
                Value = DateTime.Today;

                SendKeys.Send("%{UP}");
                SendKeys.Send("%{DOWN}");

                ResumeLayout();
            }
        }

        private void ohaDateTimePicker_KeyUp(object sender, KeyEventArgs e)
        {
            m_drop_down_on = false;

            //if (e.KeyCode == Keys.F10 && GENERAL.TemplateForm.g_admin)
            //{
            //    //Help.ShowHelp(
            //    //    this,
            //    //    Path.Combine(Application.StartupPath, "ohaERP_admin.chm"),
            //    //    HelpNavigator.KeywordIndex,
            //    //    this.Name);
            //    new mssqlinfo_tables(this).ShowDialog();
            //}
            //else if (e.KeyCode == Keys.F1)
            //{
            //    Help.ShowHelp(
            //        this,
            //        Path.Combine(Application.StartupPath, "ohaERP_help.chm"),
            //        HelpNavigator.KeywordIndex,
            //        ohaERP_Globals.cGlobal.g_act_form_for_help);
            //}
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }


        #region ReadOnly

        protected override void WndProc(ref Message m)
        {
            const int WM_PAINT = 0x000F;

            base.WndProc(ref m);

            // When we get a redraw message, ans readonly has been set
            if (m.Msg == WM_PAINT && Enabled == false)
            {
                Graphics g = CreateGraphics();
                Pen p = new Pen(Color.Black, 1);
                Brush b = new SolidBrush(SystemColors.Control);

                g.FillRectangle(b, 1, 1, Width - 2, Height - 2);
                g.DrawString(Text, Font, Brushes.Black, 0, 3);
            }
        }

        #endregion







        private void ohaDateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            if (m_drop_down_on)
                Value = ((System.Windows.Forms.DateTimePicker)sender).Value;
        }
    }
}
