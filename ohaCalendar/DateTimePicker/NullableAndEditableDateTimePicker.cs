using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Runtime.InteropServices;

namespace ohaERP_Library
{

    /// <summary>

    /// The following code is written from scratch, but was inspired by 

    /// the NullableDateTimePicker by Claudio Grazioli, http://www.grazioli.ch

    /// 

    /// This Extended DateTimePicker adds useful functionality to the DateTimePicker class espicially for Database Apllications

    /// 1. I added the abilty to set the DTP to null with the following features

    ///     a. NullText Property that can be used to display Text of your choise when DTP is null

    ///     b. Ability to type dates or times into the DTP when set to null

    ///     c. BackSpace and Delete keystrokes set the DTP to null

    /// 2. I added a ReadOnly Mode in which the UDTP hides itself and decorates a ReadOnly TextBox in its place

    ///     a. By Decorating a TextBox we get all black easily readable text.

    ///     b. The ReadOnly TextBox has automatic Clipboard access built in which is useful.

    ///     c. I added a TabStopWhenReadOnly Property that allows the ReadOnly Mode to mimic the UDTP Tabstop property or not;

    ///        If TabStopWhenReadOnly is true then the ReadOnly Mode's TabStop will be whatever the UDTP Tabstop Property is.

    ///        If TabStopWhenReadOnly is false then the the ReadOnly Mode's TabStop is false. 

    /// </summary>

    public partial class UltraDateTimePicker : System.Windows.Forms.DateTimePicker, ISupportInitialize
    {

        #region Member Variables Added To Allow Null Values

        //Format and CustomForamt are shadowed since base.Format is always Custom

        //and base.CustomFormat is used in setFormat to show the intended _Format

        //You have to keep base.Format set to Custom to avoid superfluous ValueChanged

        //events from occuring.

        private DateTimePickerFormat _Format; // Variable to store 'Format'

        private string _CustomFormat; //Variable to store 'CustomFormat'



        private string _nullText = "";  //Variable to store null Display Text



        #endregion



        #region Member Variables Added To Enable ReadOnly Mode



        private bool _readOnly = false;//Flag to denote the UDTP is in ReadOnly Mode

        private bool _visible = true; //Overridden to show the proper Display for Readonly Mode

        private bool _tabStopWhenReadOnly = false; //Variable to store whether or not the UDTP is a TabStop when in ReadOnly Mode

        private System.Windows.Forms.TextBox _textBox;//TextBox Decorated when in ReadOnly Mode





        #endregion



        #region Constructor

        /// <summary>

        /// Basic Constructer + ReadOnly _textBox initialization

        /// </summary>

        public UltraDateTimePicker()
        {

            //InitializeComponent();

            initTextBox();

            base.Format = DateTimePickerFormat.Custom;

            _Format = DateTimePickerFormat.Long;

            if (DesignMode)

                setFormat();

        }



        #endregion



        #region ISupportInitialize Members

        private bool initializing = true;

        public void BeginInit()
        {

            // TODO: Add UserControl1.BeginInit implementation    

            initializing = true;

        }

        public void EndInit()
        {

            base.Value = DateTime.Today;//Default the value to Today(makes me happy, but not necessary)

            initializing = false;

            if (DesignMode)

                return;

            if (Parent.GetType() == typeof(TableLayoutPanel))
            {

                TableLayoutPanelCellPosition cP = ((TableLayoutPanel)Parent).GetPositionFromControl(this);

                ((TableLayoutPanel)Parent).Controls.Add(_textBox, cP.Column, cP.Row);

                ((TableLayoutPanel)Parent).SetColumnSpan(_textBox, ((TableLayoutPanel)Parent).GetColumnSpan(this));

                _textBox.Anchor = Anchor;

            }

            //I added special logic here to handle positioning the _TextBox when the UDTP is in a FlowLayoutPanel

            else if (Parent.GetType() == typeof(FlowLayoutPanel))
            {

                ((FlowLayoutPanel)Parent).Controls.Add(_textBox);

                ((FlowLayoutPanel)Parent).Controls.SetChildIndex(_textBox, ((FlowLayoutPanel)Parent).Controls.IndexOf(this));

                _textBox.Anchor = Anchor;

            }

            else //not a TableLayoutPanel or FlowLayoutPanel so just assign the parent
            {

                _textBox.Parent = Parent;

                _textBox.Anchor = Anchor;

            }





            //I use the following block of code to walk up the parent-child 

            //chain and find the first member that has a Load event that I can attach to

            //I set the visiblilty during this event so that Databinding will work correctly

            //otherwise the UDTP will fail to bind properly if its visibility is false during the

            //Load event.(Strange but true, has to do with hidden controls not binding for performance reasons)

            Control parent = this;

            bool foundLoadingParent = false;



            do
            {

                parent = parent.Parent;

                if (parent.GetType().IsSubclassOf(typeof(UserControl)))
                {

                    ((UserControl)parent).Load += new EventHandler(UltraDateTimePicker_Load);

                    foundLoadingParent = true;

                }

                else if (parent.GetType().IsSubclassOf(typeof(Form)))
                {

                    ((Form)parent).Load += new EventHandler(UltraDateTimePicker_Load);

                    foundLoadingParent = true;

                }

            }

            while (!foundLoadingParent);

        }



        void UltraDateTimePicker_Load(object sender, EventArgs e)
        {

            setVisibility();

        }

        #endregion



        #region Public Properties Modified/Added To Allow Null Values



        private bool _isNull = false;



        /// <summary>

        /// Modified Value Propety now of type Object and uses MinDate to mark null values

        /// </summary>

        new public Object Value
        {

            get
            {

                //if (this.MinDate == base.Value) 
                //Check to see if set to MinDate(null), return null or base.Value accordingly

                if (_isNull)

                    return null;

                else

                    return base.Value;

            }

            set
            {

                if (value == null || value == DBNull.Value) //Check for null assignment
                {

                    if (!_isNull) //If not already null set to null and fire event
                    {

                        _isNull = true;

                        OnValueChanged(EventArgs.Empty);

                    }

                }

                else //Value is nto null
                {

                    if (_isNull && base.Value == (DateTime)value)//if null and value matches base.value take out of null and fire event

                    //(null->value needs a value changed even though base.Value did not change)
                    {

                        _isNull = false;

                        OnValueChanged(EventArgs.Empty);

                    }

                    else//change to the new value(changed event fires from base class
                    {

                        _isNull = false;

                        base.Value = (DateTime)value;

                    }

                }

                setFormat();//refresh format

                _textBox.Text = Text;

            }

        }



        /// <summary>

        /// NullText property is used to access/change the Text shown when the UDTP is null

        /// </summary>

        #region DesignerModifiers

        [Browsable(true)]

        [Category("Behavior")]

        [Description("Text shown when DateTime is 'null'")]

        [DefaultValue("")]

        #endregion

        public string NullText
        {

            get { return _nullText; }

            set { _nullText = value; }

        }



        /// <summary>

        /// Modified Format Property stores the assigned Format and the propagates the change to base.CustomFormat

        /// </summary>

        #region DesignerModifiers

        [Browsable(true)]

        [DefaultValue(DateTimePickerFormat.Long), TypeConverter(typeof(Enum))]

        #endregion

        new public DateTimePickerFormat Format
        {

            get { return _Format; }

            set
            {

                _Format = value;

                setFormat();

            }

        }



        private void setFormat()
        {

            base.CustomFormat = null;//Resets the CustomFormat(bookkeeping)

            if (_isNull)//If null apply NullText to the UDTP
            {

                base.CustomFormat = String.Concat("'", NullText, "'");

            }

            else
            {

                //The Following is used to get a string representation ot the current UDTP Format

                //And then set the CustomFormat to match the intended format

                CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;

                DateTimeFormatInfo dTFormatInfo = cultureInfo.DateTimeFormat;

                switch (_Format)
                {

                    case DateTimePickerFormat.Long:

                        base.CustomFormat = dTFormatInfo.LongDatePattern;

                        break;

                    case DateTimePickerFormat.Short:

                        base.CustomFormat = dTFormatInfo.ShortDatePattern;

                        break;

                    case DateTimePickerFormat.Time:

                        base.CustomFormat = dTFormatInfo.ShortTimePattern;

                        break;

                    case DateTimePickerFormat.Custom:

                        base.CustomFormat = _CustomFormat;

                        break;

                }

            }

        }



        /// <summary>

        /// Modified CustomFormat Property stores the assigned CustomFormat when null, otherwise functions as normal

        /// </summary>

        new public string CustomFormat
        {

            get { return _CustomFormat; }

            set
            {

                _CustomFormat = value;

                setFormat();

            }

        }



        #endregion



        #region Public Properties Modified/Added To Enable ReadOnly Mode



        /// <summary>

        /// Sets the ReadOnly Property and then call the appropriate Display Function

        /// If in Design Mode then return(If we dont do this then the Control could Disappear )

        /// </summary>

        #region DesignerModifiers

        [Browsable(true)]

        [Category("Behavior")]

        [Description("Displays Control as ReadOnly(Black on Gray) if 'true'")]

        [DefaultValue(false)]

        #endregion

        public bool ReadOnly
        {

            get { return _readOnly; }

            set
            {

                _readOnly = value;

                setVisibility();

            }

        }



        /// <summary>

        /// This useful property allows you to say wheher or not you want 

        /// the TexBox to Mimic the DTP's TabStop value when in ReadOnly Mode.

        /// I personally found this useful on Data entry forms. The default is false,

        /// which means when in ReadOnly Mode you cannot tab into it so Tab will skip

        /// ReadOnly Pickers.

        /// </summary>

        #region DesignerModifiers

        [Category("Behavior")]

        [DefaultValue(false)]

        [Browsable(true)]

        [EditorBrowsable(EditorBrowsableState.Always)]

        #endregion

        public bool TabstopWhenReadOnly
        {

            get { return _tabStopWhenReadOnly; }

            set
            {

                _tabStopWhenReadOnly = value;

                _textBox.TabStop = (_tabStopWhenReadOnly && TabStop); //TextBox is a Tabstop only if mimicing and DTP is a TabStop

            }

        }



        /// <summary>

        /// Modified the TabStop Property to support the added TabStopWhenReadOnly property

        /// </summary>

        new public bool TabStop
        {

            get { return base.TabStop; }

            set
            {

                base.TabStop = value;

                _textBox.TabStop = (_tabStopWhenReadOnly && base.TabStop);

            }

        }



        /// <summary>

        /// Sets the Visible Property and then call the appropriate Display Function

        /// If in Design Mode then return(If we dont do this then the Control could Disappear )

        /// </summary>

        new public bool Visible
        {

            get { return _visible; }

            set
            {

                _visible = value;

                setVisibility();

            }

        }



        #endregion



       

        /// <summary>

        /// Used to change the UDTP.Value on Closeup(Without this code Closeup only changes the base.Value)

        /// </summary>

        /// <param name="e"></param>



        protected override void WndProc(ref Message m)
        {

            if (m.Msg == 0x4e)
            {

                NMHDR nm = (NMHDR)m.GetLParam(typeof(NMHDR));

                if (nm.Code == -746 || nm.Code == -722)

                    Value = base.Value;//propagate change form base to UTDP

            }

            base.WndProc(ref m);

        }



        [StructLayout(LayoutKind.Sequential)]

        private struct NMHDR
        {

            public IntPtr HwndFrom;

            public int IdFrom;

            public int Code;

        }



        ///<summary>

        ///Sets UDTP Value to null when Delete or Backspace is pressed

        ///</summary>

        ///<param name="e"></param>

        protected override void OnKeyUp(KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)

                Value = null;

            base.OnKeyUp(e);

        }





        /// <summary>

        /// When Null and a Number is pressed this method takes the UDTP out of Null Mode

        /// and resends the pressed key for timign reasons

        /// </summary>

        /// <param name="e"></param>

        protected override void OnKeyPress(KeyPressEventArgs e)
        {

            base.OnKeyPress(e);

            if (_isNull && Char.IsDigit(e.KeyChar))
            {

                Value = base.Value;

                e.Handled = true;

                SendKeys.Send("{RIGHT}");

                SendKeys.Send(e.KeyChar.ToString());

            }

            else

                base.OnKeyPress(e);

        }



        



        #region OnXXXX() Modified To Enable ReadOnly Mode



        /// <summary>

        /// Refreshes the Visibility of the Control if the Parent Changes(So the _TextBox get moved and Redrawn)

        /// If in Design Mode then return(If we dont do this then the Control could Disappear )

        /// </summary>

        /// <param name="e"></param>

        protected override void OnParentChanged(EventArgs e)
        {

            base.OnParentChanged(e);

            if (DesignMode || initializing)

                return;

            updateReadOnlyTextBoxParent();//update the _TextBox parent

            setVisibility(); //Reset Visibilty for new parent

        }



        /// <summary>

        /// Propagates Location to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnLocationChanged(EventArgs e)
        {

            base.OnLocationChanged(e);

            _textBox.Location = Location;

        }



        /// <summary>

        /// Propagates Size to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnSizeChanged(EventArgs e)
        {

            base.OnSizeChanged(e);

            _textBox.Size = Size;

        }



        /// <summary>

        /// Propagates Size to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnResize(EventArgs e)
        {

            base.OnResize(e);

            _textBox.Size = Size;

        }



        /// <summary>

        /// Propagates Dock to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnDockChanged(EventArgs e)
        {

            base.OnDockChanged(e);

            _textBox.Dock = Dock;

        }



        /// <summary>

        /// Propagates RightToLeft to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnRightToLeftChanged(EventArgs e)
        {

            base.OnRightToLeftChanged(e);

            _textBox.RightToLeft = RightToLeft;

        }



        /// <summary>

        /// Propagates TabStop to the _TextBox if TabStopWhenReadOnly == true

        /// </summary>

        /// <param name="e"></param>

        protected override void OnTabStopChanged(EventArgs e)
        {

            base.OnTabStopChanged(e);

            _textBox.TabStop = _tabStopWhenReadOnly && TabStop;

        }



        /// <summary>

        /// Propagates TabIndex to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnTabIndexChanged(EventArgs e)
        {

            base.OnTabIndexChanged(e);

            _textBox.TabIndex = TabIndex;

        }



        ////// <summary>

        /// Propagates Font to the _TextBox

        /// </summary>

        /// <param name="e"></param>

        protected override void OnFontChanged(EventArgs e)
        {

            base.OnFontChanged(e);

            _textBox.Font = Font;

        }



        #endregion



        #region Private Methods Added To Enable ReadOnly Mode



        /// <summary>

        /// Added to initialize the _textBox to the default values to match the DTP

        /// </summary>

        private void initTextBox()
        {

            if (DesignMode)

                return;

            _textBox = new System.Windows.Forms.TextBox();

            _textBox.ReadOnly = true;

            _textBox.Location = Location;

            _textBox.Size = Size;

            _textBox.Dock = Dock;

            _textBox.Anchor = Anchor;

            _textBox.RightToLeft = RightToLeft;

            _textBox.Font = Font;

            _textBox.TabStop = TabStop;

            _textBox.TabIndex = TabIndex;

            _textBox.Visible = false;

            _textBox.Parent = Parent;

        }



        private void setVisibility()
        {

            if (DesignMode || initializing)//Dont actually change the visibility if in Design Mode

                return;

            if (_visible)
            {

                if (_readOnly)

                    showTextBox();//If Visible and Readonly Show TextBox

                else

                    showDTP();//If Visible and NOT ReadOnly Show DateTimePicker

            }

            else
            {

                showNone(); //If Not Visible Show Neither

            }

        }



        private void showTextBox()
        {

            base.Visible = false;

            _textBox.Visible = true;

            _textBox.TabStop = _tabStopWhenReadOnly && TabStop;

        }



        private void showDTP()
        {

            _textBox.Visible = false;

            base.Visible = true;

        }



        private void showNone()
        {

            _textBox.Visible = false;

            base.Visible = false;

        }



        private void updateReadOnlyTextBoxParent()
        {

            if (Parent == null) //If UTDP.Parent == null, set _textBox.Parent == null and return 
            {

                _textBox.Parent = null;

                return;

            }

            if (_textBox.Parent != Parent) //If the Parents DO NOT already match
            {

                //I Added Special logic here to handle positioning the _TextBox when the UDTP is in a TableLayoutPanel

                if (Parent.GetType() == typeof(TableLayoutPanel))
                {

                    TableLayoutPanelCellPosition cP = ((TableLayoutPanel)Parent).GetPositionFromControl(this);

                    ((TableLayoutPanel)Parent).Controls.Add(_textBox, cP.Column, cP.Row);

                    ((TableLayoutPanel)Parent).SetColumnSpan(_textBox, ((TableLayoutPanel)Parent).GetColumnSpan(this));

                    _textBox.Anchor = Anchor;

                }

                //I added special logic here to handle positioning the _TextBox when the UDTP is in a FlowLayoutPanel

                else if (Parent.GetType() == typeof(FlowLayoutPanel))
                {

                    ((FlowLayoutPanel)Parent).Controls.Add(_textBox);

                    ((FlowLayoutPanel)Parent).Controls.SetChildIndex(_textBox, ((FlowLayoutPanel)Parent).Controls.IndexOf(this));

                    _textBox.Anchor = Anchor;

                }

                else //not a TableLayoutPanel or FlowLayoutPanel so just assign the parent
                {

                    _textBox.Parent = Parent;

                    _textBox.Anchor = Anchor;

                }

            }

        }



        #endregion



        #region Public Methods Overriden To Enable ReadOnly Mode



        new public void Show()
        {

            Visible = true;

        }



        new public void Hide()
        {

            Visible = false;

        }



        #endregion

    }

}


