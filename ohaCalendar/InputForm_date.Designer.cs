using System.Data;

namespace ohaCalendar
{
    partial class InputForm_date
    {

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        protected void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InputForm_date));
            this.label1 = new System.Windows.Forms.Label();
            this.button_ok = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.ohaDateTimePicker = new DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.myDataTable)).BeginInit();
            this.SuspendLayout();
            // 
            // errorProvider1
            // 
            resources.ApplyResources(this.errorProvider1, "errorProvider1");
            // 
            // helpProvider1
            // 
            resources.ApplyResources(this.helpProvider1, "helpProvider1");
            // 
            // infoButton
            // 
            resources.ApplyResources(this.infoButton, "infoButton");
            this.infoButton.BackColor = System.Drawing.Color.Transparent;
            this.infoButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.errorProvider1.SetError(this.infoButton, resources.GetString("infoButton.Error"));
            this.infoButton.FlatAppearance.BorderSize = 0;
            this.infoButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Transparent;
            this.infoButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Transparent;
            this.helpProvider1.SetHelpKeyword(this.infoButton, resources.GetString("infoButton.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this.infoButton, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("infoButton.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this.infoButton, resources.GetString("infoButton.HelpString"));
            this.errorProvider1.SetIconAlignment(this.infoButton, ((System.Windows.Forms.ErrorIconAlignment)(resources.GetObject("infoButton.IconAlignment"))));
            this.errorProvider1.SetIconPadding(this.infoButton, ((int)(resources.GetObject("infoButton.IconPadding"))));
            this.helpProvider1.SetShowHelp(this.infoButton, ((bool)(resources.GetObject("infoButton.ShowHelp"))));
            this.toolTip1.SetToolTip(this.infoButton, resources.GetString("infoButton.ToolTip"));
            this.infoButton.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.errorProvider1.SetError(this.label1, resources.GetString("label1.Error"));
            this.helpProvider1.SetHelpKeyword(this.label1, resources.GetString("label1.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this.label1, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("label1.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this.label1, resources.GetString("label1.HelpString"));
            this.errorProvider1.SetIconAlignment(this.label1, ((System.Windows.Forms.ErrorIconAlignment)(resources.GetObject("label1.IconAlignment"))));
            this.errorProvider1.SetIconPadding(this.label1, ((int)(resources.GetObject("label1.IconPadding"))));
            this.label1.Name = "label1";
            this.helpProvider1.SetShowHelp(this.label1, ((bool)(resources.GetObject("label1.ShowHelp"))));
            this.toolTip1.SetToolTip(this.label1, resources.GetString("label1.ToolTip"));
            // 
            // button_ok
            // 
            resources.ApplyResources(this.button_ok, "button_ok");
            this.button_ok.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.errorProvider1.SetError(this.button_ok, resources.GetString("button_ok.Error"));
            this.helpProvider1.SetHelpKeyword(this.button_ok, resources.GetString("button_ok.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this.button_ok, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("button_ok.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this.button_ok, resources.GetString("button_ok.HelpString"));
            this.errorProvider1.SetIconAlignment(this.button_ok, ((System.Windows.Forms.ErrorIconAlignment)(resources.GetObject("button_ok.IconAlignment"))));
            this.errorProvider1.SetIconPadding(this.button_ok, ((int)(resources.GetObject("button_ok.IconPadding"))));
            this.button_ok.Name = "button_ok";
            this.helpProvider1.SetShowHelp(this.button_ok, ((bool)(resources.GetObject("button_ok.ShowHelp"))));
            this.toolTip1.SetToolTip(this.button_ok, resources.GetString("button_ok.ToolTip"));
            this.button_ok.UseVisualStyleBackColor = true;
            // 
            // button_cancel
            // 
            resources.ApplyResources(this.button_cancel, "button_cancel");
            this.button_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.errorProvider1.SetError(this.button_cancel, resources.GetString("button_cancel.Error"));
            this.helpProvider1.SetHelpKeyword(this.button_cancel, resources.GetString("button_cancel.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this.button_cancel, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("button_cancel.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this.button_cancel, resources.GetString("button_cancel.HelpString"));
            this.errorProvider1.SetIconAlignment(this.button_cancel, ((System.Windows.Forms.ErrorIconAlignment)(resources.GetObject("button_cancel.IconAlignment"))));
            this.errorProvider1.SetIconPadding(this.button_cancel, ((int)(resources.GetObject("button_cancel.IconPadding"))));
            this.button_cancel.Name = "button_cancel";
            this.helpProvider1.SetShowHelp(this.button_cancel, ((bool)(resources.GetObject("button_cancel.ShowHelp"))));
            this.toolTip1.SetToolTip(this.button_cancel, resources.GetString("button_cancel.ToolTip"));
            this.button_cancel.UseVisualStyleBackColor = true;
            // 
            // ohaDateTimePicker
            // 
            resources.ApplyResources(this.ohaDateTimePicker, "ohaDateTimePicker");
            this.errorProvider1.SetError(this.ohaDateTimePicker, resources.GetString("ohaDateTimePicker.Error"));
            this.helpProvider1.SetHelpKeyword(this.ohaDateTimePicker, resources.GetString("ohaDateTimePicker.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this.ohaDateTimePicker, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("ohaDateTimePicker.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this.ohaDateTimePicker, resources.GetString("ohaDateTimePicker.HelpString"));
            this.errorProvider1.SetIconAlignment(this.ohaDateTimePicker, ((System.Windows.Forms.ErrorIconAlignment)(resources.GetObject("ohaDateTimePicker.IconAlignment"))));
            this.errorProvider1.SetIconPadding(this.ohaDateTimePicker, ((int)(resources.GetObject("ohaDateTimePicker.IconPadding"))));
            //this.ohaDateTimePicker.Label = null;
            //this.ohaDateTimePicker.Name = "ohaDateTimePicker";
            //this.ohaDateTimePicker.NormalMode = false;
            this.helpProvider1.SetShowHelp(this.ohaDateTimePicker, ((bool)(resources.GetObject("ohaDateTimePicker.ShowHelp"))));
            this.toolTip1.SetToolTip(this.ohaDateTimePicker, resources.GetString("ohaDateTimePicker.ToolTip"));
            this.ohaDateTimePicker.Value = new System.DateTime(2013, 9, 26, 10, 30, 29, 0);
            this.ohaDateTimePicker.ValueChanged += new System.EventHandler(this.ohaDateTimePicker_ValueChanged);
            // 
            // InputForm_date
            // 
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.ohaDateTimePicker);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.button_ok);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.helpProvider1.SetHelpKeyword(this, resources.GetString("$this.HelpKeyword"));
            this.helpProvider1.SetHelpNavigator(this, ((System.Windows.Forms.HelpNavigator)(resources.GetObject("$this.HelpNavigator"))));
            this.helpProvider1.SetHelpString(this, resources.GetString("$this.HelpString"));
            this.Name = "InputForm_date";
            this.helpProvider1.SetShowHelp(this, ((bool)(resources.GetObject("$this.ShowHelp"))));
            this.toolTip1.SetToolTip(this, resources.GetString("$this.ToolTip"));
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.InputForm_FormClosing);
            this.Load += new System.EventHandler(this.InputForm_Load);
            this.Controls.SetChildIndex(this.label1, 0);
            this.Controls.SetChildIndex(this.button_ok, 0);
            this.Controls.SetChildIndex(this.button_cancel, 0);
            this.Controls.SetChildIndex(this.ohaDateTimePicker, 0);
            this.Controls.SetChildIndex(this.infoButton, 0);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.myDataTable)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected System.Windows.Forms.Label label1;
        protected System.Windows.Forms.Button button_ok;
        protected System.Windows.Forms.Button button_cancel;
        private DateTimePicker ohaDateTimePicker;
        private System.ComponentModel.IContainer components;
        ErrorProvider errorProvider1 = new ErrorProvider();
        DataTable myDataTable = new DataTable();
        Button infoButton = new Button();
        HelpProvider helpProvider1= new HelpProvider();
        ToolTip toolTip1 = new ToolTip();
    }
}
