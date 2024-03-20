using ohaERP_Library;
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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InputForm_date));
            label1 = new Label();
            button_ok = new Button();
            button_cancel = new Button();
            ohaDateTimePicker = new ohaDateTimePicker();
            SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(label1, "label1");
            label1.Name = "label1";
            // 
            // button_ok
            // 
            button_ok.DialogResult = DialogResult.OK;
            resources.ApplyResources(button_ok, "button_ok");
            button_ok.Name = "button_ok";
            button_ok.UseVisualStyleBackColor = true;
            // 
            // button_cancel
            // 
            button_cancel.DialogResult = DialogResult.Cancel;
            resources.ApplyResources(button_cancel, "button_cancel");
            button_cancel.Name = "button_cancel";
            button_cancel.UseVisualStyleBackColor = true;
            // 
            // ohaDateTimePicker
            // 
            ohaDateTimePicker.Label = null;
            resources.ApplyResources(ohaDateTimePicker, "ohaDateTimePicker");
            ohaDateTimePicker.Name = "ohaDateTimePicker";
            ohaDateTimePicker.NormalMode = false;
            ohaDateTimePicker.Value = new DateTime(2013, 9, 26, 10, 30, 29, 0);
            ohaDateTimePicker.ValueChanged += ohaDateTimePicker_ValueChanged;
            // 
            // InputForm_date
            // 
            resources.ApplyResources(this, "$this");
            Controls.Add(ohaDateTimePicker);
            Controls.Add(button_cancel);
            Controls.Add(button_ok);
            Controls.Add(label1);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Name = "InputForm_date";
            FormClosing += InputForm_FormClosing;
            Load += InputForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        protected System.Windows.Forms.Label label1;
        protected System.Windows.Forms.Button button_ok;
        protected System.Windows.Forms.Button button_cancel;
        private ohaDateTimePicker ohaDateTimePicker;
        private System.ComponentModel.IContainer components;
        ErrorProvider errorProvider1 = new ErrorProvider();
        DataTable myDataTable = new DataTable();
        Button infoButton = new Button();
        HelpProvider helpProvider1= new HelpProvider();
        ToolTip toolTip1 = new ToolTip();
    }
}
