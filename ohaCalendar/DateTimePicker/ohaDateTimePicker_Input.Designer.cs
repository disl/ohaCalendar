namespace ohaERP_Library.DateTimePicker
{
    partial class ohaDateTimePicker_Input
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ohaDateTimePicker_Input));
            this.manualTextBox = new System.Windows.Forms.TextBox();
            this.resultLabel = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new ohaERP_Library.ohaDateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.stateTextBox = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // manualTextBox
            // 
            resources.ApplyResources(this.manualTextBox, "manualTextBox");
            this.manualTextBox.Name = "manualTextBox";
            this.manualTextBox.TextChanged += new System.EventHandler(this.manualTextBox_TextChanged);
            this.manualTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.manualTextBox_KeyDown);
            this.manualTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.manualTextBox_KeyPress);
            // 
            // resultLabel
            // 
            resources.ApplyResources(this.resultLabel, "resultLabel");
            this.resultLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.resultLabel.Name = "resultLabel";
            // 
            // dateTimePicker1
            // 
            resources.ApplyResources(this.dateTimePicker1, "dateTimePicker1");
            this.dateTimePicker1.Label = null;
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.NormalMode = false;
            this.dateTimePicker1.TabStop = false;
            this.dateTimePicker1.Value = new System.DateTime(2015, 6, 25, 9, 15, 59, 0);
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            this.dateTimePicker1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dateTimePicker1_KeyDown);
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // stateTextBox
            // 
            resources.ApplyResources(this.stateTextBox, "stateTextBox");
            this.stateTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.stateTextBox.ForeColor = System.Drawing.Color.Red;
            this.stateTextBox.Name = "stateTextBox";

            this.stateTextBox.TextChanged += new System.EventHandler(this.stateTextBox_TextChanged);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";

            // 
            // ohaDateTimePicker_Input
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.ControlBox = false;
            this.Controls.Add(this.manualTextBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.stateTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.resultLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "ohaDateTimePicker_Input";
            this.TopMost = true;
            this.Deactivate += new System.EventHandler(this.ohaDateTimePicker_Input_Deactivate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ohaDateTimePicker_Input_FormClosing);
            this.Load += new System.EventHandler(this.ohaDateTimePicker_Input_Load);

            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox manualTextBox;
        private System.Windows.Forms.Label resultLabel;
        private ohaDateTimePicker  dateTimePicker1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label stateTextBox;
        private System.Windows.Forms.Label label4;
    }
}