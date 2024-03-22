namespace ohaCalendar
{
    partial class HolidayInput
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
            yearNumericUpDown = new NumericUpDown();
            label1 = new Label();
            stateComboBox = new ComboBox();
            label2 = new Label();
            button1 = new Button();
            button2 = new Button();
            ((System.ComponentModel.ISupportInitialize)yearNumericUpDown).BeginInit();
            SuspendLayout();
            // 
            // yearNumericUpDown
            // 
            yearNumericUpDown.Location = new Point(72, 21);
            yearNumericUpDown.Maximum = new decimal(new int[] { 9999, 0, 0, 0 });
            yearNumericUpDown.Name = "yearNumericUpDown";
            yearNumericUpDown.Size = new Size(92, 23);
            yearNumericUpDown.TabIndex = 0;
            yearNumericUpDown.ValueChanged += yearNumericUpDown_ValueChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 25);
            label1.Name = "label1";
            label1.Size = new Size(32, 15);
            label1.TabIndex = 1;
            label1.Text = "year:";
            // 
            // stateComboBox
            // 
            stateComboBox.FormattingEnabled = true;
            stateComboBox.Location = new Point(72, 52);
            stateComboBox.Name = "stateComboBox";
            stateComboBox.Size = new Size(237, 23);
            stateComboBox.TabIndex = 2;
            stateComboBox.SelectedValueChanged += stateComboBox_SelectedValueChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 57);
            label2.Name = "label2";
            label2.Size = new Size(35, 15);
            label2.TabIndex = 3;
            label2.Text = "state:";
            // 
            // button1
            // 
            button1.DialogResult = DialogResult.OK;
            button1.Location = new Point(72, 104);
            button1.Name = "button1";
            button1.Size = new Size(135, 23);
            button1.TabIndex = 4;
            button1.Text = "Ok";
            button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            button2.DialogResult = DialogResult.Cancel;
            button2.Location = new Point(213, 104);
            button2.Name = "button2";
            button2.Size = new Size(72, 23);
            button2.TabIndex = 5;
            button2.Text = "Cancel";
            button2.UseVisualStyleBackColor = true;
            // 
            // HolidayInput
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(321, 142);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(label2);
            Controls.Add(stateComboBox);
            Controls.Add(label1);
            Controls.Add(yearNumericUpDown);
            Name = "HolidayInput";
            Text = "HolidayInput";
            Load += HolidayInput_Load;
            ((System.ComponentModel.ISupportInitialize)yearNumericUpDown).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private NumericUpDown yearNumericUpDown;
        private Label label1;
        private ComboBox stateComboBox;
        private Label label2;
        private Button button1;
        private Button button2;
    }
}