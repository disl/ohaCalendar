namespace ohaERP_Library.DateTimePicker
{
    partial class ohaDateTimePickerWithTime
    {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            TimeDateTimePicker = new ohaDateTimePicker();
            DateDateTimePicker = new ohaDateTimePicker();
            SuspendLayout();
            // 
            // TimeDateTimePicker
            // 
            TimeDateTimePicker.CustomFormat = "HH:mm";
            TimeDateTimePicker.Format = DateTimePickerFormat.Custom;
            TimeDateTimePicker.Label = null;
            TimeDateTimePicker.Location = new Point(118, 0);
            TimeDateTimePicker.Name = "TimeDateTimePicker";
            TimeDateTimePicker.NormalMode = false;
            TimeDateTimePicker.ShowUpDown = true;
            TimeDateTimePicker.Size = new Size(77, 21);
            TimeDateTimePicker.TabIndex = 1;
            TimeDateTimePicker.Value = new DateTime(2013, 11, 12, 13, 5, 3, 0);
            TimeDateTimePicker.ValueChanged += TimeDateTimePicker_ValueChanged;
            TimeDateTimePicker.Validated += TimeDateTimePicker_Validated;
            // 
            // DateDateTimePicker
            // 
            DateDateTimePicker.Label = null;
            DateDateTimePicker.Location = new Point(0, 0);
            DateDateTimePicker.Name = "DateDateTimePicker";
            DateDateTimePicker.NormalMode = false;
            DateDateTimePicker.Size = new Size(112, 21);
            DateDateTimePicker.TabIndex = 0;
            DateDateTimePicker.Value = new DateTime(2013, 11, 12, 13, 5, 3, 0);
            DateDateTimePicker.ValueChanged += DateDateTimePicker_ValueChanged;
            // 
            // ohaDateTimePickerWithTime
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            Controls.Add(TimeDateTimePicker);
            Controls.Add(DateDateTimePicker);
            Font = new Font("Arial", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MaximumSize = new Size(195, 21);
            MinimumSize = new Size(195, 21);
            Name = "ohaDateTimePickerWithTime";
            Size = new Size(195, 21);
            //BindingContextChanged += ohaDateTimePickerWithTime_BindingContextChanged;
            Enter += ohaDateTimePickerWithTime_Enter;
            Leave += ohaDateTimePickerWithTime_Leave;
            ResumeLayout(false);
        }

        #endregion

        public ohaDateTimePicker DateDateTimePicker;
        public ohaDateTimePicker TimeDateTimePicker;

    }
}
