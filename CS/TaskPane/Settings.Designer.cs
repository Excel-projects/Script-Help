namespace ScriptHelp.TaskPane
{
    partial class Settings
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pgdSettings = new System.Windows.Forms.PropertyGrid();
            this.SuspendLayout();
            // 
            // pgdSettings
            // 
            this.pgdSettings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pgdSettings.Location = new System.Drawing.Point(0, 0);
            this.pgdSettings.Name = "pgdSettings";
            this.pgdSettings.Size = new System.Drawing.Size(650, 750);
            this.pgdSettings.TabIndex = 2;
            this.pgdSettings.PropertyValueChanged += new System.Windows.Forms.PropertyValueChangedEventHandler(this.pgdSettings_PropertyValueChanged);
            // 
            // Settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pgdSettings);
            this.Name = "Settings";
            this.Size = new System.Drawing.Size(650, 750);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.PropertyGrid pgdSettings;
    }
}
