namespace SqlHelp.TaskPane
{
    partial class DocumentViewer
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
            this.ppcDocumentViewer = new System.Windows.Forms.PrintPreviewControl();
            this.SuspendLayout();
            // 
            // ppcDocumentViewer
            // 
            this.ppcDocumentViewer.Location = new System.Drawing.Point(3, 3);
            this.ppcDocumentViewer.Name = "ppcDocumentViewer";
            this.ppcDocumentViewer.Size = new System.Drawing.Size(650, 750);
            this.ppcDocumentViewer.TabIndex = 0;
            // 
            // DocumentViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.ppcDocumentViewer);
            this.Name = "DocumentViewer";
            this.Size = new System.Drawing.Size(656, 756);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PrintPreviewControl ppcDocumentViewer;
    }
}
