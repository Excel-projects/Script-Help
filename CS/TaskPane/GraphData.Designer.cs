namespace ScriptHelp.TaskPane
{
    partial class GraphData
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
			System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
			System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
			System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
			this.dgvGraphData = new System.Windows.Forms.DataGridView();
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnStart = new System.Windows.Forms.ToolStripButton();
			this.Rpie = new System.Windows.Forms.DataVisualization.Charting.Chart();
			this.dgvGraphDataResults = new System.Windows.Forms.DataGridView();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			((System.ComponentModel.ISupportInitialize)(this.dgvGraphData)).BeginInit();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.Rpie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvGraphDataResults)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// dgvGraphData
			// 
			this.dgvGraphData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dgvGraphData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dgvGraphData.Location = new System.Drawing.Point(186, 728);
			this.dgvGraphData.Name = "dgvGraphData";
			this.dgvGraphData.Size = new System.Drawing.Size(94, 22);
			this.dgvGraphData.TabIndex = 1;
			this.dgvGraphData.Visible = false;
			// 
			// toolStrip1
			// 
			this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnStart});
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Size = new System.Drawing.Size(300, 25);
			this.toolStrip1.TabIndex = 2;
			this.toolStrip1.Text = "toolStrip1";
			// 
			// btnStart
			// 
			this.btnStart.Image = global::ScriptHelp.Properties.Resources.control_play_blue;
			this.btnStart.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnStart.Name = "btnStart";
			this.btnStart.Size = new System.Drawing.Size(51, 22);
			this.btnStart.Text = "Start";
			this.btnStart.ToolTipText = "Would you like to play a game?";
			this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
			// 
			// Rpie
			// 
			chartArea1.Area3DStyle.Enable3D = true;
			chartArea1.Area3DStyle.LightStyle = System.Windows.Forms.DataVisualization.Charting.LightStyle.Realistic;
			chartArea1.Area3DStyle.WallWidth = 10;
			chartArea1.Name = "ChartArea1";
			this.Rpie.ChartAreas.Add(chartArea1);
			legend1.Enabled = false;
			legend1.Name = "Legend1";
			this.Rpie.Legends.Add(legend1);
			this.Rpie.Location = new System.Drawing.Point(0, 28);
			this.Rpie.Name = "Rpie";
			this.Rpie.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
			series1.BorderColor = System.Drawing.Color.Silver;
			series1.ChartArea = "ChartArea1";
			series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
			series1.Font = new System.Drawing.Font("Calibri", 8F);
			series1.LabelAngle = 30;
			series1.LabelForeColor = System.Drawing.Color.White;
			series1.Legend = "Legend1";
			series1.Name = "Series1";
			this.Rpie.Series.Add(series1);
			this.Rpie.Size = new System.Drawing.Size(280, 281);
			this.Rpie.TabIndex = 3;
			this.Rpie.Text = "chart1";
			// 
			// dgvGraphDataResults
			// 
			this.dgvGraphDataResults.AllowUserToAddRows = false;
			this.dgvGraphDataResults.AllowUserToDeleteRows = false;
			this.dgvGraphDataResults.AllowUserToResizeColumns = false;
			this.dgvGraphDataResults.AllowUserToResizeRows = false;
			this.dgvGraphDataResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dgvGraphDataResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dgvGraphDataResults.ColumnHeadersVisible = false;
			this.dgvGraphDataResults.Enabled = false;
			this.dgvGraphDataResults.Location = new System.Drawing.Point(186, 327);
			this.dgvGraphDataResults.Name = "dgvGraphDataResults";
			this.dgvGraphDataResults.RowHeadersVisible = false;
			this.dgvGraphDataResults.ScrollBars = System.Windows.Forms.ScrollBars.None;
			this.dgvGraphDataResults.Size = new System.Drawing.Size(94, 400);
			this.dgvGraphDataResults.TabIndex = 4;
			this.dgvGraphDataResults.TabStop = false;
			this.dgvGraphDataResults.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvGraphDataResults_CellEndEdit);
			this.dgvGraphDataResults.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvGraphDataResults_CellFormatting);
			// 
			// pictureBox1
			// 
			this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.pictureBox1.BackgroundImage = global::ScriptHelp.Properties.Resources.table;
			this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
			this.pictureBox1.Location = new System.Drawing.Point(0, 327);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(180, 423);
			this.pictureBox1.TabIndex = 5;
			this.pictureBox1.TabStop = false;
			// 
			// GraphData
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.dgvGraphDataResults);
			this.Controls.Add(this.Rpie);
			this.Controls.Add(this.toolStrip1);
			this.Controls.Add(this.dgvGraphData);
			this.Name = "GraphData";
			this.Size = new System.Drawing.Size(300, 750);
			((System.ComponentModel.ISupportInitialize)(this.dgvGraphData)).EndInit();
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.Rpie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvGraphDataResults)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvGraphData;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnStart;
        private System.Windows.Forms.DataVisualization.Charting.Chart Rpie;
        private System.Windows.Forms.DataGridView dgvGraphDataResults;
		private System.Windows.Forms.PictureBox pictureBox1;
	}
}
