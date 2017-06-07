using System;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Windows.Forms;
using ScriptHelp.Scripts;

namespace ScriptHelp.TaskPane
{
    /// <summary>
    /// GraphData TaskPane
    /// </summary>
    public partial class GraphData : UserControl
    {
        /// <summary>
        /// random number stored for multiple processes
        /// </summary>
        public int MyRandomNumber;

        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public GraphData()
        {
            InitializeComponent();
            try
            {
                //dgvGraphData.AutoGenerateColumns = true;
                //dgvGraphData.DataSource = Data.GraphDataTable.DefaultView;
                this.Rpie.Series[0].XValueMember = "NBR_VALUE";
                this.Rpie.Series[0].YValueMembers = "VALUE";
                this.Rpie.DataSource = Data.GraphDataTable;
                this.Rpie.DataBind();

                foreach (DataRow row in Data.GraphDataTable.Rows)
                {
                    int orderNbr = orderNbr = Convert.ToInt32(row["ORDR_NBR"].ToString());
                    orderNbr = orderNbr - 1;
                    System.Drawing.Color c = System.Drawing.ColorTranslator.FromHtml(row["COLOR_ID"].ToString());
                    this.Rpie.Series[0].Points[orderNbr].Color = c;
                    Application.DoEvents();
                }
                RefreshResultsToGrid();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// To start the procedure
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                this.btnStart.Enabled = false;
                this.Rpie.Series[0].Points[MyRandomNumber]["Exploded"] = "False";
                Random random = new Random();
                int totalNbr = Data.GraphDataTable.Rows.Count;
                MyRandomNumber = random.Next(1, totalNbr);
                for (int i = 0; i < MyRandomNumber + 100; i++)
                {
                    this.Rpie.Series[0]["PieStartAngle"] = i.ToString();
                    Application.DoEvents();
                }
                InsertResultsToDataTable();
                RefreshResultsToGrid();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                this.btnStart.Enabled = true;
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                this.Rpie.Series[0].Points[MyRandomNumber]["Exploded"] = "True";

            }

        }

        /// <summary>
        /// Update the cell formatting
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void dgvGraphDataResults_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                DataGridView dgv = sender as DataGridView;
                if (dgv.Columns[e.ColumnIndex].Name.Equals("COLOR_ID") && e.Value != null)
                {
                    e.CellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml(e.Value.ToString());
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.ColorTranslator.FromHtml(e.Value.ToString());
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// Invalidate the grid formatting
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void dgvGraphDataResults_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dgvGraphDataResults.Invalidate();
        }

        /// <summary>
        /// Insert the results to the data table
        /// </summary>
        private void InsertResultsToDataTable()
        {
            try
            {
                //TODO: should be querying from the GraphData table instead of from the control
                string number = this.Rpie.Series[0].Points[MyRandomNumber].AxisLabel.ToString();
                string color = System.Drawing.ColorTranslator.ToHtml(this.Rpie.Series[0].Points[MyRandomNumber].Color);
                //MessageBox.Show(number, "Your number", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                string connection = Scripts.Data.Connection();
                string query = ("INSERT INTO GraphDataResults (NBR_VALUE, COLOR_ID) Values(@yourNumber, @yourColor)");
                using (SqlCeConnection cn = new SqlCeConnection(connection))
                {
                    using (SqlCeCommand cmd = new SqlCeCommand(query, cn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@yourNumber", number);
                        cmd.Parameters.AddWithValue("@yourColor", color);
                        cn.Open();
                        cmd.ExecuteNonQuery();
                        cn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// Refresh the results to the grid
        /// </summary>
        private void RefreshResultsToGrid()
        {
            try
            {
                string sql = "SELECT NBR_VALUE, COLOR_ID, RESULT_ID FROM GraphDataResults";
                System.Data.DataTable dt = new System.Data.DataTable();
                using (var da = new SqlCeDataAdapter(sql, Scripts.Data.Connection()))
                {
                    da.Fill(dt);
                }
                dt.DefaultView.Sort = "[RESULT_ID] DESC";
                dgvGraphDataResults.DataSource = dt.DefaultView;
                dgvGraphDataResults.AutoGenerateColumns = false;
                dgvGraphDataResults.Columns.Clear();
                dgvGraphDataResults.AllowUserToAddRows = false;

                DataGridViewTextBoxColumn txtResultColor = new DataGridViewTextBoxColumn();
                txtResultColor.Width = 0;
                txtResultColor.DataPropertyName = "COLOR_ID";
                txtResultColor.Name = "COLOR_ID";
                txtResultColor.Visible = false;
                dgvGraphDataResults.Columns.Add(txtResultColor);

                DataGridViewTextBoxColumn txtResultNumber = new DataGridViewTextBoxColumn();
                txtResultNumber.Width = 100;
                txtResultNumber.DataPropertyName = "NBR_VALUE";
                txtResultNumber.Name = "NBR_VALUE";
                txtResultNumber.HeaderText = "Results";
                txtResultNumber.Visible = true;
                txtResultNumber.ReadOnly = true;
                dgvGraphDataResults.Columns.Add(txtResultNumber);

                dgvGraphDataResults.Columns[1].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                dgvGraphDataResults.CellFormatting += dgvGraphDataResults_CellFormatting;
                dgvGraphDataResults.CellEndEdit += dgvGraphDataResults_CellEndEdit;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

    }
}
