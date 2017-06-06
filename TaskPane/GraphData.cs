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
                dgvGraphData.AutoGenerateColumns = true;
                dgvGraphData.DataSource = Data.GraphDataTable.DefaultView;
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
                this.Rpie.Series[0].Points[MyRandomNumber]["Exploded"] = "False";

                for (int i = 0; i < 360; i++)
                {
                    this.Rpie.Series[0]["PieStartAngle"] = i.ToString();
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(10);
                }
                Random random = new Random();
                int randomNumber = random.Next(1, 38);
                this.Rpie.Series[0].Points[randomNumber]["Exploded"] = "True";
                MyRandomNumber = randomNumber;
                string yourNumber = this.Rpie.Series[0].Points[randomNumber].AxisLabel.ToString();
                MessageBox.Show(yourNumber, "Your number", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

    }
}
