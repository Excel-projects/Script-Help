using System;
using System.Data.SqlServerCe;
using System.Linq;
using System.Windows.Forms;
using ScriptHelp.Scripts;

namespace ScriptHelp.TaskPane
{
    /// <summary>
    /// TableData TaskPane
    /// </summary>
    public partial class TableData : UserControl
    {
        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public TableData()
        {
            InitializeComponent();
            try
            {
                dgvList.AutoGenerateColumns = true;
                string tableName = Ribbon.AppVariables.TableName;
                this.Text = "List of " + tableName;
                switch (tableName)
                {
                    case "TableAlias":
                        dgvList.DataSource = Data.TableAliasTable;
                        break;
                    case "DateFormat":
                        dgvList.DataSource = Data.DateFormatTable;
                        break;
                }
                this.dgvList.Columns[0].Width = dgvList.Width - 75;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// Save the changed values in the table
        /// </summary>
        /// <param name="sender">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                this.Validate();
                if (dgvList.IsCurrentRowDirty || dgvList.IsCurrentCellDirty)
                {
                    dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dgvList.EndEdit();
                }

                string tableName = Ribbon.AppVariables.TableName;
                string sql = "SELECT * FROM " + tableName;
                //string sql = "SELECT * FROM @tableName";
                SqlCeConnection cn = new SqlCeConnection(Data.Connection());
                SqlCeCommandBuilder scb = default(SqlCeCommandBuilder);
                SqlCeDataAdapter sda = new SqlCeDataAdapter(sql, cn);
                //TODO: use this parameterized query...
                //sda.SelectCommand.Parameters.AddWithValue("@tableName", tableName);
                //sda.SelectCommand.Parameters.AddWithValue("@tableName", System.Data.SqlDbType.NVarChar).Value =  tableName;
                //sda.SelectCommand.Parameters.Add(new SqlCeParameter
                //{
                //      ParameterName = "@tableName"
                //    , Value = tableName
                //    , SqlDbType = System.Data.SqlDbType.NVarChar
                //    , Size = 50 
                //});

                sda.TableMappings.Add("Table", tableName);
                scb = new SqlCeCommandBuilder(sda);
                switch (tableName)
                {
                    case "TableAlias":
                        sda.Update(Data.TableAliasTable);
                        Data.CreateTableAliasTable();
                        break;
                    case "DateFormat":
                        sda.Update(Data.DateFormatTable);
                        Data.CreateDateFormatTable();
                        break;
                }
                Ribbon.ribbonref.InvalidateRibbon();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

    }
}
