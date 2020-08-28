using System;
using System.IO;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Security.AccessControl;
using System.Windows.Forms;

namespace ScriptHelp.Scripts
{
    class Data
    {

        const string dataFolder = "App_Data";

        public static string Connection()
        {
            Data.SetUserPath();
            string databaseFile = "Data Source=" + Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".sdf");
            return databaseFile;
        }


        public static DataTable TableAliasTable = new DataTable();
        public static DataTable DateFormatTable = new DataTable();
        public static DataTable TimeFormatTable = new DataTable();
        public static DataTable GraphDataTable = new DataTable();

        public static void CreateTableAliasTable()
        {
            try
            {
                string tableName = "TableAlias";
                string columnName = "TableName";
                string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;
                dynamic dcTableName = new DataColumn(columnName, typeof(string));
                TableAliasTable.Rows.Clear();
                DataColumnCollection columns = TableAliasTable.Columns;
                if (columns.Contains(columnName) == false)
                {
                    TableAliasTable.Columns.Add(dcTableName);
                }

                using (var da = new SqlCeDataAdapter(sql, Connection()))
                {
                    da.Fill(TableAliasTable);
                }
                TableAliasTable.DefaultView.Sort = columnName + " asc";
                TableAliasTable.TableName = tableName;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        public static void CreateDateFormatTable()
        {
            try
            {
                string tableName = "DateFormat";
                string columnName = "FormatString";
                string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;
                dynamic dcFormatString = new DataColumn(columnName, typeof(string));
                DateFormatTable.Rows.Clear();
                DataColumnCollection columns = DateFormatTable.Columns;
                if (columns.Contains(columnName) == false)
                {
                    DateFormatTable.Columns.Add(dcFormatString);
                }

                using (var da = new SqlCeDataAdapter(sql, Connection()))
                {
                    da.Fill(DateFormatTable);
                }
                DateFormatTable.DefaultView.Sort = columnName + " asc";
                DateFormatTable.TableName = tableName;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        public static void CreateTimeFormatTable()
        {
            try
            {
                string tableName = "TimeFormat";
                string columnName = "FormatString";
                string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;
                dynamic dcFormatString = new DataColumn(columnName, typeof(string));
                TimeFormatTable.Rows.Clear();
                DataColumnCollection columns = TimeFormatTable.Columns;
                if (columns.Contains(columnName) == false)
                {
                    TimeFormatTable.Columns.Add(dcFormatString);
                }

                using (var da = new SqlCeDataAdapter(sql, Connection()))
                {
                    da.Fill(TimeFormatTable);
                }
                TimeFormatTable.DefaultView.Sort = columnName + " asc";
                TimeFormatTable.TableName = tableName;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        public static void CreateGraphDataTable()
        {
            try
            {
                string columnName = "ORDR_NBR";
                dynamic dcFormatString = new DataColumn(columnName, typeof(string));
                GraphDataTable.Rows.Clear();
                DataColumnCollection columns = GraphDataTable.Columns;
                if (columns.Contains(columnName) == false)
                {
                    GraphDataTable.Columns.Add(dcFormatString);
                }

                string tableName = "GraphData";
                string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;

                using (var da = new SqlCeDataAdapter(sql, Connection()))
                {
                    da.Fill(GraphDataTable);
                }
                GraphDataTable.DefaultView.Sort = columnName + " asc";

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
            finally
            {

            }

        }

        public static void SetServerPath()
        {
            try
            {
                string versionNumber = AssemblyInfo.versionFolderNumber;
                //return Path.Combine(Properties.Settings.Default.App_PathDeploy, "Application Files", AssemblyInfo.Product + versionNumber, dataFolder, AssemblyInfo.Product + ".sdf.deploy"); //for internal server
                Uri baseUrl = new Uri(Properties.Settings.Default.App_PathDeploy);
                string relativeUrl = "Application Files/" + AssemblyInfo.Product + versionNumber + "/" + dataFolder + "/";
                Uri combined = new Uri(baseUrl, relativeUrl);
                Properties.Settings.Default.App_PathDeployData = combined.ToString();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        public static void SetUserPath()
        {
            string userFilePath = string.Empty;
            try
            {
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                {
                    string versionNumber = AssemblyInfo.versionFolderNumber;
                    string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    userFilePath = Path.Combine(localAppData, AssemblyInfo.Copyright.Replace(" ", "_"), AssemblyInfo.Product, dataFolder);
                    System.IO.Directory.CreateDirectory(userFilePath);
                    DirectoryInfo info = new DirectoryInfo(userFilePath);
                    DirectorySecurity security = info.GetAccessControl();
                    security.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.Modify, InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow));
                    security.AddAccessRule(new FileSystemAccessRule("Everyone", FileSystemRights.Modify, InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow));
                    info.SetAccessControl(security);

                }
                else
                {
                    userFilePath = System.IO.Path.Combine(AssemblyInfo.GetClickOnceLocation(), dataFolder);
                }
                Properties.Settings.Default.App_PathLocalData = userFilePath;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        public static void InsertRecord(DataTable tbl, string text)
        {
            string tableName = tbl.TableName.ToString();
            string columnName = tbl.Columns[0].ColumnName.ToString();
            string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;
            if (tbl.Select(columnName + " = '" + text.Replace("'", "''") + "'").Length == 0)
            {
                DialogResult dr = MessageBox.Show("Would you like to add '" + text + "' to the list?", "Add New Value", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                switch (dr)
                {
                    case DialogResult.Yes:
                        tbl.Rows.Add(new Object[] { text });
                        SqlCeConnection cn = new SqlCeConnection(Data.Connection());
                        SqlCeCommandBuilder scb = default(SqlCeCommandBuilder);
                        SqlCeDataAdapter sda = new SqlCeDataAdapter(sql, cn);
                        sda.TableMappings.Add("Table", tableName);
                        scb = new SqlCeCommandBuilder(sda);
                        sda.Update(tbl);

                        dynamic dcFormatString = new DataColumn(columnName, typeof(string));
                        tbl.Rows.Clear();
                        DataColumnCollection columns = tbl.Columns;
                        if (columns.Contains(columnName) == false)
                        {
                            tbl.Columns.Add(dcFormatString);
                        }

                        using (var da = new SqlCeDataAdapter(sql, Connection()))
                        {
                            da.Fill(tbl);
                        }
                        tbl.DefaultView.Sort = columnName + " asc";

                        break;

                    case DialogResult.No:
                        break;
                }
            }
        }

    }
}
