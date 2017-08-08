using System;
using System.IO;
using System.Data;
using System.Data.SqlServerCe;
using System.Linq;
using System.Security.AccessControl;

namespace ScriptHelp.Scripts
{
	class Data
	{

		const string dataFolder = "App_Data";

		/// <summary>
		/// Relative database connection string
		/// </summary>
		/// <returns>the data source of the database</returns>
		public static string Connection()
		{
			Data.SetUserPath();
			string databaseFile = "Data Source=" + Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".sdf");
			return databaseFile;
		}

		/// <summary>
		/// List of common table alias
		/// </summary>
		public static DataTable TableAliasTable = new DataTable();

		/// <summary>
		/// List of date format strings
		/// </summary>
		public static DataTable DateFormatTable = new DataTable();

		/// <summary>
		/// List of graph data
		/// </summary>
		public static DataTable GraphDataTable = new DataTable();

		/// <summary>
		/// Creates the datatable for the list of common table alias
		/// </summary>
		public static void CreateTableAliasTable()
		{
			try
			{
				string columnName = "TableName";
				dynamic dcTableName = new DataColumn(columnName, typeof(string));
				TableAliasTable.Rows.Clear();
				DataColumnCollection columns = TableAliasTable.Columns;
				if (columns.Contains(columnName) == false)
				{
					TableAliasTable.Columns.Add(dcTableName);
				}

				string tableName = "TableAlias";
				string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;

				using (var da = new SqlCeDataAdapter(sql, Connection()))
				{
					da.Fill(TableAliasTable);
				}
				TableAliasTable.DefaultView.Sort = columnName + " asc";

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);

			}

		}

		/// <summary>
		/// Creates the datatable for the date format strings
		/// </summary>
		public static void CreateDateFormatTable()
		{
			try
			{
				string columnName = "FormatString";
				dynamic dcFormatString = new DataColumn(columnName, typeof(string));
				DateFormatTable.Rows.Clear();
				DataColumnCollection columns = DateFormatTable.Columns;
				if (columns.Contains(columnName) == false)
				{
					DateFormatTable.Columns.Add(dcFormatString);
				}

				string tableName = "DateFormat";
				string sql = "SELECT * FROM " + tableName + " ORDER BY " + columnName;

				using (var da = new SqlCeDataAdapter(sql, Connection()))
				{
					da.Fill(DateFormatTable);
				}
				DateFormatTable.DefaultView.Sort = columnName + " asc";

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);

			}

		}

		/// <summary>
		/// Creates the datatable for the graph data
		/// </summary>
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

		/// <summary>
		/// 
		/// </summary>
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

		/// <summary>
		/// 
		/// </summary>
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
					//if (!Directory.Exists(userFilePath)) Directory.CreateDirectory(userFilePath);
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

	}
}
