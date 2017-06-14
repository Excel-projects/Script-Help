using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

// <summary> 
// This namespaces if for ribbon classes and methods
// </summary>
namespace ScriptHelp.Scripts
{
	/// <summary> 
	/// Class for the ribbon procedures
	/// </summary>
	[ComVisible(true)]
	public class Ribbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		/// <summary>
		/// Used to reference the ribbon object
		/// </summary>
		public static Ribbon ribbonref;

		#region | Task Panes |

		/// <summary>
		/// Settings TaskPane
		/// </summary>
		public TaskPane.Settings mySettings;

		/// <summary>
		/// Script TaskPane
		/// </summary>
		public TaskPane.Script myScript;

		/// <summary>
		/// TableData TaskPane
		/// </summary>
		public TaskPane.TableData myTableData;

		/// <summary>
		/// TableData TaskPane
		/// </summary>
		public TaskPane.GraphData myGraphData;

		/// <summary>
		/// Settings Custom Task Pane
		/// </summary>
		public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;

		/// <summary>
		/// Script Custom Task Pane
		/// </summary>
		public Microsoft.Office.Tools.CustomTaskPane myTaskPaneScript;

		/// <summary>
		/// TableData Custom Task Pane
		/// </summary>
		public Microsoft.Office.Tools.CustomTaskPane myTaskPaneTableData;

		/// <summary>
		/// TableData Custom Task Pane
		/// </summary>
		public Microsoft.Office.Tools.CustomTaskPane myTaskPaneGraphData;

		#endregion

		/// <summary>
		/// Used for values across different classes
		/// </summary>
		public static class AppVariables
		{
			/// <summary>
			/// variable used for sending the copied range to the form for export
			/// </summary>
			public static string ScriptRange { get; set; }

			/// <summary>
			/// variable used for saving the script file
			/// </summary>
			public static string FileType { get; set; }

			/// <summary>
			/// variable used for the table name used to populate a datagrid
			/// </summary>
			public static string TableName { get; set; }

			/// <summary>
			/// The first visible column name in the table
			/// </summary>
			public static string FirstColumnName { get; set; }

		}

		#region | Helpers |

		/// <summary> 
		/// The Sql Help ribbon
		/// </summary>
		public Ribbon()
		{
		}

		#region | IRibbonExtensibility Members |

		/// <summary> 
		/// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
		/// </summary>
		/// <param name="ribbonID">Represents the XML customization file </param>
		/// <returns>A method that returns a bitmap image for the control id. </returns> 
		/// <remarks></remarks>
		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("ScriptHelp.Ribbon.xml");
		}

		#endregion

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion

		#region | Ribbon Events |

		/// <summary> 
		/// loads the ribbon UI and creates a log record
		/// </summary>
		/// <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code. </param>
		/// <remarks></remarks>
		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			try
			{
				this.ribbon = ribbonUI;
				ribbonref = this;
				ThisAddIn.e_ribbon = ribbonUI;
				ErrorHandler.SetLogPath();
				ErrorHandler.CreateLogRecord();
				AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
				Data.CreateTableAliasTable();
				Data.CreateDateFormatTable();
				Data.CreateGraphDataTable();

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		/// <summary> 
		/// Assigns an image to a button on the ribbon in the xml file
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns a bitmap image for the control id. </returns> 
		public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "btnQueryTypeDqlAppend":
					case "btnQueryTypeDqlAppendLocked":
					case "btnQueryTypeDqlCreate":
					case "btnQueryTypeDqlTruncateAppend":
					case "btnQueryTypeDqlUpdate":
					case "btnQueryTypeDqlUpdateLocked":
						return Properties.Resources.QueryTypeDql;
					case "btnQueryTypeTSqlCreateTable":
					case "btnQueryTypeTSqlInsertValues":
					case "btnQueryTypeTSqlMergeValues":
					case "btnQueryTypeTSqlSelectValues":
					case "btnQueryTypeTSqlSelectUnion":
					case "btnQueryTypeTSqlUpdateValues":
						return Properties.Resources.QueryTypeTSql;
					case "btnQueryTypePlSqlSelectUnion":
					case "btnQueryTypePlSqlInsertValues":
					case "btnQueryTypePlSqlUpdateValues":
						return Properties.Resources.QueryTypePlSql;
					case "btnQueryTypeGithubTable":
						return Properties.Resources.QueryTypeGitHub;
					case "btnQueryTypeXmlValues":
						return Properties.Resources.QueryTypeXML;
					default:
						return null;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return null;
			}
		}

		/// <summary> 
		/// Assigns the enabled to controls
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns true or false if the control is enabled </returns> 
		public bool GetEnabled(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "btnCopyVisibleCells":
					case "btnCleanData":
					case "btnZeroToNull":
					case "btnFormatSqlDateColumns":
					case "btnClearInteriorColor":
					case "btnAddScriptColumn":
						return ErrorHandler.IsEnabled(false);
					default:
						return false;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return false;
			}
		}

		/// <summary> 
		/// Assigns text to a label on the ribbon from the xml file
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns a string for a label. </returns> 
		public string GetLabelText(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "tabScriptHelp":
						return AssemblyInfo.Title;
					case "txtCopyright":
						return " " + AssemblyInfo.Copyright;
					case "txtDescription":
						return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
					case "txtInstallDate":
						DateTime dteCreateDate = File.GetCreationTime(Assembly.GetExecutingAssembly().Location);
						return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");
					default:
						return string.Empty;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// Assigns the number of items for a combobox or dropdown
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns an integer of total count of items used for a combobox or dropdown </returns> 
		public int GetItemCount(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "cboDateFormat":
					case "cboDatePasteFormat":
						return Data.DateFormatTable.Rows.Count;
					case "cboTableAlias":
						return Data.TableAliasTable.Rows.Count;
					default:
						return 0;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return 0;
			}
		}

		/// <summary> 
		/// Assigns the values to a combobox or dropdown based on an index
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <param name="index">Represents the index of the combobox or dropdown value </param>
		/// <returns>A method that returns a string per index of a combobox or dropdown </returns> 
		public string GetItemLabel(Office.IRibbonControl control, int index)
		{
			try
			{
				switch (control.Id)
				{
					case "cboDateFormat":
					case "cboDatePasteFormat":
						return UpdateDateFormatComboBoxSource(index);
					case "cboTableAlias":
						return UpdateTableAliasComboBoxSource(index);

					default:
						return string.Empty;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// Assigns default values to comboboxes
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns a string for the default value of a combobox </returns> 
		public string GetText(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "cboDateFormat":
						return Properties.Settings.Default.Sheet_Column_Date_Format_Replace;
					case "cboDatePasteFormat":
						return Properties.Settings.Default.Sheet_Column_Date_Format_Find;
					case "cboTableAlias":
						return Properties.Settings.Default.Sheet_Column_Table_Alias;
					default:
						return string.Empty;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// Assigns the visiblity to controls
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns true or false if the control is visible </returns> 
		public bool GetVisible(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "cboTableAlias":
					case "cboDateFormat":
					case "cboDatePasteFormat":
					case "separator2":
					case "btnDateFormat":
					case "btnTableAlias":
					case "btnPasteFormat":
						return Properties.Settings.Default.Visible_Options;
					case "ComAddInsDialog":
						return Properties.Settings.Default.Visible_ComAddInsDialog;
					case "FormatAsTableGallery":
						return Properties.Settings.Default.Visible_FormatAsTableGallery;
					case "ViewFreezePanesGallery":
						return Properties.Settings.Default.Visible_ViewFreezePanesGallery;
					case "RemoveDuplicates":
						return Properties.Settings.Default.Visible_RemoveDuplicates;
					case "btnCopyVisibleCells":
						return Properties.Settings.Default.Visible_btnCopyVisibleCells;
					case "btnClearInteriorColor":
						return Properties.Settings.Default.Visible_btnClearInteriorColor;
					case "btnZeroToNull":
						return Properties.Settings.Default.Visible_btnZeroToNull;
					case "btnSeparateValues":
						return Properties.Settings.Default.Visible_btnSeparateValues;
					case "btnFileList":
						return Properties.Settings.Default.Visible_btnFileList;
					case "mnuScriptType":
					case "separator1":
						return Properties.Settings.Default.Visible_mnuScriptType;
					case "btnQueryTypeTSqlCreateTable":
					case "btnQueryTypeTSqlInsertValues":
					case "btnQueryTypeTSqlMergeValues":
					case "btnQueryTypeTSqlSelectValues":
					case "btnQueryTypeTSqlSelectUnion":
					case "btnQueryTypeTSqlUpdateValues":
						return Properties.Settings.Default.Visible_mnuScriptType_TSQL;
					case "btnQueryTypePlSqlSelectUnion":
					case "btnQueryTypePlSqlInsertValues":
					case "btnQueryTypePlSqlUpdateValues":
						return Properties.Settings.Default.Visible_mnuScriptType_PLSQL;
					case "btnQueryTypeDqlAppend":
					case "btnQueryTypeDqlAppendLocked":
					case "btnQueryTypeDqlCreate":
					case "btnQueryTypeDqlTruncateAppend":
					case "btnQueryTypeDqlUpdate":
					case "btnQueryTypeDqlUpdateLocked":
						return Properties.Settings.Default.Visible_mnuScriptType_DQL;
					case "btnQueryTypeGithubTable":
						return Properties.Settings.Default.Visible_mnuScriptType_Github;
					case "btnQueryTypeXmlValues":
						return Properties.Settings.Default.Visible_mnuScriptType_XML;
					default:
						return false;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return false;
			}
		}

		/// <summary>
		/// Assigns the value to an application setting
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <returns>A method that returns true or false if the control is enabled </returns> 
		public void OnAction(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "btnQueryTypeTSqlCreateTable":
					case "btnQueryTypeTSqlSelectValues":
					case "btnQueryTypeTSqlSelectUnion":
					case "btnQueryTypePlSqlSelectUnion":
					case "btnQueryTypePlSqlInsertValues":
					case "btnQueryTypeDqlUpdate":
					case "btnQueryTypeDqlCreate":
					case "btnQueryTypeDqlAppend":
					case "btnQueryTypeDqlUpdateLocked":
					case "btnQueryTypeDqlTruncateAppend":
					case "btnQueryTypeDqlAppendLocked":
					case "btnQueryTypeTSqlInsertValues":
					case "btnQueryTypeTSqlUpdateValues":
					case "btnQueryTypePlSqlUpdateValues":
					case "btnQueryTypeTSqlMergeValues":
					case "btnQueryTypeGithubTable":
					case "btnQueryTypeXmlValues":
						AddScriptColumn(control);
						break;
					case "btnDateFormat":
					case "btnTableAlias":
					case "btnPasteFormat":
						AppVariables.TableName = control.Tag;
						OpenTableDataPane();
						break;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}

		}

		/// <summary> 
		/// Return the updated value from the comboxbox
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <param name="text">Represents the text from the combobox value </param>
		public void OnChange(Office.IRibbonControl control, string text)
		{
			try
			{
				switch (control.Id)
				{
					case "cboDateFormat":
						Properties.Settings.Default.Sheet_Column_Date_Format_Replace = text;
						break;
					case "cboDatePasteFormat":
						Properties.Settings.Default.Sheet_Column_Date_Format_Find = text;
						break;
					case "cboTableAlias":
						Properties.Settings.Default.Sheet_Column_Table_Alias = text;
						break;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		#endregion

		#region | Ribbon Buttons |

		/// <summary> 
		/// Copy only the visible cells that are selected
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void CopyVisibleCells(Office.IRibbonControl control)
		{
			Excel.Range visibleRange = null;
			try
			{
				if (ErrorHandler.IsEnabled(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				visibleRange = Globals.ThisAddIn.Application.Selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
				visibleRange.Copy();
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				if (visibleRange != null)
					Marshal.ReleaseComObject(visibleRange);
			}
		}

		/// <summary> 
		/// Removes all nonprintable characters from text and returns number of cells altered
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void CleanData(Office.IRibbonControl control)
		{
			Excel.ListObject tbl = null;
			Excel.Range cell = null;
			Excel.Range usedRange = null;
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				cell = default(Excel.Range);
				string c = string.Empty;
				string cc = string.Empty;
				int cnt = 0;
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				usedRange = tbl.Range;
				int n = tbl.ListColumns.Count;
				int m = tbl.ListRows.Count;
				for (int i = 0; i <= m; i++) // by row
				{
					for (int j = 1; j <= n; j++) // by column
					{
						if (usedRange[i + 1, j].Value2 != null)
						{
							c = usedRange[i + 1, j].Value2.ToString();  // can't convert null to string
							if (Globals.ThisAddIn.Application.WorksheetFunction.IsText(c))
							{
								cc = Globals.ThisAddIn.Application.WorksheetFunction.Clean(c.Trim());
								if ((cc != c))
								{
									cell = tbl.Range.Cells[i + 1, j];
									if (Convert.ToBoolean(cell.HasFormula) == false)
									{
										cell.Value = cc;
										cell.Interior.Color = Properties.Settings.Default.Sheet_Column_Cleaned_Colour;
										cnt = cnt + 1;
									}
								}
								cell = tbl.Range.Cells[i + 1, j];
								string qt = Properties.Settings.Default.Sheet_Column_Script_Quote;
								if (cell.PrefixCharacter == qt)  // show the leading apostrophe in the cell by doubling the value.
								{
									cell.Value = qt + qt + cell.Value;
									cell.Interior.Color = Properties.Settings.Default.Sheet_Column_Cleaned_Colour;
								}
							}
						}
					}
				}
				MessageBox.Show("The number of cells cleaned: " + cnt.ToString(), "Cleaning has finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (cell != null)
					Marshal.ReleaseComObject(cell);
				if (usedRange != null)
					Marshal.ReleaseComObject(usedRange);
			}
		}

		/// <summary> 
		/// Finds dates columns with SSMS crap format and alters to use standard SQL date format
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void FormatSqlDateColumns(Office.IRibbonControl control)
		{
			Excel.ListObject tbl = null;
			Excel.Range cell = null;
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				cell = default(Excel.Range);
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					cell = FirstNotNullCellInColumn(col.DataBodyRange);
					if (((cell != null)))
					{
						if (cell.NumberFormat.ToString() == Properties.Settings.Default.Sheet_Column_Date_Format_Find | ErrorHandler.IsDate(cell.Value))
						{
							col.DataBodyRange.NumberFormat = Properties.Settings.Default.Sheet_Column_Date_Format_Replace;
							col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter;
						}
					}
				}
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (cell != null)
					Marshal.ReleaseComObject(cell);
			}
		}

		/// <summary> 
		/// Remove interior cell color format
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void ClearInteriorColor(Office.IRibbonControl control)
		{
			Excel.ListObject tbl = null;
			Excel.Range rng = null;
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				tbl.DataBodyRange.Interior.ColorIndex = Excel.Constants.xlNone;
				tbl.DataBodyRange.Font.ColorIndex = Excel.Constants.xlAutomatic;
				rng = tbl.Range;
				for (int i = 1; i <= rng.Columns.Count; i++)
				{
					if (rng.Columns.EntireColumn[i].Hidden == false)
					{
						((Excel.Range)rng.Cells[1, i]).Interior.ColorIndex = Excel.Constants.xlNone;
						((Excel.Range)rng.Cells[1, i]).HorizontalAlignment = Excel.Constants.xlCenter;
						((Excel.Range)rng.Cells[1, i]).VerticalAlignment = Excel.Constants.xlCenter;
					}
				}
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (rng != null)
					Marshal.ReleaseComObject(rng);
			}
		}

		/// <summary> 
		/// Change zero string cell values to string "NULL"
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void ZeroStringToNull(Office.IRibbonControl control)
		{
			Excel.ListObject tbl = null;
			Excel.Range cell = null;
			Excel.Range usedRange = null;
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				cell = default(Excel.Range);
				int cnt = 0;
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				usedRange = tbl.Range;
				int n = tbl.ListColumns.Count;
				int m = tbl.ListRows.Count;
				for (int i = 0; i <= m; i++) // by row
				{
					for (int j = 1; j <= n; j++) // by column
					{
						if (usedRange[i + 1, j].Value2 == null)
						{
							cell = tbl.Range.Cells[i + 1, j];
							cell.Value = Properties.Settings.Default.Sheet_Column_Script_Null;
							cell.Interior.Color = Properties.Settings.Default.Sheet_Column_Cleaned_Colour;
							cnt = cnt + 1;
						}
					}
				}
				MessageBox.Show("The number of cells converted to " + Properties.Settings.Default.Sheet_Column_Script_Null + ": " + cnt.ToString(), "Converting has finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (cell != null)
					Marshal.ReleaseComObject(cell);
				if (usedRange != null)
					Marshal.ReleaseComObject(usedRange);
			}
		}

		/// <summary>
		/// Add a row per delimited string value from a column
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void SeparateValues(Office.IRibbonControl control)
		{
			Excel.ListObject tbl = null;
			Excel.Range cell = null;
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				cell = Globals.ThisAddIn.Application.ActiveCell;
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				int m = tbl.ListRows.Count;
				int a = m;
				int columnIndex = cell.Column;

				for (int i = 1; i <= m + 1; i++) // by row
				{
					string cellValue = tbl.Range.Cells[i, columnIndex].Value2;
					if (string.IsNullOrEmpty(cellValue) == false)
					{
						string[] metadata = cellValue.Split(Properties.Settings.Default.Sheet_Column_Separate_Values_Delimiter);
						int countValues = metadata.Length - 1;
						if (countValues > 0)
						{
							//if the column value has multiple values then create a new row per value
							for (int j = 1; j <= countValues; j++) // by value 
							{
								tbl.ListRows.Add(i);
								tbl.Range.Rows[i + 1].Value = tbl.Range.Rows[i].Value;
								tbl.Range.Cells[i + 1, columnIndex].Value2 = metadata[j - 1].Trim();  // get the next value in the string
							}
							tbl.Range.Cells[i, columnIndex].Value2 = metadata[countValues].Trim(); // reset the first row value
							m += countValues; //reset the total rows
							i += countValues; //reset the current row
						}
					}

				}
				MessageBox.Show("The number of row(s) added is " + (m - a).ToString(), "Finished Separating Values", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (cell != null)
					Marshal.ReleaseComObject(cell);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void AddScriptColumn(Office.IRibbonControl control)
		{
			try
			{
				if (ErrorHandler.IsAvailable(true) == false)
				{
					return;
				}
				ErrorHandler.CreateLogRecord();
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				switch (control.Id)
				{
					case "btnQueryTypeTSqlCreateTable":
						AddFormulaTSqlCreateTable();
						break;
					case "btnQueryTypeTSqlInsertValues":
						AddFormulaTSqlInsertValues();
						break;
					case "btnQueryTypeTSqlMergeValues":
						AddFormulaTSqlMergeValues();
						break;
					case "btnQueryTypeTSqlSelectValues":
						AddFormulaTSqlSelectValues();
						break;
					case "btnQueryTypeTSqlSelectUnion":
						AddFormulaTSqlSelectUnion();
						break;
					case "btnQueryTypeTSqlUpdateValues":
						AddFormulaTSqlUpdateValues();
						break;
					case "btnQueryTypePlSqlInsertValues":
						AddFormulaPlSqlInsertValues();
						break;
					case "btnQueryTypePlSqlSelectUnion":
						AddFormulaPlSqlSelectUnion();
						break;
					case "btnQueryTypePlSqlUpdateValues":
						AddFormulaPlSqlUpdateValues();
						break;
					case "btnQueryTypeDqlAppend":
						AddFormulaDqlAppend();
						break;
					case "btnQueryTypeDqlAppendLocked":
						AddFormulaDqlAppendLocked();
						break;
					case "btnQueryTypeDqlCreate":
						AddFormulaDqlCreate();
						break;
					case "btnQueryTypeDqlTruncateAppend":
						AddFormulaDqlTruncateAppend();
						break;
					case "btnQueryTypeDqlUpdate":
						AddFormulaDqlUpdate();
						break;
					case "btnQueryTypeDqlUpdateLocked":
						AddFormulaDqlUpdateLocked();
						break;
					case "btnQueryTypeGithubTable":
						AddFormulaGithubTable();
						break;
					case "btnQueryTypeXmlValues":
						AddFormulaXmlValues();
						break;
				}
			}

			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
			}
		}

		/// <summary> 
		/// Opens the graph data pane
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void OpenGraphData(Office.IRibbonControl control)
		{
			try
			{
				if (myTaskPaneGraphData != null)
				{
					if (myTaskPaneGraphData.Visible == true)
					{
						myTaskPaneGraphData.Visible = false;
					}
					else
					{
						myTaskPaneGraphData.Visible = true;
					}
				}
				else
				{
					myGraphData = new TaskPane.GraphData();
					myTaskPaneGraphData = Globals.ThisAddIn.CustomTaskPanes.Add(myGraphData, "Graph Data for " + Scripts.AssemblyInfo.Title);
					myTaskPaneGraphData.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
					myTaskPaneGraphData.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
					myTaskPaneGraphData.Width = 300;
					myTaskPaneGraphData.Visible = true;
				}

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}

		}

		/// <summary> 
		/// Opens the settings form
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void OpenSettingsForm(Office.IRibbonControl control)
		{
			try
			{
				if (myTaskPaneSettings != null)
				{
					if (myTaskPaneSettings.Visible == true)
					{
						myTaskPaneSettings.Visible = false;
					}
					else
					{
						myTaskPaneSettings.Visible = true;
					}
				}
				else
				{
					mySettings = new TaskPane.Settings();
					myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + Scripts.AssemblyInfo.Title);
					myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
					myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
					myTaskPaneSettings.Width = 675;
					myTaskPaneSettings.Visible = true;
				}

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		/// <summary> 
		/// Creates a recursive file listing based on the users selected directory
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void CreateBatchFileList(Office.IRibbonControl control)
		{
			string filePath = Properties.Settings.Default.App_PathFileListing;
			try
			{
				ErrorHandler.CreateLogRecord();
				DialogResult msgDialogResult = DialogResult.None;
				FolderBrowserDialog dlg = new FolderBrowserDialog();
				if (Properties.Settings.Default.Option_FileListingPathSelect == true)
				{
					dlg.RootFolder = Environment.SpecialFolder.MyComputer;
					dlg.SelectedPath = filePath;
					msgDialogResult = dlg.ShowDialog();
					filePath = dlg.SelectedPath;
				}
				if (msgDialogResult == DialogResult.OK | Properties.Settings.Default.Option_FileListingPathSelect == false)
				{
					filePath += @"\";
					string scriptCommands = string.Empty;
					string currentDate = DateTime.Now.ToString("dd.MMM.yyyy_hh.mm.tt");
					string batchFileName = filePath + "FileListing_" + currentDate + "_" + Environment.UserName + ".bat";
					scriptCommands = "echo off" + Environment.NewLine;
					scriptCommands += "cd %1" + Environment.NewLine;
					scriptCommands += @"dir """ + filePath + @""" /s /a-h /b /-p /o:gen >""" + filePath + "FileListing_" + currentDate + "_" + Environment.UserName + @".csv""" + Environment.NewLine;
					scriptCommands += @"""" + filePath + "FileListing_" + currentDate + "_" + Environment.UserName + @".csv""" + Environment.NewLine;
					scriptCommands += "cd .. " + Environment.NewLine;
					scriptCommands += "echo on" + Environment.NewLine;
					System.IO.File.WriteAllText(batchFileName, scriptCommands);
					AssemblyInfo.OpenFile(batchFileName);
				}
			}
			catch (System.UnauthorizedAccessException)
			{
				MessageBox.Show("You don't have access to this folder, bro!" + Environment.NewLine + Environment.NewLine + filePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		/// <summary> 
		/// Opens an as built help file
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void OpenHelpAsBuiltFile(Office.IRibbonControl control)
		{
			ErrorHandler.CreateLogRecord();
			string clickOnceLocation = AssemblyInfo.GetClickOnceLocation();
			AssemblyInfo.OpenFile(Path.Combine(clickOnceLocation, @"Documentation\\As Built.docx"));

		}

		/// <summary> 
		/// Opens a api help file
		/// </summary>
		/// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
		/// <remarks></remarks>
		public void OpenHelpApiFile(Office.IRibbonControl control)
		{
			ErrorHandler.CreateLogRecord();
			string clickOnceLocation = AssemblyInfo.GetClickOnceLocation();
			AssemblyInfo.OpenFile(Path.Combine(clickOnceLocation, @"Documentation\\Api Help.chm"));
		}

		#endregion

		#region | Subroutines |

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlAppend()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					switch (col.Name.IndexOfAny(new char[] { '[', ']', '"' }))
					{
						case -1:
							break;
						default:
							MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							valuePrefix = "\"" + col.Name + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								valuePrefix = "\"APPEND " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				formula = "=\"UPDATE " + tableAlias + " objects \" & " + formula + whereCheck + " & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlAppendLocked()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;
				string whereClause = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					switch (col.Name.IndexOfAny(new char[] { '[', ']', '"' }))
					{
						case -1:
							break;
						default:
							MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							valuePrefix = "\"" + col.Name + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								valuePrefix = "\"APPEND " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							if (afterWhere == true)
							{
								whereClause += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";

							}
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							if (afterWhere == true)
							{
								whereClause += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
							}
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				string formulaUnlock = "\"UPDATE " + tableAlias + "(all) objects SET r_immutable_flag = 0 \" & " + whereClause + " & CHAR(10) & \"GO \"  & CHAR(10) & ";
				string formulaLock = "\"UPDATE " + tableAlias + "(all) objects SET r_immutable_flag = 1 \" & " + whereClause + " & CHAR(10) & \"GO \"";
				formula = "=" + formulaUnlock + "\"UPDATE " + tableAlias + "(all) objects \" & " + formula + whereCheck + " & CHAR(10) & \"GO \" & CHAR(10) & " + formulaLock;
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlCreate()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & CHAR(10) & \",\" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						string valuePrefix = string.Empty;
						valuePrefix = "\" SET " + col.Name + " = \" & ";
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				formula = "=\"CREATE " + tableAlias + " objects \" & CHAR(10) & " + formula + " & \"; \" & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlTruncateAppend()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					switch (col.Name.IndexOfAny(new char[] { '[', ']', '"' }))
					{
						case -1:
							break;
						default:
							MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							valuePrefix = "\"" + col.Name + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								valuePrefix = "\"TRUNCATE " + col.Name + ", APPEND " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				formula = "=\"UPDATE " + tableAlias + " objects \" & " + formula + whereCheck + " & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlUpdate()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							valuePrefix = "\"" + col.Name + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								valuePrefix = "\"SET " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				formula = "=\"UPDATE " + tableAlias + " objects \" & " + formula + whereCheck + " & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaDqlUpdateLocked()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;
				string whereClause = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					switch (col.Name.IndexOfAny(new char[] { '[', ']', '"' }))
					{
						case -1:
							break;
						default:
							MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							valuePrefix = "\"" + col.Name + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								valuePrefix = "\"SET " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							if (afterWhere == true)
							{
								whereClause += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";

							}
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							if (afterWhere == true)
							{
								whereClause += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
							}
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}
				//replace NULL values with DQL format
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "', '" + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + "'\", \"'nulldate'\")";
				formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Sheet_Column_Script_Null + "'\", \"nullstring\")";
				formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Sheet_Column_Script_Null + "\", \"nullint\")";

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				string formulaUnlock = "\"UPDATE " + tableAlias + "(all) objects SET r_immutable_flag = 0 \" & " + whereClause + " & CHAR(10) & \"GO \"  & CHAR(10) & ";
				string formulaLock = "\"UPDATE " + tableAlias + "(all) objects SET r_immutable_flag = 1 \" & " + whereClause + " & CHAR(10) & \"GO \"";
				formula = "=" + formulaUnlock + "\"UPDATE " + tableAlias + "(all) objects \" & " + formula + whereCheck + " & CHAR(10) & \"GO \" & CHAR(10) & " + formulaLock;
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "DQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}
		
		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaPlSqlInsertValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string insertPrefix = "INSERT INTO " + tableAlias + " (" + ConcatenateColumnNames(tbl.Range) + ") VALUES(";
				formula = "=\"" + insertPrefix + "\" & " + formula + " & \");\"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaPlSqlSelectUnion()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						string dqt = "\"\"";
						string valuePlSuffix = "& \" AS " + dqt + col.Name + dqt + " \"";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valuePlSuffix;
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
				formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \"\", \"UNION \") & " + "\"SELECT \" & " + formula + " & \" FROM DUAL\"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaPlSqlUpdateValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							string firstWhereColumn = col.Name;
							firstWhereColumn = System.Text.RegularExpressions.Regex.Replace(firstWhereColumn, "where ", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
							firstWhereColumn = firstWhereColumn.Trim();
							valuePrefix = "\"WHERE " + firstWhereColumn + " = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND " + col.Name + " = \" & ";
							}
							else
							{
								string useComma = string.Empty;
								if (columnCount != 1)
								{
									useComma = ",";
								}
								valuePrefix = "\"" + useComma + " " + col.Name + " = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				formula = "=\"UPDATE " + tableAlias + " SET \" & " + formula + whereCheck + " & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlCreateTable()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string insertPrefix = "INSERT INTO " + tableAlias + " (" + ConcatenateColumnNames(tbl.Range) + ") VALUES(";
				formula = "=\"" + insertPrefix + "\" & " + formula + " & \");\"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				string createTable = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'" + tableAlias + "') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE " + tableAlias + Environment.NewLine + "; " + Environment.NewLine + "CREATE TABLE " + tableAlias + " (" + tableAlias + "_ID [int] PRIMARY KEY IDENTITY(1,1) NOT NULL, " + ConcatenateColumnNames(tbl.Range, "", Environment.NewLine + "[", "] [varchar](max) NULL") + Environment.NewLine + ");" + Environment.NewLine;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = createTable + (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlInsertValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string insertPrefix = "INSERT INTO " + tableAlias + " (" + ConcatenateColumnNames(tbl.Range) + ") VALUES(";
				formula = "=\"" + insertPrefix + "\" & " + formula + " & \");\"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlMergeValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAliasTemp = tableAlias + "_temp";
				string sqlColName = string.Empty;

				sqlColName = "SELECT " + tableAliasTemp + ".*" + " FROM (VALUES";

				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				// Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					//if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
				formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				tbl.ShowTotals = true;
				string totalsColumnValue = ") " + tableAliasTemp + " (" + ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
				tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
					AppVariables.ScriptRange = "SET XACT_ABORT ON" + Environment.NewLine + "BEGIN TRANSACTION;" + Environment.NewLine + Environment.NewLine + ";WITH " + Environment.NewLine + tableAliasTemp + Environment.NewLine + "AS " + Environment.NewLine + "(" + Environment.NewLine + AppVariables.ScriptRange + ") " + Environment.NewLine + "MERGE " + tableAlias + " AS T" + Environment.NewLine + "USING " + tableAliasTemp + " AS S" + Environment.NewLine + "ON " + ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "WHEN NOT MATCHED BY TARGET" + Environment.NewLine + "THEN INSERT" + Environment.NewLine + "(" + Environment.NewLine + ConcatenateColumnNames(tbl.Range, "", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "VALUES" + Environment.NewLine + "(" + Environment.NewLine + ConcatenateColumnNames(tbl.Range, "S", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "WHEN MATCHED" + Environment.NewLine + "THEN UPDATE SET" + Environment.NewLine + ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "--WHEN NOT MATCHED BY SOURCE AND 'ADD WHERE CLAUSE HERE'" + Environment.NewLine + "--THEN DELETE" + Environment.NewLine + "OUTPUT $action, inserted.*, deleted.*;" + Environment.NewLine + Environment.NewLine + "ROLLBACK TRANSACTION;" + Environment.NewLine + "--COMMIT TRANSACTION;" + Environment.NewLine + "GO";
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}
			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlSelectUnion()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				sqlColName = Properties.Settings.Default.Sheet_Column_Name;

				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				// Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					//if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						string valueTsuffix = "& \" AS [" + col.Name + "] \"";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valueTsuffix;
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
				formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \"\", \"UNION \") & " + "\"SELECT \" & " + formula + " & \" \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}
			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlSelectValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;

				sqlColName = "SELECT " + tableAlias + ".*" + " FROM (VALUES";

				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];

				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				// Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					//if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \", \" & ";
						}
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
						formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
					}
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
				formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				tbl.ShowTotals = true;
				string totalsColumnValue = ") " + tableAlias + " (" + ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
				tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}
			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaTSqlUpdateValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;
				bool afterWhere = false;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string columnName = col.Name;
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);
						if (String.Compare(firstPart.Trim(), "WHERE", true) == 0)
						{
							string firstWhereColumn = col.Name;
							firstWhereColumn = System.Text.RegularExpressions.Regex.Replace(firstWhereColumn, "where ", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
							firstWhereColumn = firstWhereColumn.Trim();
							valuePrefix = "\"WHERE [" + firstWhereColumn + "] = \" & ";
							afterWhere = true;
						}
						else
						{
							if (afterWhere == true)
							{
								valuePrefix = "\"AND [" + col.Name + "] = \" & ";
							}
							else
							{
								string useComma = string.Empty;
								if (columnCount != 1)
								{
									useComma = ",";
								}
								valuePrefix = "\"" + useComma + " [" + col.Name + "] = \" & ";
							}
						}
						if (GetSqlDataType(col) == Properties.Settings.Default.Script_Type_Date)
						{
							formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Sheet_Column_Date_Format_Replace + qt + ")\"";
						}
						else
						{
							formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
						}
					}
				}

				string whereCheck = string.Empty;
				if (afterWhere == false)
				{
					whereCheck = " & \" WHERE \" ";
				}
				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				formula = "=\"UPDATE " + tableAlias + " SET\" & " + formula + whereCheck + " & CHAR(10) & \"GO \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "SQL";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaGithubTable()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;

				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				sqlCol = tbl.ListColumns.Add();
				sqlCol.Name = lastColumnName;
				int lastColumnIndex = tbl.Range.Columns.Count;

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \"|\" & ";
						}
						formula += GetColumnFormat(col).ToString();
					}
				}
				formula = "=\"" + "|" + "\" & " + formula + " & \"|\"";
				string sqlColName = ConcatenateColumnNames(tbl.Range, string.Empty, "|") + "|";
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.Copy();
					AppVariables.FileType = "TXT";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}
			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}

		}

		/// <summary> 
		/// Add a formula at the end of the table to use as a script
		/// </summary>
		/// <remarks></remarks>
		public void AddFormulaXmlValues()
		{
			Excel.ListObject tbl = null;
			Excel.ListColumn sqlCol = null;
			try
			{
				ErrorHandler.CreateLogRecord();
				string lastColumnName = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string tableAlias = Properties.Settings.Default.Sheet_Column_Table_Alias;
				string sqlColName = string.Empty;
				int columnCount = 0;

				sqlColName = Properties.Settings.Default.Sheet_Column_Name;
				tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
				int lastColumnIndex = tbl.Range.Columns.Count;
				sqlCol = tbl.ListColumns[lastColumnIndex];
				if (sqlCol.Name == sqlColName)
				{
					lastColumnName = sqlCol.Name;
				}
				else
				{
					sqlCol = tbl.ListColumns.Add();
					sqlCol.Name = lastColumnName;
					lastColumnIndex = tbl.Range.Columns.Count;
				}

				sqlCol.DataBodyRange.NumberFormat = "General";
				string formula = string.Empty;
				string qt = string.Empty;

				foreach (Excel.ListColumn col in tbl.ListColumns)
				{
					if (col.Name.IndexOfAny(new char[] { '[', ']', '"' }) != -1)
					{
						MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + "\" " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}
					if (col.Name == lastColumnName | col.Range.EntireColumn.Hidden)
					{
						//DO NOTHING - because the column is hidden or the last column with the sql script
					}
					else
					{
						if (!string.IsNullOrEmpty(formula))
						{
							formula = formula + " & \" \" & ";
						}
						if (columnCount == 0)
						{
							AppVariables.FirstColumnName = col.Name;
						}
						columnCount += 1;
						qt = ApplyTextQuotes(col);
						string colRef = GetColumnFormat(col).ToString();
						colRef = colRef.Replace("'", "''");
						colRef = colRef.Replace("#", "'#");
						colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

						string valuePrefix = string.Empty;
						string valueSuffix = string.Empty;
						string columnName = col.Name.ToLower();
						string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);

						valuePrefix = " CHAR(10) & \"<" + columnName + ">\" & ";
						valueSuffix = " & \"</" + columnName + ">\" ";
						formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valueSuffix;
					}
				}

				string nullValue = Properties.Settings.Default.Sheet_Column_Script_Null;
				formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
				formula = "=\"<row> \" & " + formula + " & CHAR(10) & \"</row> \"";
				tbl.ShowTotals = false;
				lastColumnName = sqlColName;  // maximum header characters are 255
				tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
				try
				{
					sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
					sqlCol.Range.Columns.AutoFit();
					sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
					sqlCol.Range.WrapText = true;
					sqlCol.DataBodyRange.Copy();
					AppVariables.FileType = "XML";
					AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
					AppVariables.ScriptRange = AppVariables.ScriptRange.Replace(@"""", String.Empty);
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
				}
				finally
				{
					OpenScriptPane();
				}

			}
			catch (System.OutOfMemoryException)
			{
				MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
			finally
			{
				Cursor.Current = System.Windows.Forms.Cursors.Arrow;
				if (tbl != null)
					Marshal.ReleaseComObject(tbl);
				if (sqlCol != null)
					Marshal.ReleaseComObject(sqlCol);
			}
		}

		/// <summary> 
		/// Some columns in SQL will need quoting and others will not
		/// </summary>
		/// <param name="col">Represents the list column </param>
		/// <returns>A method that returns a string of a quote based on application settings for this value. </returns> 
		/// <remarks></remarks>
		private string ApplyTextQuotes(Excel.ListColumn col)
		{
			try
			{
				if ((GetSqlDataType(col) != Properties.Settings.Default.Script_Type_Numeric))
				{
					return Properties.Settings.Default.Sheet_Column_Script_Quote;
				}
				else
				{
					return string.Empty;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary>
		/// Return the list of column names in formatted string for SQL
		/// </summary>
		/// <param name="rng">Represents the Excel Range value</param>
		/// <param name="tableAliasName">Table alias used to prefix column names</param>
		/// <param name="prefixChar">The prefix character for the column name e.g. "["</param>
		/// <param name="suffixChar">The suffix character for the column name e.g. "]"</param>
		/// <returns>A method that returns a string of the column names</returns>
		public string ConcatenateColumnNames(Excel.Range rng, string tableAliasName = "", string prefixChar = "", string suffixChar = "")
		{
			try
			{
				string columnNames = string.Empty;
				string selectionChar = string.Empty;
				if (prefixChar != "|")
				{
					selectionChar = ", ";
				}
				if (tableAliasName != string.Empty)
				{
					tableAliasName = tableAliasName + ".";
				}
				for (int i = 1; i <= rng.Columns.Count - 1; i++)
				{
					if (rng.Columns.EntireColumn[i].Hidden == false)
					{
						columnNames = columnNames + selectionChar + tableAliasName + prefixChar + ((Excel.Range)rng.Cells[1, i]).Value2 + suffixChar;
					}
				}
				if (columnNames.Substring(0, selectionChar.Length).Contains(selectionChar) && selectionChar.Length > 0)
				{
					columnNames = columnNames.Substring(2, columnNames.Length - 2);
				}
				return columnNames;
			}
			catch (Exception)
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Return the list of column names in formatted string for SQL
		/// </summary>
		/// <param name="rng">Represents the Excel Range value</param>
		/// <param name="tableAliasNameTarget">Table alias used to prefix column names</param>
		/// <param name="tableAliasNameSource">Table alias used to prefix column names</param>
		/// <returns>A method that returns a string of the column names</returns>
		public string ConcatenateColumnNamesJoin(Excel.Range rng, string tableAliasNameTarget, string tableAliasNameSource)
		{
			try
			{
				string columnNames = string.Empty;
				for (int i = 1; i <= rng.Columns.Count - 1; i++)
				{
					if (rng.Columns.EntireColumn[i].Hidden == false)
					{
						columnNames = columnNames + ", " + tableAliasNameTarget + ".[" + ((Excel.Range)rng.Cells[1, i]).Value2 + "] = " + tableAliasNameSource + ".[" + ((Excel.Range)rng.Cells[1, i]).Value2 + "]" + Environment.NewLine;
					}
				}
				columnNames = columnNames.Substring(2, columnNames.Length - 2);
				return columnNames;
			}
			catch (Exception)
			{
				return string.Empty;
			}
		}

		/// <summary> 
		/// Get the first not null cell
		/// </summary>   
		/// <param name="rng">Represents the cell range </param>
		/// <returns>A method that returns a range </returns> 
		/// <remarks> 
		/// TODO: find a way to do this without looping.
		/// NOTE: SpecialCells is unreliable when called from VBA UDFs (Odd ??!)               
		///</remarks> 
		private Excel.Range FirstNotNullCellInColumn(Excel.Range rng)
		{
			try
			{
				if ((rng == null))
				{
					return null;
				}

				foreach (Excel.Range row in rng.Rows)
				{
					Excel.Range cell = row.Cells[1, 1];
					if ((cell.Value != null))
					{
						string cellValue = cell.Value2.ToString();
						if (String.Compare(cellValue, Properties.Settings.Default.Sheet_Column_Script_Null, true) != 0)
						{
							return cell;
						}
					}
				}
				return null;
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return null;
			}
		}

		/// <summary> 
		/// Generate a formula reference with text formatting
		/// </summary>
		/// <param name="col">Represents the list column </param>
		/// <param name="fmt">Represents the formatting string </param>
		/// <returns>A method that returns a string of a formula </returns> 
		/// <remarks></remarks>
		private string FormatCellText(Excel.ListColumn col, string fmt)
		{
			string functionReturnValue = null;
			try
			{
				functionReturnValue = "[" + col.Name + "]";
				if ((string.IsNullOrEmpty(fmt)))
				{
					return functionReturnValue;
				}
				return "TEXT(" + functionReturnValue + ",\"" + fmt + "\")";
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// To return a data format for a column
		/// </summary>
		/// <param name="col">Represents the list column </param>
		/// <returns>A method that returns a string </returns> 
		/// <remarks></remarks>
		private string GetColumnFormat(Excel.ListColumn col)
		{
			try
			{
				string fmt = string.Empty;
				string nFmt = string.Empty;

				switch (GetSqlDataType(col))
				{
					case 2:
						fmt = Properties.Settings.Default.Sheet_Column_Date_Format_Replace;
						return FormatCellText(col, fmt);
					case 1:
						// we will use the column formatting if some is applied
						if ((col.DataBodyRange.NumberFormat != null))
						{
							nFmt = col.DataBodyRange.NumberFormat.ToString();
							if (!(nFmt == "General"))
							{
								fmt = nFmt;
							}
						}
						return FormatCellText(col, fmt);
				}
				return FormatCellText(col, fmt);
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// Determine the likely SQL type of the column
		/// </summary>
		/// <param name="col">Represents the list column </param>
		/// <returns>A method that returns an integer of the column data type </returns> 
		/// <remarks></remarks>
		private int GetSqlDataType(Excel.ListColumn col)
		{
			try
			{
				// default to text
				double numCnt = 0;
				double notNullCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(col.DataBodyRange, "<>" + Properties.Settings.Default.Sheet_Column_Script_Null);

				// If all values are nulls then assume text
				if ((notNullCnt == 0))
				{
					return Properties.Settings.Default.Script_Type_Text;
				}

				numCnt = Globals.ThisAddIn.Application.WorksheetFunction.Count(col.DataBodyRange);
				// if no numbers then assume text
				if ((numCnt == 0))
				{
					return Properties.Settings.Default.Script_Type_Text;
				}

				// if a mix of numbers and not numbers then assume text
				if ((numCnt != notNullCnt))
				{
					return Properties.Settings.Default.Script_Type_Text;
				}

				//Excel changes the case of date formats on custom cell format types
				bool result = Properties.Settings.Default.Sheet_Column_Date_Format_Replace.Equals(col.DataBodyRange.NumberFormat.ToString(), StringComparison.OrdinalIgnoreCase);
				// NOTE: next test relies consistent formatting on numerics in a column
				// so we only have to test the first cell
				if (ErrorHandler.IsDate(FirstNotNullCellInColumn(col.DataBodyRange)) | result == true)
				{
					return Properties.Settings.Default.Script_Type_Date;
				}
				else
				{
					return Properties.Settings.Default.Script_Type_Numeric;
				}
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return Properties.Settings.Default.Script_Type_Text;
			}
		}

		/// <summary> 
		/// Return the count of items in a delimited list
		/// </summary>
		/// <param name="valueList">Represents the list of values in a string </param>
		/// <param name="delimiter">Represents the list delimiter </param>
		/// <returns>the number of values in a delimited string</returns>
		public int GetListItemCount(string valueList, string delimiter)
		{
			try
			{
				string[] comboList = valueList.Split((delimiter).ToCharArray());
				return comboList.GetUpperBound(0) + 1;

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return 0;
			}
		}

		/// <summary> 
		/// Update the source of the combobox to a delimited string
		/// </summary>
		/// <param name="itemIndex">Represents the index of the list value </param>
		/// <returns>the label value for the combobox index</returns>
		public string UpdateDateFormatComboBoxSource(int itemIndex)
		{
			try
			{
				return Data.DateFormatTable.Rows[itemIndex]["FormatString"].ToString();
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}
		}

		/// <summary> 
		/// Update the source of the combobox to a delimited string
		/// </summary>
		/// <param name="itemIndex">Represents the index of the list value </param>
		/// <returns>the label value for the combobox index</returns>
		public string UpdateTableAliasComboBoxSource(int itemIndex)
		{
			try
			{
				return Data.TableAliasTable.Rows[itemIndex]["TableName"].ToString();
			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
				return string.Empty;
			}

		}

		/// <summary>
		/// Used to update/reset the ribbon values
		/// </summary>
		public void InvalidateRibbon()
		{
			ribbon.Invalidate();
		}

		/// <summary>
		/// Opens the task pane for the script
		/// </summary>
		public void OpenScriptPane()
		{
			try
			{
				if (myTaskPaneScript != null)
				{
					myTaskPaneScript.Dispose();
					myScript.Dispose();
				}
				myScript = new TaskPane.Script();
				myTaskPaneScript = Globals.ThisAddIn.CustomTaskPanes.Add(myScript, "Script for " + Scripts.AssemblyInfo.Title);
				myTaskPaneScript.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
				myTaskPaneScript.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				myTaskPaneScript.Width = 675;
				myTaskPaneScript.Visible = true;

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		/// <summary>
		/// Opens the task pane for the a list
		/// </summary>
		public void OpenTableDataPane()
		{
			try
			{
				if (myTaskPaneTableData != null)
				{
					myTaskPaneTableData.Dispose();
					myTableData.Dispose();
				}
				myTableData = new TaskPane.TableData();
				myTaskPaneTableData = Globals.ThisAddIn.CustomTaskPanes.Add(myTableData, "List of " + Ribbon.AppVariables.TableName + " for " + Scripts.AssemblyInfo.Title);
				myTaskPaneTableData.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
				myTaskPaneTableData.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				myTaskPaneTableData.Width = 300;
				myTaskPaneTableData.Visible = true;

			}
			catch (Exception ex)
			{
				ErrorHandler.DisplayMessage(ex);
			}
		}

		#endregion

	}
}