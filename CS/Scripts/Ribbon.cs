using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
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

        #region | Task Panes |

        /// <summary>
        /// Settings TaskPane
        /// </summary>
        public TaskPane.Settings mySettings;

        /// <summary>
        /// Script TaskPane
        /// </summary>
        public static TaskPane.Script myScript;

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
        public static Microsoft.Office.Tools.CustomTaskPane myTaskPaneScript;

        /// <summary>
        /// TableData Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneTableData;

        /// <summary>
        /// TableData Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneGraphData;

        #endregion

        #region | Ribbon Events |

        /// <summary> 
        /// The ribbon
        /// </summary>
        public Ribbon()
        {
        }

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

        /// <summary>
        /// Called by the GetCustomUI method to obtain the contents of the Ribbon XML file.
        /// </summary>
        /// <param name="resourceName">name of  the XML file</param>
        /// <returns>the contents of the XML file</returns>
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
                AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
                AssemblyInfo.SetAssemblyFolderVersion();
                Data.SetServerPath();
                Data.SetUserPath();
                ErrorHandler.SetLogPath();
                ErrorHandler.CreateLogRecord();

                string destFilePath = Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".sdf");
                if (!(File.Exists(destFilePath)))
                {
                    using (var client = new System.Net.WebClient())
                    {
                        client.DownloadFile(Properties.Settings.Default.App_PathDeployData + AssemblyInfo.Product + ".sdf.deploy", Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".sdf"));
                    }

                }

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
                    case "btnScriptTypeDqlAppend":
                    case "btnScriptTypeDqlAppendLocked":
                    case "btnScriptTypeDqlCreate":
                    case "btnScriptTypeDqlTruncateAppend":
                    case "btnScriptTypeDqlUpdate":
                    case "btnScriptTypeDqlUpdateLocked":
                        return Properties.Resources.ScriptTypeDql;
                    case "btnScriptTypeTSqlCreateTable":
                    case "btnScriptTypeTSqlInsertValues":
                    case "btnScriptTypeTSqlMergeValues":
                    case "btnScriptTypeTSqlSelectValues":
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btnScriptTypeTSqlUpdateValues":
                        return Properties.Resources.ScriptTypeTSql;
                    case "btnScriptTypePlSqlCreateTable":
                    case "btnScriptTypePlSqlInsertValues":
                    case "btnScriptTypePlSqlMergeValues":
                    case "btnScriptTypePlSqlSelectValues":
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btnScriptTypePlSqlUpdateValues":
                        return Properties.Resources.ScriptTypePlSql;
                    case "btnScriptTypeGithubTable":
                        return Properties.Resources.ScriptTypeGitHub;
                    case "btnScriptTypeHtmlTable":
                    case "btnScriptTypeXmlValues":
                        return Properties.Resources.ScriptTypeXML;
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
                    case "btnFormatDateColumns":
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
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
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
                    case "cboDateFormatReplace":
                    case "cboDateFormatFind":
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
                    case "cboDateFormatReplace":
                    case "cboDateFormatFind":
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
                    case "cboDateFormatReplace":
                        return Properties.Settings.Default.Table_ColumnDateFormatReplace;
                    case "cboDateFormatFind":
                        return Properties.Settings.Default.Table_ColumnDateFormatFind;
                    case "cboTableAlias":
                        return Properties.Settings.Default.Table_ColumnTableAlias;
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
                    case "grpClipboard":
                        return Properties.Settings.Default.Visible_grpClipboard;
                    case "grpFormatDataTable":
                        return Properties.Settings.Default.Visible_grpFormatDataTable;
                    case "grpScriptVariables":
                        return Properties.Settings.Default.Visible_grpScriptVariables;
                    case "btnClearInteriorColor":
                        return Properties.Settings.Default.Visible_btnClearInteriorColor;
                    case "btnZeroToNull":
                        return Properties.Settings.Default.Visible_btnZeroToNull;
                    case "btnSeparateValues":
                        return Properties.Settings.Default.Visible_btnSeparateValues;
                    case "btnFileList":
                        return Properties.Settings.Default.Visible_btnFileList;
                    case "btnScriptTypeTSqlCreateTable":
                    case "btnScriptTypeTSqlInsertValues":
                    case "btnScriptTypeTSqlMergeValues":
                    case "btnScriptTypeTSqlSelectValues":
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btnScriptTypeTSqlUpdateValues":
                        return Properties.Settings.Default.Visible_mnuScriptType_TSQL;
                    case "btnScriptTypePlSqlCreateTable":
                    case "btnScriptTypePlSqlInsertValues":
                    case "btnScriptTypePlSqlMergeValues":
                    case "btnScriptTypePlSqlSelectValues":
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btnScriptTypePlSqlUpdateValues":
                        return Properties.Settings.Default.Visible_mnuScriptType_PLSQL;
                    case "btnScriptTypeDqlAppend":
                    case "btnScriptTypeDqlAppendLocked":
                    case "btnScriptTypeDqlCreate":
                    case "btnScriptTypeDqlTruncateAppend":
                    case "btnScriptTypeDqlUpdate":
                    case "btnScriptTypeDqlUpdateLocked":
                        return Properties.Settings.Default.Visible_mnuScriptType_DQL;
                    case "btnScriptTypeGithubTable":
                        return Properties.Settings.Default.Visible_mnuScriptType_Github;
                    case "btnScriptTypeHtmlTable":
                    case "btnScriptTypeXmlValues":
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
                    case "btnStart":
                        OpenGraphData();
                        break;
                    case "btnCopyVisibleCells":
                        CopyVisibleCells();
                        break;
                    case "btnCleanData":
                        CleanData();
                        break;
                    case "btnZeroToNull":
                        ZeroStringToNull();
                        break;
                    case "btnFormatDateColumns":
                        FormatDateColumns();
                        break;
                    case "btnClearInteriorColor":
                        ClearInteriorColor();
                        break;
                    case "btnSeparateValues":
                        SeparateValues();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnFileList":
                        CreateFileList();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                    case "btnDownloadNewVersion":
                        //DownloadNewVersion();
                        break;
                    case "btnScriptTypeDqlAppend":
                        Formula.DqlAppend();
                        break;
                    case "btnScriptTypeDqlAppendLocked":
                        Formula.DqlAppendLocked();
                        break;
                    case "btnScriptTypeDqlCreate":
                        Formula.DqlCreate();
                        break;
                    case "btnScriptTypeDqlTruncateAppend":
                        Formula.DqlTruncateAppend();
                        break;
                    case "btnScriptTypeDqlUpdate":
                        Formula.DqlUpdate();
                        break;
                    case "btnScriptTypeDqlUpdateLocked":
                        Formula.DqlUpdateLocked();
                        break;
                    case "btnScriptTypeGithubTable":
                        Formula.GithubTable();
                        break;
                    case "btnScriptTypeHtmlTable":
                        Formula.HtmlTable();
                        break;
                    case "btnScriptTypePlSqlCreateTable":
                        Formula.PlSqlCreateTable();
                        break;
                    case "btnScriptTypePlSqlInsertValues":
                        Formula.PlSqlInsertValues();
                        break;
                    case "btnScriptTypePlSqlMergeValues":
                        Formula.PlSqlMergeValues();
                        break;
                    case "btnScriptTypePlSqlSelectValues":
                        Formula.PlSqlSelectValues();
                        break;
                    case "btnScriptTypePlSqlSelectUnion":
                        Formula.PlSqlSelectUnion();
                        break;
                    case "btnScriptTypePlSqlUpdateValues":
                        Formula.PlSqlUpdateValues();
                        break;
                    case "btnScriptTypeTSqlCreateTable":
                        Formula.TSqlCreateTable();
                        break;
                    case "btnScriptTypeTSqlInsertValues":
                        Formula.TSqlInsertValues();
                        break;
                    case "btnScriptTypeTSqlMergeValues":
                        Formula.TSqlMergeValues();
                        break;
                    case "btnScriptTypeTSqlSelectValues":
                        Formula.TSqlSelectValues();
                        break;
                    case "btnScriptTypeTSqlSelectUnion":
                        Formula.TSqlSelectUnion();
                        break;
                    case "btnScriptTypeTSqlUpdateValues":
                        Formula.TSqlUpdateValues();
                        break;
                    case "btnScriptTypeXmlValues":
                        Formula.XmlValues();
                        break;
                    case "btnDateFormatFind":
                    case "btnDateFormatReplace":
                    case "btnTableAlias":
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
                    case "cboDateFormatReplace":
                        Properties.Settings.Default.Table_ColumnDateFormatReplace = text;
                        break;
                    case "cboDateFormatFind":
                        Properties.Settings.Default.Table_ColumnDateFormatFind = text;
                        break;
                    case "cboTableAlias":
                        Properties.Settings.Default.Table_ColumnTableAlias = text;
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
        /// <remarks></remarks>
        public void CopyVisibleCells()
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
        /// <remarks></remarks>
        public void CleanData()
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
                                        cell.Interior.Color = Properties.Settings.Default.Table_ColumnCleanedColour;
                                        cnt = cnt + 1;
                                    }
                                }
                                cell = tbl.Range.Cells[i + 1, j];
                                string qt = Properties.Settings.Default.Table_ColumnScriptQuote;
                                if (cell.PrefixCharacter == qt)  // show the leading apostrophe in the cell by doubling the value.
                                {
                                    cell.Value = qt + qt + cell.Value;
                                    cell.Interior.Color = Properties.Settings.Default.Table_ColumnCleanedColour;
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
        /// Change zero string cell values to string "NULL"
        /// </summary>
        /// <remarks></remarks>
        public void ZeroStringToNull()
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
                            cell.Value = Properties.Settings.Default.Table_ColumnScriptNull;
                            cell.Interior.Color = Properties.Settings.Default.Table_ColumnCleanedColour;
                            cnt = cnt + 1;
                        }
                    }
                }
                MessageBox.Show("The number of cells converted to " + Properties.Settings.Default.Table_ColumnScriptNull + ": " + cnt.ToString(), "Converting has finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        /// Finds dates columns with the paste format settings or date specific columns and updates with date format setting
        /// </summary>
        /// <remarks></remarks>
        public void FormatDateColumns()
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
                        if (cell.NumberFormat.ToString() == Properties.Settings.Default.Table_ColumnDateFormatFind | ErrorHandler.IsDate(cell.Value))
                        {
                            col.DataBodyRange.NumberFormat = Properties.Settings.Default.Table_ColumnDateFormatReplace;
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
        /// Convert a range of cells to a table with a specific table format
        /// </summary>
        /// <remarks></remarks>
        public void FormatAsTable()
        {
            Excel.Range range = null;
            string tableName = AssemblyInfo.Title + " " + DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss:fffzzz");
            string tableStyle = Properties.Settings.Default.Table_StyleName;
            try
            {
                if (ErrorHandler.IsValidListObject(false) == true)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                range = Globals.ThisAddIn.Application.Selection;
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

                range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = tableName;
                range.Select();
                range.Worksheet.ListObjects[tableName].TableStyle = tableStyle;
                ribbon.ActivateTab("tabScriptHelp");

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
                if (range != null)
                    Marshal.ReleaseComObject(range);
            }
        }

        /// <summary> 
        /// Remove interior cell color format
        /// </summary>
        /// <remarks></remarks>
        public void ClearInteriorColor()
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
        /// Add a row per delimited string value from a column
        /// </summary>
        /// <remarks></remarks>
        public void SeparateValues()
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
                    string cellValue = tbl.Range.Cells[i, columnIndex].Value2.ToString();
                    if (string.IsNullOrEmpty(cellValue) == false)
                    {
                        string[] metadata = cellValue.Split(Properties.Settings.Default.Table_ColumnSeparateValuesDelimiter);
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
        /// Creates a recursive file listing based on the users selected directory
        /// </summary>
        /// <remarks></remarks>
        public void CreateFileList()
        {
            string filePath = Properties.Settings.Default.Option_PathFileListing;
            try
            {
                ErrorHandler.CreateLogRecord();
                DialogResult msgDialogResult = DialogResult.None;
                FolderBrowserDialog dlg = new FolderBrowserDialog();
                if (Properties.Settings.Default.Option_PathFileListingSelect == true)
                {
                    dlg.RootFolder = Environment.SpecialFolder.MyComputer;
                    dlg.SelectedPath = filePath;
                    msgDialogResult = dlg.ShowDialog();
                    filePath = dlg.SelectedPath;
                }
                if (msgDialogResult == DialogResult.OK | Properties.Settings.Default.Option_PathFileListingSelect == false)
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
        /// Opens the settings taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenSettings()
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
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReportIssue);

        }

        /// <summary> 
        /// Opens a api help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenHelpApi()
        {
            ErrorHandler.CreateLogRecord();
            string clickOnceLocation = AssemblyInfo.GetClickOnceLocation();
            AssemblyInfo.OpenFile(Path.Combine(clickOnceLocation, @"Documentation\\Api Help.chm"));
        }

        /// <summary> 
        /// Opens the graph taskpane
        /// </summary>
        /// <remarks></remarks>
        public void OpenGraphData()
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

        #endregion

        #region | Subroutines |

        /// <summary> 
        /// Used to apply quoting based on the column type
        /// </summary>
        /// <param name="col">Represents the list column </param>
        /// <returns>A method that returns a string of a quote based on application settings for this value. </returns> 
        /// <remarks></remarks>
        public static string ApplyTextQuotes(Excel.ListColumn col)
        {
            Excel.Range cell = FirstNotNullCellInColumn(col.DataBodyRange);
            string timeFormat = "h:mm:ss AM/PM";  //TODO: reference a list of time formats
            try
            {
                if ((GetSqlDataType(col) != Properties.Settings.Default.Column_TypeNumeric)) //or date/time
                {
                    return Properties.Settings.Default.Table_ColumnScriptQuote;
                }
                else
                {
                    if (cell.NumberFormat.ToString() == timeFormat)
                    {
                        return Properties.Settings.Default.Table_ColumnScriptQuote;
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Return the list of column names in formatted string for SQL select statements
        /// </summary>
        /// <param name="rng">Represents the Excel Range value</param>
        /// <param name="tableAliasName">Table alias used to prefix column names</param>
        /// <param name="prefixChar">The prefix character for the column name e.g. "["</param>
        /// <param name="suffixChar">The suffix character for the column name e.g. "]"</param>
        /// <param name="selectionChar">The selection character for the column name e.g. ", "</param>
        /// <returns>A method that returns a string of the column names</returns>
        public static string ConcatenateColumnNames(Excel.Range rng, string tableAliasName = "", string prefixChar = "", string suffixChar = "", string selectionChar = ", ")
        {
            try
            {
                string columnNames = string.Empty;
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
        /// Return the list of column names in formatted string for SQL joins
        /// </summary>
        /// <param name="rng">Represents the Excel Range value</param>
        /// <param name="tableAliasNameTarget">Table alias used to prefix column names</param>
        /// <param name="tableAliasNameSource">Table alias used to prefix column names</param>
        /// <returns>A method that returns a string of the column names</returns>
        public static string ConcatenateColumnNamesJoin(Excel.Range rng, string tableAliasNameTarget, string tableAliasNameSource)
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
        public static Excel.Range FirstNotNullCellInColumn(Excel.Range rng)
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
                        if (String.Compare(cellValue, Properties.Settings.Default.Table_ColumnScriptNull, true) != 0)
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
        public static string FormatCellText(Excel.ListColumn col, string fmt)
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
        public static string GetColumnFormat(Excel.ListColumn col)
        {
            try
            {
                string fmt = string.Empty;
                string nFmt = string.Empty;

                switch (GetSqlDataType(col))
                {
                    case 2:
                        fmt = Properties.Settings.Default.Table_ColumnDateFormatReplace;
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
        public static int GetSqlDataType(Excel.ListColumn col)
        {
            try
            {
                // default to text
                double numCnt = 0;
                double notNullCnt = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(col.DataBodyRange, "<>" + Properties.Settings.Default.Table_ColumnScriptNull);

                // If all values are nulls then assume text
                if ((notNullCnt == 0))
                {
                    return Properties.Settings.Default.Column_TypeText;
                }

                numCnt = Globals.ThisAddIn.Application.WorksheetFunction.Count(col.DataBodyRange);
                // if no numbers then assume text
                if ((numCnt == 0))
                {
                    return Properties.Settings.Default.Column_TypeText;
                }

                // if a mix of numbers and not numbers then assume text
                if ((numCnt != notNullCnt))
                {
                    return Properties.Settings.Default.Column_TypeText;
                }

                //Excel changes the case of date formats on custom cell format types
                bool result = Properties.Settings.Default.Table_ColumnDateFormatReplace.Equals(col.DataBodyRange.NumberFormat.ToString(), StringComparison.OrdinalIgnoreCase);
                // NOTE: next test relies consistent formatting on numerics in a column
                // so we only have to test the first cell
                if (ErrorHandler.IsDate(FirstNotNullCellInColumn(col.DataBodyRange)) | result == true)
                {
                    return Properties.Settings.Default.Column_TypeDate;
                }
                else
                {
                    return Properties.Settings.Default.Column_TypeNumeric;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return Properties.Settings.Default.Column_TypeText;
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
        /// Used to update/reset the ribbon values
        /// </summary>
        public void InvalidateRibbon()
        {
            ribbon.Invalidate();
        }

        /// <summary>
        /// Open the taskpane for a script view
        /// </summary>
        public static void OpenScriptPane()
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
        /// Open the taskpane for the a table view list
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

        /// <summary> 
        /// Update the value of the combobox from a data table by index
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
        /// Update the value of the combobox from a data table by index
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

        #endregion

    }
}