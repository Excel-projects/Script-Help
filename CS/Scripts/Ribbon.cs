using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScriptHelp.Scripts
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public static Ribbon ribbonref;

        public static class AppVariables
        {

            public static string ScriptRange { get; set; }
            public static string FileType { get; set; }
            public static string TableName { get; set; }
            public static string FirstColumnName { get; set; }
            public static string ControlLabel { get; set; }

        }

        #region | Task Panes |

        public TaskPane.Settings mySettings;
        public static TaskPane.Script myScript;
        public TaskPane.TableData myTableData;
        public TaskPane.GraphData myGraphData;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;
        public static Microsoft.Office.Tools.CustomTaskPane myTaskPaneScript;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneTableData;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneGraphData;

        #endregion

        #region | Ribbon Events |

        public Ribbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ScriptHelp.Ribbon.xml");
        }

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
                Data.CreateTimeFormatTable();
                Data.CreateGraphDataTable();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

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
                    case "btn1ScriptTypeDqlAppend":
                    case "btn1ScriptTypeDqlAppendLocked":
                    case "btn1ScriptTypeDqlCreate":
                    case "btn1ScriptTypeDqlTruncateAppend":
                    case "btn1ScriptTypeDqlUpdate":
                    case "btn1ScriptTypeDqlUpdateLocked":
                        return Properties.Resources.ScriptTypeDql;
                    case "btnScriptTypeTSqlCreateTable":
                    case "btnScriptTypeTSqlInsertValues":
                    case "btnScriptTypeTSqlMergeValues":
                    case "btnScriptTypeTSqlSelectValues":
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btnScriptTypeTSqlUpdateValues":
                    case "btn1ScriptTypeTSqlCreateTable":
                    case "btn1ScriptTypeTSqlInsertValues":
                    case "btn1ScriptTypeTSqlMergeValues":
                    case "btn1ScriptTypeTSqlSelectValues":
                    case "btn1ScriptTypeTSqlSelectUnion":
                    case "btn1ScriptTypeTSqlUpdateValues":
                        return Properties.Resources.ScriptTypeTSql;
                    case "btnScriptTypePlSqlCreateTable":
                    case "btnScriptTypePlSqlInsertValues":
                    case "btnScriptTypePlSqlMergeValues":
                    case "btnScriptTypePlSqlSelectValues":
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btnScriptTypePlSqlUpdateValues":
                    case "btn1ScriptTypePlSqlCreateTable":
                    case "btn1ScriptTypePlSqlInsertValues":
                    case "btn1ScriptTypePlSqlMergeValues":
                    case "btn1ScriptTypePlSqlSelectValues":
                    case "btn1ScriptTypePlSqlSelectUnion":
                    case "btn1ScriptTypePlSqlUpdateValues":
                        return Properties.Resources.ScriptTypePlSql;
                    case "btnScriptTypeMarkdownTable":
                    case "btn1ScriptTypeMarkdownTable":
                        return Properties.Resources.ScriptTypeMarkdown;
                    case "btnScriptTypeHtmlTable":
                    case "btnScriptTypeXmlValues":
                    case "btn1ScriptTypeHtmlTable":
                    case "btn1ScriptTypeXmlValues":
                        return Properties.Resources.ScriptTypeMarkup;
                    case "btnProblemStepRecorder":
                    case "btnProblemStepRecorder1":
                        return Properties.Resources.problem_steps_recorder;
                    case "btnSnippingTool":
                    case "btnSnippingTool1":
                        return Properties.Resources.snipping_tool;
                    case "btnSaveVersion":
                        return Properties.Resources.SaveVersion;
                    case "btnSaveCode":
                        return Properties.Resources.SaveCode;
                    case "btnCamera":
                        return Properties.Resources.camera;
                    case "btn1Start":
                        return Properties.Resources.Play;
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
                    case "btnFormatTimeColumns":
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
                    case "btnScriptTypeDqlAppend":
                    case "btn1ScriptTypeDqlAppend":
                        return "DQL Append";
                    case "btnScriptTypeDqlAppendLocked":
                    case "btn1ScriptTypeDqlAppendLocked":
                        return "DQL Append/Locked";
                    case "btnScriptTypeDqlCreate":
                    case "btn1ScriptTypeDqlCreate":
                        return "DQL Create";
                    case "btnScriptTypeDqlTruncateAppend":
                    case "btn1ScriptTypeDqlTruncateAppend":
                        return "DQL Truncate/Append";
                    case "btnScriptTypeDqlUpdate":
                    case "btn1ScriptTypeDqlUpdate":
                        return "DQL Update";
                    case "btnScriptTypeDqlUpdateLocked":
                    case "btn1ScriptTypeDqlUpdateLocked":
                        return "DQL Update/Locked";
                    case "btnScriptTypeMarkdownTable":
                    case "btn1ScriptTypeMarkdownTable":
                        return "Markdown Table";
                    case "btnScriptTypeHtmlTable":
                    case "btn1ScriptTypeHtmlTable":
                        return "HTML Table";
                    case "btnScriptTypePlSqlCreateTable":
                    case "btn1ScriptTypePlSqlCreateTable":
                        return "PL/SQL Create Table";
                    case "btnScriptTypePlSqlInsertValues":
                    case "btn1ScriptTypePlSqlInsertValues":
                        return "PL/SQL Insert Values";
                    case "btnScriptTypePlSqlMergeValues":
                    case "btn1ScriptTypePlSqlMergeValues":
                        return "PL/SQL Merge Values";
                    case "btnScriptTypePlSqlSelectValues":
                    case "btn1ScriptTypePlSqlSelectValues":
                        return "PL/SQL Select Values";
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btn1ScriptTypePlSqlSelectUnion":
                        return "PL/SQL Select Union";
                    case "btnScriptTypePlSqlUpdateValues":
                    case "btn1ScriptTypePlSqlUpdateValues":
                        return "PL/SQL Update Values";
                    case "btnScriptTypeTSqlCreateTable":
                    case "btn1ScriptTypeTSqlCreateTable":
                        return "T-SQL Create Table";
                    case "btnScriptTypeTSqlInsertValues":
                    case "btn1ScriptTypeTSqlInsertValues":
                        return "T-SQL Insert Values";
                    case "btnScriptTypeTSqlMergeValues":
                    case "btn1ScriptTypeTSqlMergeValues":
                        return "T-SQL Merge Values";
                    case "btnScriptTypeTSqlSelectValues":
                    case "btn1ScriptTypeTSqlSelectValues":
                        return "T-SQL Select Values";
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btn1ScriptTypeTSqlSelectUnion":
                        return "T-SQL Select Union";
                    case "btnScriptTypeTSqlUpdateValues":
                    case "btn1ScriptTypeTSqlUpdateValues":
                        return "T-SQL Update Values";
                    case "btnScriptTypeXmlValues":
                    case "btn1ScriptTypeXmlValues":
                        return "XML Values";
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

        public int GetItemCount(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboFormatDate":
                        return Data.DateFormatTable.Rows.Count;
                    case "cboFormatTime":
                        return Data.TimeFormatTable.Rows.Count;
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

        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboFormatDate":
                        return UpdateDateFormatComboBoxSource(index);
                    case "cboFormatTime":
                        return UpdateTimeFormatComboBoxSource(index);
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

        public string GetText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboFormatDate":
                        return Properties.Settings.Default.Table_ColumnFormatDate;
                    case "cboFormatTime":
                        return Properties.Settings.Default.Table_ColumnFormatTime;
                    case "cboTableAlias":
                        return Properties.Settings.Default.Table_ColumnTableAlias;
                    case "txtColumnSeparateValuesDelimiter":
                        return Properties.Settings.Default.Table_ColumnSeparateValuesDelimiter.ToString();
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

        public bool GetPressed(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {

                    case "chkBackstageTsql":
                        return Properties.Settings.Default.Visible_mnuScriptType_TSQL;
                    case "chkBackstagePlsql":
                        return Properties.Settings.Default.Visible_mnuScriptType_PLSQL;
                    case "chkBackstageDql":
                        return Properties.Settings.Default.Visible_mnuScriptType_DQL;
                    case "chkBackstageMarkdown":
                        return Properties.Settings.Default.Visible_mnuScriptType_Markdown;
                    case "chkBackstageMarkup":
                        return Properties.Settings.Default.Visible_mnuScriptType_Markup;
                    default:
                        return true;
                }

            }
            catch (Exception)
            {
                return true;
                //ErrorHandler.DisplayMessage(ex);
            }

        }

        public bool GetVisible(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "grpClipboard":
                        return Properties.Settings.Default.Visible_grpClipboard;
                    case "grpAnnotation":
                        return Properties.Settings.Default.Visible_grpAnnotation;
                    case "btnScriptTypeTSqlCreateTable":
                    case "btnScriptTypeTSqlInsertValues":
                    case "btnScriptTypeTSqlMergeValues":
                    case "btnScriptTypeTSqlSelectValues":
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btnScriptTypeTSqlUpdateValues":
                    case "btn1ScriptTypeTSqlCreateTable":
                    case "btn1ScriptTypeTSqlInsertValues":
                    case "btn1ScriptTypeTSqlMergeValues":
                    case "btn1ScriptTypeTSqlSelectValues":
                    case "btn1ScriptTypeTSqlSelectUnion":
                    case "btn1ScriptTypeTSqlUpdateValues":
                        return Properties.Settings.Default.Visible_mnuScriptType_TSQL;
                    case "btnScriptTypePlSqlCreateTable":
                    case "btnScriptTypePlSqlInsertValues":
                    case "btnScriptTypePlSqlMergeValues":
                    case "btnScriptTypePlSqlSelectValues":
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btnScriptTypePlSqlUpdateValues":
                    case "btn1ScriptTypePlSqlCreateTable":
                    case "btn1ScriptTypePlSqlInsertValues":
                    case "btn1ScriptTypePlSqlMergeValues":
                    case "btn1ScriptTypePlSqlSelectValues":
                    case "btn1ScriptTypePlSqlSelectUnion":
                    case "btn1ScriptTypePlSqlUpdateValues":
                        return Properties.Settings.Default.Visible_mnuScriptType_PLSQL;
                    case "btnScriptTypeDqlAppend":
                    case "btnScriptTypeDqlAppendLocked":
                    case "btnScriptTypeDqlCreate":
                    case "btnScriptTypeDqlTruncateAppend":
                    case "btnScriptTypeDqlUpdate":
                    case "btnScriptTypeDqlUpdateLocked":
                    case "btn1ScriptTypeDqlAppend":
                    case "btn1ScriptTypeDqlAppendLocked":
                    case "btn1ScriptTypeDqlCreate":
                    case "btn1ScriptTypeDqlTruncateAppend":
                    case "btn1ScriptTypeDqlUpdate":
                    case "btn1ScriptTypeDqlUpdateLocked":
                        return Properties.Settings.Default.Visible_mnuScriptType_DQL;
                    case "btnScriptTypeMarkdownTable":
                    case "btn1ScriptTypeMarkdownTable":
                        return Properties.Settings.Default.Visible_mnuScriptType_Markdown;
                    case "btnScriptTypeHtmlTable":
                    case "btnScriptTypeXmlValues":
                    case "btn1ScriptTypeHtmlTable":
                    case "btn1ScriptTypeXmlValues":
                        return Properties.Settings.Default.Visible_mnuScriptType_Markup;
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

        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                Ribbon.AppVariables.ControlLabel = GetLabelText(control);
                switch (control.Id)
                {
                    case "btnStart":
                    case "btn1Start":
                        OpenGraphData();
                        break;
                    case "btnCopyVisibleCells":
                        CopyVisibleCells();
                        break;
                    case "btnCleanData":
                    case "btn1CleanData":
                        CleanData();
                        break;
                    case "btnZeroToNull":
                        ZeroStringToNull();
                        break;
                    case "btnFormatDateColumns":
                    case "btn1FormatDateColumns":
                        FormatDateColumns();
                        break;
                    case "btnFormatDateColumnsAll":
                        FormatDateColumnsAll();
                        break;
                    case "btnFormatTimeColumns":
                    case "btn1FormatTimeColumns":
                        FormatTimeColumns();
                        break;
                    case "btnClearInteriorColor":
                        ClearInteriorColor();
                        break;
                    case "btnSeparateValues":
                    case "btn1SeparateValues":
                        SeparateValues();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                    case "btnScriptTypeDqlAppend":
                    case "btn1ScriptTypeDqlAppend":
                        Formula.DqlAppend();
                        break;
                    case "btnScriptTypeDqlAppendLocked":
                    case "btn1ScriptTypeDqlAppendLocked":
                        Formula.DqlAppendLocked();
                        break;
                    case "btnScriptTypeDqlCreate":
                    case "btn1ScriptTypeDqlCreate":
                        Formula.DqlCreate();
                        break;
                    case "btnScriptTypeDqlTruncateAppend":
                    case "btn1ScriptTypeDqlTruncateAppend":
                        Formula.DqlTruncateAppend();
                        break;
                    case "btnScriptTypeDqlUpdate":
                    case "btn1ScriptTypeDqlUpdate":
                        Formula.DqlUpdate();
                        break;
                    case "btnScriptTypeDqlUpdateLocked":
                    case "btn1ScriptTypeDqlUpdateLocked":
                        Formula.DqlUpdateLocked();
                        break;
                    case "btnScriptTypeMarkdownTable":
                    case "btn1ScriptTypeMarkdownTable":
                        Formula.MarkdownTable();
                        break;
                    case "btnScriptTypeHtmlTable":
                    case "btn1ScriptTypeHtmlTable":
                        Formula.HtmlTable();
                        break;
                    case "btnScriptTypePlSqlCreateTable":
                    case "btn1ScriptTypePlSqlCreateTable":
                        Formula.PlSqlCreateTable();
                        break;
                    case "btnScriptTypePlSqlInsertValues":
                    case "btn1ScriptTypePlSqlInsertValues":
                        Formula.PlSqlInsertValues();
                        break;
                    case "btnScriptTypePlSqlMergeValues":
                    case "btn1ScriptTypePlSqlMergeValues":
                        Formula.PlSqlMergeValues();
                        break;
                    case "btnScriptTypePlSqlSelectValues":
                    case "btn1ScriptTypePlSqlSelectValues":
                        Formula.PlSqlSelectValues();
                        break;
                    case "btnScriptTypePlSqlSelectUnion":
                    case "btn1ScriptTypePlSqlSelectUnion":
                        Formula.PlSqlSelectUnion();
                        break;
                    case "btnScriptTypePlSqlUpdateValues":
                    case "btn1ScriptTypePlSqlUpdateValues":
                        Formula.PlSqlUpdateValues();
                        break;
                    case "btnScriptTypeTSqlCreateTable":
                    case "btn1ScriptTypeTSqlCreateTable":
                        Formula.TSqlCreateTable();
                        break;
                    case "btnScriptTypeTSqlInsertValues":
                    case "btn1ScriptTypeTSqlInsertValues":
                        Formula.TSqlInsertValues();
                        break;
                    case "btnScriptTypeTSqlMergeValues":
                    case "btn1ScriptTypeTSqlMergeValues":
                        Formula.TSqlMergeValues();
                        break;
                    case "btnScriptTypeTSqlSelectValues":
                    case "btn1ScriptTypeTSqlSelectValues":
                        Formula.TSqlSelectValues();
                        break;
                    case "btnScriptTypeTSqlSelectUnion":
                    case "btn1ScriptTypeTSqlSelectUnion":
                        Formula.TSqlSelectUnion();
                        break;
                    case "btnScriptTypeTSqlUpdateValues":
                    case "btn1ScriptTypeTSqlUpdateValues":
                        Formula.TSqlUpdateValues();
                        break;
                    case "btnScriptTypeXmlValues":
                    case "btn1ScriptTypeXmlValues":
                        Formula.XmlValues();
                        break;
                    case "btnFormatDate":
                    case "btnFormatTime":
                    case "btnTableAlias":
                        AppVariables.TableName = control.Tag;
                        OpenTableDataPane();
                        break;
                    case "btnSnippingTool":
                    case "btnSnippingTool1":
                        OpenSnippingTool();
                        break;
                    case "btnProblemStepRecorder":
                    case "btnProblemStepRecorder1":
                        OpenProblemStepRecorder();
                        break;
                    case "btnSaveVersion":
                        Ribbon.SaveVersion();
                        break;
                    case "btnSaveCode":
                        Ribbon.SaveCode();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        public void OnAction_Checkbox(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                switch (control.Id)
                {

                    case "chkBackstageTsql":
                        Properties.Settings.Default.Visible_mnuScriptType_TSQL = pressed;
                        break;
                    case "chkBackstagePlsql":
                        Properties.Settings.Default.Visible_mnuScriptType_PLSQL = pressed;
                        break;
                    case "chkBackstageDql":
                        Properties.Settings.Default.Visible_mnuScriptType_DQL = pressed;
                        break;
                    case "chkBackstageMarkdown":
                        Properties.Settings.Default.Visible_mnuScriptType_Markdown = pressed;
                        break;
                    case "chkBackstageMarkup":
                        Properties.Settings.Default.Visible_mnuScriptType_Markup = pressed;
                        break;
                }

                ribbon.Invalidate();

            }
            catch (Exception)
            {
                //ErrorHandler.DisplayMessage(ex);
            }

        }

        public void OnChange(Office.IRibbonControl control, string text)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboFormatDate":
                        Properties.Settings.Default.Table_ColumnFormatDate = text;
                        Data.InsertRecord(Data.DateFormatTable, text);
                        break;
                    case "cboFormatTime":
                        Properties.Settings.Default.Table_ColumnFormatTime = text;
                        Data.InsertRecord(Data.TimeFormatTable, text);
                        break;
                    case "cboTableAlias":
                        Properties.Settings.Default.Table_ColumnTableAlias = text;
                        Data.InsertRecord(Data.TableAliasTable, text);
                        break;
                    case "txtColumnSeparateValuesDelimiter":
                        Properties.Settings.Default.Table_ColumnSeparateValuesDelimiter = Convert.ToChar(text);
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                Properties.Settings.Default.Save();
                ribbon.InvalidateControl(control.Id);
            }
        }

        #endregion

        #region | Ribbon Buttons |

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

        public void FormatDateColumns()
        {
            Excel.ListObject tbl = null;
            Excel.Range cell = null;
            Excel.Range cellCurrent = null;
            try
            {
                if (ErrorHandler.IsAvailable(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                cell = default(Excel.Range);
                cellCurrent = Globals.ThisAddIn.Application.ActiveCell;
                int columnIndex = cellCurrent.Column;
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                foreach (Excel.ListColumn col in tbl.ListColumns)
                {
                    cell = FirstNotNullCellInColumn(col.DataBodyRange);
                    if (((col.Index == columnIndex)))
                    {
                        col.DataBodyRange.NumberFormat = Properties.Settings.Default.Table_ColumnFormatDate;
                        col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter;
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

        public void FormatDateColumnsAll()
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
                        if (cell.NumberFormat.ToString() == Properties.Settings.Default.Table_ColumnFormatDatePaste | ErrorHandler.IsDate(cell.Value))
                        {
                            col.DataBodyRange.NumberFormat = Properties.Settings.Default.Table_ColumnFormatDate;
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

        public void FormatTimeColumns()
        {
            Excel.ListObject tbl = null;
            Excel.Range cell = null;
            Excel.Range cellCurrent = null;
            try
            {
                if (ErrorHandler.IsAvailable(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                cell = default(Excel.Range);
                cellCurrent = Globals.ThisAddIn.Application.ActiveCell;
                int columnIndex = cellCurrent.Column;
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                foreach (Excel.ListColumn col in tbl.ListColumns)
                {
                    cell = FirstNotNullCellInColumn(col.DataBodyRange);
                    if (((col.Index == columnIndex)))
                    {
                        col.DataBodyRange.NumberFormat = Properties.Settings.Default.Table_ColumnFormatTime;
                        col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter;
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
                    if (tbl.Range.Cells[i, columnIndex].Value2 != null)
                    {
                        string cellValue = tbl.Range.Cells[i, columnIndex].Value2.ToString();
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

        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);

        }

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

        public void OpenSnippingTool()
        {
            string filePath;
            try
            {
                if (System.Environment.Is64BitOperatingSystem)
                {
                    filePath = "C:\\Windows\\sysnative\\SnippingTool.exe";
                }
                else
                {
                    filePath = "C:\\Windows\\system32\\SnippingTool.exe";
                }
                System.Diagnostics.Process.Start(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenProblemStepRecorder()
        {
            string filePath = @"C:\Windows\System32\psr.exe";
            try
            {
                System.Diagnostics.Process.Start(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public static void SaveVersion()
        {
            try
            {
                var filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName; //AssemblyInfo.GetCurrentFileName();
                filePath = string.Concat(
                    Path.GetFileNameWithoutExtension(filePath),
                    "_",
                    DateTime.Now.ToString("yyyy.MM.dd_HH.mm.ss.fff"),
                    "_",
                    Environment.UserName,
                    Path.GetExtension(filePath)
                );
                Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public static void SaveCode()
        {
            try
            {
                //update this for VBA and XML part of current file
                var filePath = AssemblyInfo.GetCurrentFileName();
                filePath = string.Concat(
                    Path.GetFileNameWithoutExtension(filePath),
                    "_",
                    DateTime.Now.ToString("yyyy.MM.dd_HH.mm.ss.fff"),
                    "_",
                    Environment.UserName,
                    Path.GetExtension(filePath)
                );
                Globals.ThisAddIn.Application.ActiveWorkbook.SaveCopyAs(filePath);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        #endregion

        #region | Subroutines |

        public static string ApplyTextQuotes(Excel.ListColumn col)
        {
            Excel.Range cell = FirstNotNullCellInColumn(col.DataBodyRange);
            string timeFormat = Properties.Settings.Default.Table_ColumnFormatTime;
            try
            {
                if ((GetSqlDataType(col) != Properties.Settings.Default.Column_TypeNumeric)) 
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

        public static string ConcatenateColumnNamesJoin(Excel.Range rng, string tableAliasNameTarget, string tableAliasNameSource, string joinPrefix)
        {
            try
            {
                string columnNames = string.Empty;
                for (int i = 1; i <= rng.Columns.Count - 1; i++)
                {
                    if (rng.Columns.EntireColumn[i].Hidden == false)
                    {
                        columnNames = columnNames + joinPrefix + tableAliasNameTarget + ".[" + ((Excel.Range)rng.Cells[1, i]).Value2 + "] = " + tableAliasNameSource + ".[" + ((Excel.Range)rng.Cells[1, i]).Value2 + "]" + Environment.NewLine;
                    }
                }
                columnNames = new string(' ', joinPrefix.Length) + columnNames.Substring(joinPrefix.Length, columnNames.Length - joinPrefix.Length);
                return columnNames;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

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

        public static string GetColumnFormat(Excel.ListColumn col)
        {
            try
            {
                string fmt = string.Empty;
                string nFmt = string.Empty;

                switch (GetSqlDataType(col))
                {
                    case 2:
                        fmt = Properties.Settings.Default.Table_ColumnFormatDate;
                        return FormatCellText(col, fmt);
                    case 1:
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
                bool result = Properties.Settings.Default.Table_ColumnFormatDate.Equals(col.DataBodyRange.NumberFormat.ToString(), StringComparison.OrdinalIgnoreCase);
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

        public int GetListItemCount(string valueList, string delimiter)
        {
            try
            {
                string[] comboList = valueList.Split(delimiter.ToCharArray());
                return comboList.Length;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return 0;
            }
        }

        public void InvalidateRibbon()
        {
            ribbon.Invalidate();
        }

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
                myTaskPaneScript = Globals.ThisAddIn.CustomTaskPanes.Add(myScript, Ribbon.AppVariables.ControlLabel);
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
                myTaskPaneTableData = Globals.ThisAddIn.CustomTaskPanes.Add(myTableData, Ribbon.AppVariables.TableName);
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

        public string UpdateTimeFormatComboBoxSource(int itemIndex)
        {
            try
            {
                return Data.TimeFormatTable.Rows[itemIndex]["FormatString"].ToString();
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

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

        public static string GetCommentHeader(string purposeLine = "", string prefix = "/*", string suffix = "*/")
        {
            string noteLine = string.Concat("Generated from ", AssemblyInfo.Title, " ", AssemblyInfo.FileVersion, "  ", DateTime.Now.ToString("yyyy.MM.dd HH.mm.ss"));
            string dividerLine = string.Concat(System.Linq.Enumerable.Repeat("*", 75));
            string headerComment = prefix + dividerLine + Environment.NewLine;
            headerComment += "** Purpose:  " + purposeLine + Environment.NewLine;
            headerComment += "** Note:     " + noteLine + Environment.NewLine;
            headerComment += dividerLine + suffix + Environment.NewLine + Environment.NewLine;
            return headerComment;
        }

        #endregion

    }
}