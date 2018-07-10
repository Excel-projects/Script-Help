using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScriptHelp.Scripts
{
    class Formula
    {
        //static readonly bool isListObject;

        static Formula()
        {
            //isListObject = ErrorHandler.IsValidListObject(true);
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlAppend()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
                        }
                        else
                        {
                            formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                        }
                    }
                }
                //replace NULL values with DQL format
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlAppendLocked()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            if (afterWhere == true)
                            {
                                whereClause += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";

                            }
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
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
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

                string whereCheck = string.Empty;
                if (afterWhere == false)
                {
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlCreate()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        string valuePrefix = string.Empty;
                        valuePrefix = "\" SET " + col.Name + " = \" & ";
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
                        }
                        else
                        {
                            formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                        }
                    }
                }
                //replace NULL values with DQL format
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlTruncateAppend()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
                        }
                        else
                        {
                            formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                        }
                    }
                }
                //replace NULL values with DQL format
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

                string whereCheck = string.Empty;
                if (afterWhere == false)
                {
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlUpdate()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
                        }
                        else
                        {
                            formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                        }
                    }
                }
                //replace NULL values with DQL format
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void DqlUpdateLocked()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            if (afterWhere == true)
                            {
                                whereClause += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";

                            }
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
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
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "', '" + Properties.Settings.Default.Table_ColumnFormatDate + "'\", \"'nulldate'\")";
                formula = "SUBSTITUTE(" + formula + ", \"'" + Properties.Settings.Default.Table_ColumnScriptNull + "'\", \"nullstring\")";
                formula = "SUBSTITUTE(" + formula + ", \"" + Properties.Settings.Default.Table_ColumnScriptNull + "\", \"nullint\")";

                string whereCheck = string.Empty;
                if (afterWhere == false)
                {
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    Ribbon.AppVariables.FileType = "DQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void MarkdownTable()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;

                string sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        formula += Ribbon.GetColumnFormat(col).ToString();
                    }
                }
                formula = "=\"" + "|" + "\" & " + formula + " & \"|\"";
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.DataBodyRange.Copy();
                    Ribbon.AppVariables.FileType = "TXT";
					string headerColumn = Ribbon.ConcatenateColumnNames(tbl.Range, string.Empty, "|", string.Empty, string.Empty) + "|" + Environment.NewLine;
                    string headerSeparator = "|:" + new String('-', 10);
                    string headerLine = new System.Text.StringBuilder(headerSeparator.Length * lastColumnIndex).Insert(0, headerSeparator, lastColumnIndex).ToString().Substring(0, ((headerSeparator.Length * lastColumnIndex) - (headerSeparator.Length - 1))) + Environment.NewLine;
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("Markdown table", "<!---" , "--->") + headerColumn + headerLine + (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }

        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void HtmlTable()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = string.Empty; // Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";

                        string valuePrefix = string.Empty;
                        string valueSuffix = string.Empty;
                        string columnName = col.Name.ToLower();
                        string firstPart = columnName.Substring(0, columnName.LastIndexOf(" ") + 1);

                        valuePrefix = " CHAR(10) & \"<td>\" & ";
                        valueSuffix = " & \"</td>\" ";
                        formula += valuePrefix + "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valueSuffix;
                    }
                }

                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                formula = "=\"<tr> \" & " + formula + " & CHAR(10) & \"</tr> \"";
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
                    Ribbon.AppVariables.FileType = "XML";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("HTML table", "<!---", "--->") + "<table>" + Environment.NewLine + "<tr>" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "<th>", "</th>", Environment.NewLine) + Environment.NewLine + "</tr>" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + "</table>";
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlCreateTable()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string insertPrefix = "INSERT INTO " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES(";
                formula = "=\"" + insertPrefix + "\" & " + formula + " & \");\"";
                tbl.ShowTotals = false;
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                string createTable = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'" + tableAlias + "') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE " + tableAlias + Environment.NewLine + "; " + Environment.NewLine + "CREATE TABLE " + tableAlias + " (" + tableAlias + "_ID [int] PRIMARY KEY IDENTITY(1,1) NOT NULL, " + Ribbon.ConcatenateColumnNames(tbl.Range, "", Environment.NewLine + "[", "] [varchar](max) NULL") + Environment.NewLine + ");" + Environment.NewLine;
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.DataBodyRange.Copy();
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = createTable + (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlInsertValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string insertPrefix = "INSERT INTO " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES(";
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlMergeValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAliasTemp = tableAlias + "_source";
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
                formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                tbl.ShowTotals = true;
                string totalsColumnValue = ") " + tableAliasTemp + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
                tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.Range.Copy();
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                    Ribbon.AppVariables.ScriptRange = "SET XACT_ABORT ON" + Environment.NewLine + "BEGIN TRANSACTION;" + Environment.NewLine + Environment.NewLine + ";WITH " + Environment.NewLine + tableAliasTemp + Environment.NewLine + "AS " + Environment.NewLine + "(" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + ") " + Environment.NewLine + "MERGE " + tableAlias + " AS T" + Environment.NewLine + "USING " + tableAliasTemp + " AS S" + Environment.NewLine + "ON " + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "WHEN NOT MATCHED BY TARGET" + Environment.NewLine + "THEN INSERT" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "VALUES" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "S", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "WHEN MATCHED" + Environment.NewLine + "THEN UPDATE SET" + Environment.NewLine + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "--WHEN NOT MATCHED BY SOURCE AND 'ADD WHERE CLAUSE HERE'" + Environment.NewLine + "--THEN DELETE" + Environment.NewLine + "OUTPUT $action, inserted.*, deleted.*;" + Environment.NewLine + Environment.NewLine + "ROLLBACK TRANSACTION;" + Environment.NewLine + "--COMMIT TRANSACTION;" + Environment.NewLine + "GO";
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlSelectValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
                formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                tbl.ShowTotals = true;
                string totalsColumnValue = ") " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
                tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.Range.Copy();
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlSelectUnion()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        string dqt = "\"\"";
                        string valuePlSuffix = "& \" AS " + dqt + col.Name + dqt + " \"";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valuePlSuffix;
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void PlSqlUpdateValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
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
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlCreateTable()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string insertPrefix = "INSERT INTO " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES(";
                formula = "=\"" + insertPrefix + "\" & " + formula + " & \");\"";
                tbl.ShowTotals = false;
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                string createTable = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'" + tableAlias + "') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE " + tableAlias + Environment.NewLine + "; " + Environment.NewLine + "CREATE TABLE " + tableAlias + " (" + tableAlias + "_ID [int] PRIMARY KEY IDENTITY(1,1) NOT NULL, " + Ribbon.ConcatenateColumnNames(tbl.Range, "", Environment.NewLine + "[", "] [varchar](max) NULL") + Environment.NewLine + ");" + Environment.NewLine;
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.DataBodyRange.Copy();
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = createTable + (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To create and insert records") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlInsertValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string insertPrefix = "INSERT INTO " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES(";
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To insert records") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlMergeValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAliasTemp = tableAlias + "_source";
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
                formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                tbl.ShowTotals = true;
                string totalsColumnValue = ") " + tableAliasTemp + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
                tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.Range.Copy();
					Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To update, insert & delete rows") + "SET XACT_ABORT ON" + Environment.NewLine + "BEGIN TRANSACTION;" + Environment.NewLine + Environment.NewLine + ";WITH " + Environment.NewLine + tableAliasTemp + Environment.NewLine + "AS " + Environment.NewLine + "(" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + ") " + Environment.NewLine + "MERGE " + tableAlias + " AS T" + Environment.NewLine + "USING " + tableAliasTemp + " AS S" + Environment.NewLine + "ON " + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "WHEN NOT MATCHED BY TARGET" + Environment.NewLine + "THEN INSERT" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "VALUES" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "S", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "WHEN MATCHED" + Environment.NewLine + "THEN UPDATE SET" + Environment.NewLine + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "--WHEN NOT MATCHED BY SOURCE AND 'ADD WHERE CLAUSE HERE'" + Environment.NewLine + "--THEN DELETE" + Environment.NewLine + "OUTPUT $action, inserted.*, deleted.*;" + Environment.NewLine + Environment.NewLine + "ROLLBACK TRANSACTION;" + Environment.NewLine + "--COMMIT TRANSACTION;" + Environment.NewLine + "GO";
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlSelectUnion()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                sqlColName = Properties.Settings.Default.Table_ColumnName;

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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        string valueTsuffix = "& \" AS [" + col.Name + "] \"";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"" + valueTsuffix;
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To select values with a union operator") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlSelectValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
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
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
                        colRef = colRef.Replace("'", "''");
                        colRef = colRef.Replace("#", "'#");
                        colRef = "SUBSTITUTE(" + colRef + ", " + "\"" + qt + "\", \"" + qt + qt + "\")";
                        formula += "\"" + qt + "\" & " + colRef + " & \"" + qt + "\"";
                    }
                }
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
                formula = "SUBSTITUTE(" + formula + ", \"'" + nullValue + "'\", \"" + nullValue + "\")";
                int firstRowNbr = tbl.Range[1, 1].Row + 1; // must use the offset for the first row number
                formula = "=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, \" \", \",\") & " + "\" ( \" & " + formula + " & \")\"";
                lastColumnName = sqlColName;  // maximum header characters are 255
                tbl.HeaderRowRange[lastColumnIndex].Value2 = lastColumnName;
                tbl.ShowTotals = true;
                string totalsColumnValue = ") " + tableAlias + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") ";
                tbl.TotalsRowRange[lastColumnIndex].Value2 = totalsColumnValue; // totals row has a maximum limit of 32,767 characters
                try
                {
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula;
                    sqlCol.Range.Columns.AutoFit();
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft;
                    sqlCol.Range.Copy();
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To select values") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void TSqlUpdateValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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
                        if (Ribbon.GetSqlDataType(col) == Properties.Settings.Default.Column_TypeDate)
                        {
                            formula += valuePrefix + "\"DATE(" + qt + "\" & " + colRef + " & \"" + qt + ", " + qt + Properties.Settings.Default.Table_ColumnFormatDate + qt + ")\"";
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
                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
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
                    Ribbon.AppVariables.FileType = "SQL";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("To update records") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

        /// <summary> 
        /// Add a formula at the end of the table to use as a script
        /// </summary>
        /// <remarks></remarks>
        public static void XmlValues()
        {
            Excel.ListObject tbl = null;
            Excel.ListColumn sqlCol = null;
            try
            {
                ErrorHandler.CreateLogRecord();
                if (ErrorHandler.IsValidListObject(true) == false) { return; };
                string lastColumnName = Properties.Settings.Default.Table_ColumnTableAlias;
                string tableAlias = Properties.Settings.Default.Table_ColumnTableAlias;
                string sqlColName = string.Empty;
                int columnCount = 0;

                sqlColName = Properties.Settings.Default.Table_ColumnName;
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
                            Ribbon.AppVariables.FirstColumnName = col.Name;
                        }
                        columnCount += 1;
                        qt = Ribbon.ApplyTextQuotes(col);
                        string colRef = Ribbon.GetColumnFormat(col).ToString();
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

                string nullValue = Properties.Settings.Default.Table_ColumnScriptNull;
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
                    Ribbon.AppVariables.FileType = "XML";
                    Ribbon.AppVariables.ScriptRange = (string)Clipboard.GetData(DataFormats.Text);
                    Ribbon.AppVariables.ScriptRange = Ribbon.GetCommentHeader("XML table", "<!---", "--->") + Ribbon.AppVariables.ScriptRange.Replace(@"""", String.Empty);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    Ribbon.AppVariables.ScriptRange = "There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine + formula;
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
                Ribbon.OpenScriptPane();
            }
        }

    }
}
