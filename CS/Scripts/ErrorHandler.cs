using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using log4net;
using log4net.Config;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

// <summary> 
// This namespaces if for generic application classes
// </summary>
namespace ScriptHelp.Scripts
{
    /// <summary> 
    /// Used to handle exceptions
    /// </summary>
    public static class ErrorHandler
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorHandler));

        /// <summary>
        /// Applies a new path for the log file by FileAppender name
        /// </summary>
        public static void SetLogPath()
        {
            XmlConfigurator.Configure();
            log4net.Repository.Hierarchy.Hierarchy h = (log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository();
            string logFileName = System.IO.Path.Combine(Properties.Settings.Default.App_PathLocalData, AssemblyInfo.Product + ".log");

            foreach (var a in h.Root.Appenders)
            {
                if (a is log4net.Appender.FileAppender)
                {
                    if (a.Name.Equals("FileAppender"))
                    {
                        log4net.Appender.FileAppender fa = (log4net.Appender.FileAppender)a;
                        fa.File = logFileName;
                        fa.ActivateOptions();
                    }
                }
            }
        }

        /// <summary>
        /// Create a log record to track which methods are being used.
        /// </summary>
        public static void CreateLogRecord()
        {
            try
            {
                // gather context
                var sf = new System.Diagnostics.StackFrame(1);
                var caller = sf.GetMethod();
                var currentProcedure = caller.Name.Trim();

                // handle log record
                var logMessage = string.Concat(new Dictionary<string, string>
                {
                    ["PROCEDURE"] = currentProcedure,
                    ["USER NAME"] = Environment.UserName,
                    ["MACHINE NAME"] = Environment.MachineName
                }.Select(x => $"[{x.Key}]=|{x.Value}|"));
                log.Info(logMessage);

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }
        }

        /// <summary> 
        /// Used to produce an error message and create a log record
        /// <example>
        /// <code lang="C#">
        /// ErrorHandler.DisplayMessage(ex);
        /// </code>
        /// </example> 
        /// </summary>
        /// <param name="ex">Represents errors that occur during application execution.</param>
        /// <param name="isSilent">Used to show a message to the user and log an error record or just log a record.</param>
        /// <remarks></remarks>
        public static void DisplayMessage(Exception ex, Boolean isSilent = false)
        {
            // gather context
            var sf = new System.Diagnostics.StackFrame(1);
            var caller = sf.GetMethod();
            var errorDescription = ex.ToString().Replace("\r\n", " "); // the carriage returns were messing up my log file
            var currentProcedure = caller.Name.Trim();
            var currentFileName = AssemblyInfo.GetCurrentFileName();

            // handle log record
            var logMessage = string.Concat(new Dictionary<string, string>
            {
                ["PROCEDURE"] = currentProcedure,
                ["USER NAME"] = Environment.UserName,
                ["MACHINE NAME"] = Environment.MachineName,
                ["FILE NAME"] = currentFileName,
                ["DESCRIPTION"] = errorDescription,
            }.Select(x => $"[{x.Key}]=|{x.Value}|"));
            log.Error(logMessage);

            // format message
            var userMessage = new StringBuilder()
                .AppendLine("Contact your system administrator. A record has been created in the log file.")
                .AppendLine("Procedure: " + currentProcedure)
                .AppendLine("Description: " + errorDescription)
                .ToString();

            // handle message
            if (isSilent == false)
            {
                MessageBox.Show(userMessage, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary> 
        /// Check to see if there is an active document
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>A method that returns true or false if there is an active document </returns> 
        /// <remarks></remarks>
        public static bool IsActiveDocument(bool showMsg = false)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
                {
                    if (showMsg == true)
                    {
                        MessageBox.Show("The command could not be completed.  Please open a document and select a range.", AssemblyInfo.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
        }

        /// <summary> 
        /// Check to see if there is an active selection
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>A method that returns true or false if there is an active selection </returns> 
        /// <remarks></remarks>
        public static bool IsActiveSelection(bool showMsg = false)
        {
            Excel.Range checkRange = null;
            try
            {
                checkRange = Globals.ThisAddIn.Application.Selection as Excel.Range; //must cast the selection as range or errors
                if (null == checkRange)
                {
                    if (showMsg == true)
                    {
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [Range]", AssemblyInfo.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
            finally
            {
                if (checkRange != null)
                    Marshal.ReleaseComObject(checkRange);
            }
        }

        /// <summary> 
        /// Check to see if there is an named range selected
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>A method that returns true or false if there is a valid list object </returns> 
        /// <remarks></remarks>
        public static bool IsValidListObject(bool showMsg = false)
        {
            Excel.ListObject tbl = null;
            try
            {
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;  // directly after the table is created this is not true
                if ((tbl == null))
                {
                    if (showMsg == true)
                    {
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [ListObject]", AssemblyInfo.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (tbl != null)
                    Marshal.ReleaseComObject(tbl);
            }
        }

        /// <summary>
        /// This method check whether Excel is in Cell Editing mode or not
        /// There are few ways to check this (eg. check to see if a standard menu item is disabled etc.)
        /// I know in cell editing mode app.DisplayAlerts throws an Exception, so here I'm relying on that behaviour
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>true if Excel is in cell editing mode</returns>
        private static bool IsInCellEditingMode(bool showMsg = false)
        {
            bool flag = false;
            try
            {
                Globals.ThisAddIn.Application.DisplayAlerts = false; //This will throw an Exception if Excel is in Cell Editing Mode
            }
            catch (Exception)
            {
                if (showMsg == true)
                {
                    MessageBox.Show("The procedure can not run while you are editing a cell.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                flag = true;
            }
            return flag;
        }

        /// <summary> 
        /// Can an object be inserted
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>A method that returns true or false if an object can be enabled </returns> 
        /// <remarks></remarks>
        public static bool IsEnabled(bool showMsg = false)
        {
            try
            {
                if (IsActiveDocument(showMsg) == false)
                {
                    return false;
                }
                else
                {
                    if (IsActiveSelection(showMsg) == false)
                    {
                        return false;
                    }
                    else
                    {
                        if (IsInCellEditingMode(showMsg) == true)
                        {
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
        }

        /// <summary> 
        /// Can an object be inserted
        /// </summary>
        /// <param name="showMsg">To show a message </param>
        /// <returns>A method that returns true or false if an object can be used </returns> 
        /// <remarks></remarks>
        public static bool IsAvailable(bool showMsg = false)
        {
            try
            {
                if (IsEnabled(showMsg) == false)
                {
                    return false;
                }
                else
                {
                    if (IsValidListObject(showMsg) == false)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return false;
            }
        }

        /// <summary> 
        /// Check if the object is a date     
        /// </summary>
        /// <param name="expression">Represents the cell value </param>
        /// <returns>A method that returns true or false if there is a valid date </returns> 
        /// <remarks></remarks>
        public static bool IsDate(object expression)
        {
            if (expression != null)
            {
                if (expression is DateTime)
                {
                    return true;
                }
                if (expression is string)
                {
                    DateTime time1;
                    return DateTime.TryParse((string)expression, out time1);
                }
            }
            return false;
        }

        /// <summary> 
        /// Check if the object is a time     
        /// </summary>
        /// <param name="expression">Represents the cell value </param>
        /// <returns>A method that returns true or false if there is a valid time </returns> 
        /// <remarks></remarks>
        public static bool IsTime(object expression)
        {
            try
            {
                string timeValue = Convert.ToString(expression);
                //timeValue = String.Format("{0:" + Properties.Settings.Default.Table_ColumnFormatTime + "}", expression);
                //timeValue = timeValue.ToString(Properties.Settings.Default.Table_ColumnFormatTime, System.Globalization.CultureInfo.InvariantCulture);
                //timeValue = timeValue.ToString(Properties.Settings.Default.Table_ColumnFormatTime);
                //string timeValue = expression.ToString().Format(Properties.Settings.Default.Table_ColumnFormatTime);
                TimeSpan time1;
                //return TimeSpan.TryParse((string)expression, out time1);
                return TimeSpan.TryParse(timeValue, out time1);
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}