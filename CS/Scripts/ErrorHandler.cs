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

namespace ScriptHelp.Scripts
{
    public static class ErrorHandler
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(ErrorHandler));

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

        public static void CreateLogRecord()
        {
            try
            {
                var sf = new System.Diagnostics.StackFrame(1);
                var caller = sf.GetMethod();
                var currentProcedure = caller.Name.Trim();

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

        public static void DisplayMessage(Exception ex, Boolean isSilent = false)
        {
            var sf = new System.Diagnostics.StackFrame(1);
            var caller = sf.GetMethod();
            var errorDescription = ex.ToString().Replace("\r\n", " "); 
            var currentProcedure = caller.Name.Trim();
            var currentFileName = AssemblyInfo.GetCurrentFileName();

            var logMessage = string.Concat(new Dictionary<string, string>
            {
                ["PROCEDURE"] = currentProcedure,
                ["USER NAME"] = Environment.UserName,
                ["MACHINE NAME"] = Environment.MachineName,
                ["FILE NAME"] = currentFileName,
                ["DESCRIPTION"] = errorDescription,
            }.Select(x => $"[{x.Key}]=|{x.Value}|"));
            log.Error(logMessage);

            var userMessage = new StringBuilder()
                .AppendLine("Contact your system administrator. A record has been created in the log file.")
                .AppendLine("Procedure: " + currentProcedure)
                .AppendLine("Description: " + errorDescription)
                .ToString();

            if (isSilent == false)
            {
                MessageBox.Show(userMessage, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static bool IsActiveDocument(bool showMsg = false)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
                {
                    if (showMsg == true)
                    {
                        MessageBox.Show("The command could not be completed.  Please open a document and select a range."
                            , AssemblyInfo.Description
                            , MessageBoxButtons.OK
                            , MessageBoxIcon.Exclamation);
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

        public static bool IsActiveSelection(bool showMsg = false)
        {
            Excel.Range checkRange = null;
            try
            {
                checkRange = Globals.ThisAddIn.Application.Selection as Excel.Range; 
                if (null == checkRange)
                {
                    if (showMsg == true)
                    {
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [Range]"
                            , AssemblyInfo.Description
                            , MessageBoxButtons.OK
                            , MessageBoxIcon.Exclamation);
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
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [ListObject]"
                            , AssemblyInfo.Description
                            , MessageBoxButtons.OK
                            , MessageBoxIcon.Exclamation);
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
                    MessageBox.Show("The procedure can not run while you are editing a cell.", "No action taken."
                        , MessageBoxButtons.OK
                        , MessageBoxIcon.Information);
                }
                flag = true;
            }
            return flag;
        }

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

        public static bool IsTime(object expression)
        {
            try
            {
                string timeValue = Convert.ToString(expression);
                TimeSpan time1;
                return TimeSpan.TryParse(timeValue, out time1);
            }
            catch (Exception)
            {
                return false;
            }
        }

    }
}