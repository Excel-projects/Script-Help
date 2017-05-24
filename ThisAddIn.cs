using System;
using System.Linq;
using ScriptHelp.Scripts;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ScriptHelp
{
    /// <summary> 
    /// Used to handle startup and shutdown of the addin
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary> 
        /// This method is used to handle events in the ribbon
        /// </summary>
        public static Excel.Application e_application;
        /// <summary> 
        /// This method is used to handle events in the ribbon
        /// </summary>
        public static Office.IRibbonUI e_ribbon;

        /// <summary>
        /// This method is triggered on the startup of the addin
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            e_application = this.Application;
            e_application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(e_Application_SheetSelectionChange);
        }

        /// <summary> 
        /// This method is triggered after the sheet selection change event
        /// </summary>
        /// <param name="sh">the name of the sheet </param>
        /// <param name="target">the currently selected range </param>
        /// <remarks></remarks>
        private void e_Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            e_ribbon.Invalidate();
        }

        /// <summary>
        /// This method is triggered on the shutdown of the addin
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            e_application.SheetSelectionChange -= new Excel.AppEvents_SheetSelectionChangeEventHandler(e_Application_SheetSelectionChange);
            e_application = null;
        }

        /// <summary> 
        /// This method is used to create the ribbon
        /// </summary>
        /// <returns></returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
