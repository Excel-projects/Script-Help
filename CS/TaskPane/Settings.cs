using System;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;

namespace ScriptHelp.TaskPane
{
    /// <summary>
    /// Settings TaskPane
    /// </summary>
    public partial class Settings : UserControl
    {
        /// <summary>
        /// Initialize the controls in the object
        /// </summary>
        public Settings()
        {
            InitializeComponent();
            this.pgdSettings.SelectedObject = Properties.Settings.Default;
        }

        /// <summary> 
        /// Sets the column width of a property grid 
        /// </summary>
        /// <param name="grid">Represents the property grid object. </param>
        /// <param name="width">Represents the width of the column. </param>
        /// <remarks></remarks>
        public static void SetLabelColumnWidth(PropertyGrid grid, int width)
        {
            if (grid == null)
                return;

            FieldInfo fi = grid.GetType().GetField("gridView", BindingFlags.Instance | BindingFlags.NonPublic);
            if (fi == null)
                return;

            Control view = fi.GetValue(grid) as Control;
            if (view == null)
                return;

            MethodInfo mi = view.GetType().GetMethod("MoveSplitterTo", BindingFlags.Instance | BindingFlags.NonPublic);
            if (mi == null)
                return;
            mi.Invoke(view, new object[] { width });
        }

        /// <summary>
        /// Update the ribbon with the changed settings values
        /// </summary>
        /// <param name="s">contains the sender of the event, so if you had one method bound to multiple controls, you can distinguish them.</param>
        /// <param name="e">refers to the event arguments for the used event, they usually come in the form of properties/functions/methods that get to be available on it.</param>
        private void pgdSettings_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            Properties.Settings.Default.Save();
            Scripts.Ribbon.ribbonref.InvalidateRibbon();
        }
        
    }
}
