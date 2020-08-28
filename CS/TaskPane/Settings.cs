using System;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using ScriptHelp.Scripts;

namespace ScriptHelp.TaskPane
{
    public partial class Settings : UserControl
    {
        public Settings()
        {
            try
            {
                InitializeComponent();
                this.pgdSettings.SelectedObject = Properties.Settings.Default;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public static void SetLabelColumnWidth(PropertyGrid grid, int width)
        {
            try
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
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        private void pgdSettings_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            try
            {
                Properties.Settings.Default.Save();
                Scripts.Ribbon.ribbonref.InvalidateRibbon();
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

    }
}
