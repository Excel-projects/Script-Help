using System;
using System.Drawing;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// ColorUtils
    /// </summary>
    public class ColorUtils
    {
        /// <summary>
        /// ColorToRtfTableEntry
        /// </summary>
        /// <param name="color">this is the system drawing color</param>
        /// <returns>returns the string format of the color</returns>
        public static string ColorToRtfTableEntry(Color color)
        {
            return String.Format(@"\red{0}\green{1}\blue{2}", color.R, color.G, color.B);
        }
    }
}
