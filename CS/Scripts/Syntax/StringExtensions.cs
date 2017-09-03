using System;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// StringExtensions
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// NormalizeLineBreaks
        /// </summary>
        /// <param name="instance">instance</param>
        /// <param name="preferredLineBreak">preferredLineBreak</param>
        /// <returns>returns the preferred way of doing line breaks</returns>
        public static string NormalizeLineBreaks(this string instance, string preferredLineBreak)
        {
            return instance.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", preferredLineBreak);
        }

        /// <summary>
        /// NormalizeLineBreaks
        /// </summary>
        /// <param name="instance">instance</param>
        /// <returns>returns the line break with Environment.NewLine</returns>
        public static string NormalizeLineBreaks(this string instance)
        {
            return NormalizeLineBreaks(instance, Environment.NewLine);
        }
    }
}
