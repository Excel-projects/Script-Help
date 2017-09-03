using System;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// StyleGroupPair
    /// </summary>
    internal class StyleGroupPair
    {
        /// <summary>
        /// Index
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// SyntaxStyle
        /// </summary>
        public SyntaxStyle SyntaxStyle { get; set; }
        /// <summary>
        /// GroupName
        /// </summary>
        public string GroupName { get; set; }

        /// <summary>
        /// StyleGroupPair
        /// </summary>
        /// <param name="syntaxStyle">syntaxStyle</param>
        /// <param name="groupName">groupName</param>
        public StyleGroupPair(SyntaxStyle syntaxStyle, string groupName)
        {
            if (syntaxStyle == null)
                throw new ArgumentNullException("syntaxStyle");
            if (groupName == null)
                throw new ArgumentNullException("groupName");

            SyntaxStyle = syntaxStyle;
            GroupName = groupName;
        }
    }
}
