using System;

namespace ScriptHelp.Scripts.Syntax
{
    internal class StyleGroupPair
    {
        public int Index { get; set; }
        public SyntaxStyle SyntaxStyle { get; set; }
        public string GroupName { get; set; }

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
