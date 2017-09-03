using System;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// Expression
    /// </summary>
    public class Expression
    {
        /// <summary>
        /// ExpressionType
        /// </summary>
        public ExpressionType Type { get; private set; }
        /// <summary>
        /// Content
        /// </summary>
        public string Content { get; private set; }
        /// <summary>
        /// Group
        /// </summary>
        public string Group { get; private set; }

        /// <summary>
        /// Expression
        /// </summary>
        /// <param name="content">content</param>
        /// <param name="type">type</param>
        /// <param name="group">group</param>
        public Expression(string content, ExpressionType type, string group)
        {
            if (content == null)
                throw new ArgumentNullException("content");
            if (group == null)
                throw new ArgumentNullException("group");

            Type = type;
            Content = content;
            Group = group;
        }

        /// <summary>
        /// Expression
        /// </summary>
        /// <param name="content">content</param>
        /// <param name="type">type</param>
        public Expression(string content, ExpressionType type)
            : this(content, type, String.Empty)
        {
        }

        /// <summary>
        /// To override the system ToString function
        /// </summary>
        /// <returns>returns the overrided string</returns>
        public override string ToString()
        {
            if (Type == ExpressionType.Newline)
                return String.Format("({0})", Type);

            return String.Format("({0} --> {1}{2})", Content, Type, Group.Length > 0 ? " --> " + Group : String.Empty);
        }
    }
}
