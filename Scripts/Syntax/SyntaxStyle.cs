using System.Drawing;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// SyntaxStyle
    /// </summary>
    public class SyntaxStyle
    {
        /// <summary>
        /// Bold
        /// </summary>
        public bool Bold { get; set; }
        /// <summary>
        /// Italic
        /// </summary>
        public bool Italic { get; set; }
        /// <summary>
        /// Color
        /// </summary>
        public Color Color { get; set; }

        /// <summary>
        /// SyntaxStyle
        /// </summary>
        /// <param name="color">color</param>
        /// <param name="bold">bold</param>
        /// <param name="italic">italic</param>
        public SyntaxStyle(Color color, bool bold, bool italic)
        {
            Color = color;
            Bold = bold;
            Italic = italic;
        }

        /// <summary>
        /// SyntaxStyle
        /// </summary>
        /// <param name="color">color</param>
        public SyntaxStyle(Color color)
            : this(color, false, false)
        {
        }
    }
}
