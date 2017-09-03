using System;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// PatternStyleMap
    /// </summary>
    internal class PatternStyleMap
    {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// PatternDefinition
        /// </summary>
        public PatternDefinition PatternDefinition { get; set; }
        /// <summary>
        /// SyntaxStyle
        /// </summary>
        public SyntaxStyle SyntaxStyle { get; set; }

        /// <summary>
        /// PatternStyleMap
        /// </summary>
        /// <param name="name">name</param>
        /// <param name="patternDefinition">patternDefinition</param>
        /// <param name="syntaxStyle">syntaxStyle</param>
        public PatternStyleMap(string name, PatternDefinition patternDefinition, SyntaxStyle syntaxStyle)
        {
            if (patternDefinition == null)
                throw new ArgumentNullException("patternDefinition");
            if (syntaxStyle == null)
                throw new ArgumentNullException("syntaxStyle");
            if (String.IsNullOrEmpty(name))
                throw new ArgumentException("name must not be null or empty", "name");

            Name = name;
            PatternDefinition = patternDefinition;
            SyntaxStyle = syntaxStyle;
        }
    }
}