using System.Collections.Generic;

namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// CaseInsensitivePatternDefinition
    /// </summary>
    public class CaseInsensitivePatternDefinition : PatternDefinition
    {
        /// <summary>
        /// CaseInsensitivePatternDefinition
        /// </summary>
        /// <param name="tokens">string values</param>
        public CaseInsensitivePatternDefinition(IEnumerable<string> tokens)
            : base(false, tokens)
        {
        }

        /// <summary>
        /// CaseInsensitivePatternDefinition
        /// </summary>
        /// <param name="tokens">string values</param>
        public CaseInsensitivePatternDefinition(params string[] tokens)
            : base(false, tokens)
        {
        }
    }
}
