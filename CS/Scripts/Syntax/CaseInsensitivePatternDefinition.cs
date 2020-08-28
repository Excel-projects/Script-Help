using System.Collections.Generic;

namespace ScriptHelp.Scripts.Syntax
{
    public class CaseInsensitivePatternDefinition : PatternDefinition
    {
        public CaseInsensitivePatternDefinition(IEnumerable<string> tokens)
            : base(false, tokens)
        {
        }

        public CaseInsensitivePatternDefinition(params string[] tokens)
            : base(false, tokens)
        {
        }
    }
}
