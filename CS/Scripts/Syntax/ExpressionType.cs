namespace ScriptHelp.Scripts.Syntax
{
    /// <summary>
    /// Enumerates the type of the parsed content
    /// </summary>
    public enum ExpressionType
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        /// <summary>
        /// i.e. a word which is neither keyword nor inside any word-group
        /// </summary>
        Identifier, 
        /// <summary>
        /// Operator
        /// </summary>
        Operator,
        /// <summary>
        /// Number
        /// </summary>
        Number,
        /// <summary>
        /// Whitespace
        /// </summary>
        Whitespace,
        /// <summary>
        /// Newline
        /// </summary>
        Newline,
        /// <summary>
        /// Keyword
        /// </summary>
        Keyword,
        /// <summary>
        /// Comment
        /// </summary>
        Comment,
        /// <summary>
        /// CommentLine
        /// </summary>
        CommentLine,
        /// <summary>
        /// String
        /// </summary>
        String,
        /// <summary>
        /// needs extra argument
        /// </summary>
        DelimitedGroup,     
        /// <summary>
        /// needs extra argument
        /// </summary>
        WordGroup           
    }
}
