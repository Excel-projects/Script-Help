namespace ScriptHelp.Scripts.Syntax
{
    public enum ExpressionType
    {

        None = 0,
        Identifier, 
        Operator,
        Number,
        Whitespace,
        Newline,
        Keyword,
        Comment,
        CommentLine,
        String,
        DelimitedGroup,     
        WordGroup           
    }
}
