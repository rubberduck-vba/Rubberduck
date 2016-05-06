using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public static class StatementContext
    {
        public static DeclarationType GetSearchDeclarationType(ResolutionStatementContext statementContext)
        {
            switch(statementContext)
            {
                case ResolutionStatementContext.LetStatement:
                    return DeclarationType.PropertyLet;
                case ResolutionStatementContext.SetStatement:
                    return DeclarationType.PropertySet;
                default:
                    return DeclarationType.PropertyGet;
            }
        }
    }
}
