using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public static class StatementContext
    {
        public static DeclarationType GetSearchDeclarationType(StatementResolutionContext statementContext)
        {
            switch (statementContext)
            {
                case StatementResolutionContext.LetStatement:
                    return DeclarationType.PropertyLet;
                case StatementResolutionContext.SetStatement:
                    return DeclarationType.PropertySet;
                default:
                    return DeclarationType.PropertyGet;
            }
        }
    }
}
