using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueVisitorFactory
    {
        IParseTreeValueVisitor Create(RubberduckParserState state, IParseTreeValueFactory valueFactory);
    }

    public class ParseTreeValueVisitorFactory : IParseTreeValueVisitorFactory
    {
        public IParseTreeValueVisitor Create(RubberduckParserState state, IParseTreeValueFactory valueFactory)
        {
            return new ParseTreeValueVisitor(state, valueFactory);
        }
    }
}
