using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIParseTreeValueVisitorFactory
    {
        IUCIParseTreeValueVisitor Create(RubberduckParserState state, IUCIValueFactory valueFactory);
    }

    public class UCIParseTreeValueVisitorFactory : IUCIParseTreeValueVisitorFactory
    {
        public IUCIParseTreeValueVisitor Create(RubberduckParserState state, IUCIValueFactory valueFactory)
        {
            return new UCIParseTreeValueVisitor(state, valueFactory);
        }
    }
}
