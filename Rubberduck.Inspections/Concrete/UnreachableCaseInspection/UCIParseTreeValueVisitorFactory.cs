using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIParseTreeValueVisitorFactory
    {
        IUCIParseTreeValueVisitor Create(RubberduckParserState state);
    }

    public class UCIParseTreeValueVisitorFactory : IUCIParseTreeValueVisitorFactory
    {
        public IUCIParseTreeValueVisitor Create(RubberduckParserState state)
        {
            return new UCIParseTreeValueVisitor(state, new UCIValueFactory());
        }
    }
}
