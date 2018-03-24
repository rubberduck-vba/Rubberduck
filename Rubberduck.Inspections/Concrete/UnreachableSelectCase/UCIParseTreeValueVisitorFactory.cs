using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
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
