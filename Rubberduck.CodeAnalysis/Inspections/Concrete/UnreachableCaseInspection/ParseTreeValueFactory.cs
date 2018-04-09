
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueFactory
    {
        IParseTreeValue Create(string valueToken, string declaredTypeName = null);
    }

    public class ParseTreeValueFactory : IParseTreeValueFactory
    {
        public IParseTreeValue Create(string valueToken, string declaredTypeName = null)
        {
            return new ParseTreeValue(valueToken, declaredTypeName);
        }
    }
}
