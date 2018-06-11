namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueFactory
    {
        IParseTreeValue Create(string valueToken, string declaredTypeName = null, string conformToTypeName = null);
    }

    public class ParseTreeValueFactory : IParseTreeValueFactory
    {
        public IParseTreeValue Create(string valueToken, string declaredTypeName = null, string conformToTypeName = null)
        {
            return new ParseTreeValue(valueToken, declaredTypeName, conformToTypeName);
        }
    }
}
