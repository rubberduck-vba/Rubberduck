
namespace Rubberduck.Refactorings
{
    public interface IParseTreeValueFactory
    {
        IParseTreeValue CreateMismatchExpression(string expression, string typeName);
        IParseTreeValue CreateExpression(string expression, string typeName);
        IParseTreeValue CreateDeclaredType(string expression, string typeName);
        IParseTreeValue CreateValueType(string expression, string typeName);
        IParseTreeValue Create(string valueToken);
        IParseTreeValue Create(byte value);
        IParseTreeValue Create(int value);
        IParseTreeValue Create(long value);
        IParseTreeValue Create(float value);
        IParseTreeValue Create(double value);
        IParseTreeValue Create(decimal value);
        IParseTreeValue Create(bool value);
        IParseTreeValue CreateDate(string value);
        IParseTreeValue CreateDate(double value);
    }
}
