
namespace Rubberduck.Refactorings
{
    public interface IParseTreeValue
    {
        string Token { get; }
        string ValueType { get; }
        bool ParsesToConstantValue { get; }
        bool IsOverflowExpression { get; }
        bool IsMismatchExpression { get; }
    }
}
