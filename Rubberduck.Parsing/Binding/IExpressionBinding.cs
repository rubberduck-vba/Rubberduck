namespace Rubberduck.Parsing.Binding
{
    public interface IExpressionBinding
    {
        IBoundExpression Resolve();
    }
}
