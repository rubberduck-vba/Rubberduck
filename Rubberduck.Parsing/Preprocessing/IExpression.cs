namespace Rubberduck.Parsing.Preprocessing
{
    public interface IExpression
    {
        IValue Evaluate();
        bool EvaluateCondition();
    }
}
