namespace Rubberduck.Parsing.PreProcessing
{
    public interface IExpression
    {
        IValue Evaluate();
        bool EvaluateCondition();
    }
}
