namespace Rubberduck.Parsing.PreProcessing
{
    public abstract class Expression : IExpression
    {
        public abstract IValue Evaluate();

        public bool EvaluateCondition()
        {
            var val = Evaluate();
            return val != null && val.AsBool;
        }
    }
}
