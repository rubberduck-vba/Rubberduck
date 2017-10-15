namespace Rubberduck.Parsing.Preprocessing
{
    public abstract class Expression : IExpression
    {
        public abstract IValue Evaluate();

        public bool EvaluateCondition()
        {
            var val = Evaluate();
            if (val == null)
            {
                return false;
            }
            return val.AsBool;
        }
    }
}
