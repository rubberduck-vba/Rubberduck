namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentListArgument
    {
        private readonly IExpressionBinding _binding;
        private IBoundExpression _expression;
        private readonly ArgumentListArgumentType _argumentType;

        public ArgumentListArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType)
        {
            _binding = binding;
            _argumentType = argumentType;
        }

        public ArgumentListArgumentType ArgumentType
        {
            get
            {
                return _argumentType;
            }
        }

        public IBoundExpression Expression
        {
            get
            {
                return _expression;
            }
        }

        public void Resolve()
        {
            _expression = _binding.Resolve();
        }
    }
}
