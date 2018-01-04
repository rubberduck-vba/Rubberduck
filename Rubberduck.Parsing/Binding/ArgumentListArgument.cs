using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentListArgument
    {
        private readonly IExpressionBinding _binding;
        private readonly Func<Declaration, IBoundExpression> _namedArgumentExpressionCreator;

        public ArgumentListArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType)
            : this (binding, argumentType, calledProcedure => null)
        {
        }

        public ArgumentListArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType, Func<Declaration, IBoundExpression> namedArgumentExpressionCreator)
        {
            _binding = binding;
            ArgumentType = argumentType;
            _namedArgumentExpressionCreator = namedArgumentExpressionCreator;
        }

        public ArgumentListArgumentType ArgumentType { get; }

        public IBoundExpression NamedArgumentExpression { get; private set; }

        public IBoundExpression Expression { get; private set; }

        public void Resolve(Declaration calledProcedure)
        {
            Expression = _binding.Resolve();
            if (calledProcedure != null)
            {
                NamedArgumentExpression = _namedArgumentExpressionCreator(calledProcedure);
            }
        }
    }
}
