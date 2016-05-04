using Rubberduck.Parsing.Symbols;
using System;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentListArgument
    {
        private readonly IExpressionBinding _binding;
        private IBoundExpression _expression;
        private IBoundExpression _namedArgumentExpression;
        private readonly ArgumentListArgumentType _argumentType;
        private readonly Func<Declaration, IBoundExpression> _namedArgumentExpressionCreator;

        public ArgumentListArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType)
            : this (binding, argumentType, calledProcedure => null)
        {
        }

        public ArgumentListArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType, Func<Declaration, IBoundExpression> namedArgumentExpressionCreator)
        {
            _binding = binding;
            _argumentType = argumentType;
            _namedArgumentExpressionCreator = namedArgumentExpressionCreator;
        }

        public ArgumentListArgumentType ArgumentType
        {
            get
            {
                return _argumentType;
            }
        }

        public IBoundExpression NamedArgumentExpression
        {
            get
            {
                return _namedArgumentExpression;
            }
        }

        public IBoundExpression Expression
        {
            get
            {
                return _expression;
            }
        }

        public void Resolve(Declaration calledProcedure)
        {
            _expression = _binding.Resolve();
            if (calledProcedure != null)
            {
                _namedArgumentExpression = _namedArgumentExpressionCreator(calledProcedure);
            }
        }
    }
}
