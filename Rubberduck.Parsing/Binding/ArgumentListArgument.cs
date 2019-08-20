using Rubberduck.Parsing.Symbols;
using System;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentListArgument
    {
        private readonly IExpressionBinding _binding;
        private readonly ParserRuleContext _context;
        private readonly Func<Declaration, IBoundExpression> _namedArgumentExpressionCreator;
        private readonly bool _isAddressOfArgument;

        public ArgumentListArgument(IExpressionBinding binding, ParserRuleContext context, ArgumentListArgumentType argumentType, bool isAddressOfArgument = false)
            : this (binding, context, argumentType, calledProcedure => null, isAddressOfArgument)
        {
        }

        public ArgumentListArgument(IExpressionBinding binding, ParserRuleContext context, ArgumentListArgumentType argumentType, Func<Declaration, IBoundExpression> namedArgumentExpressionCreator, bool isAddressOfArgument = false)
        {
            _binding = binding;
            _context = context;
            ArgumentType = argumentType;
            _namedArgumentExpressionCreator = namedArgumentExpressionCreator;
            _isAddressOfArgument = isAddressOfArgument;
        }

        public ArgumentListArgumentType ArgumentType { get; }
        public IBoundExpression NamedArgumentExpression { get; private set; }
        public IBoundExpression Expression { get; private set; }

        public void Resolve(Declaration calledProcedure, int parameterIndex)
        {
            var binding = _binding;
            if (calledProcedure != null)
            {
                NamedArgumentExpression = _namedArgumentExpressionCreator(calledProcedure);

                if (!_isAddressOfArgument && !CanBeObject(calledProcedure, parameterIndex))
                {
                    binding = new LetCoercionDefaultBinding(_context, binding);
                }
            }

            Expression = binding.Resolve();
        }

        private bool CanBeObject(Declaration calledProcedure, int parameterIndex)
        {
            if (NamedArgumentExpression != null)
            {
                var correspondingParameter = NamedArgumentExpression.ReferencedDeclaration as ParameterDeclaration;
                return CanBeObject(correspondingParameter);
            }

            if (parameterIndex >= 0 && calledProcedure is IParameterizedDeclaration parameterizedDeclaration)
            {
                var parameters = parameterizedDeclaration.Parameters.ToList();
                if (parameterIndex >= parameters.Count)
                {
                    return parameters.Any(param => param.IsParamArray);
                }

                var correspondingParameter = parameters[parameterIndex];
                return CanBeObject(correspondingParameter);

            }

            return true;
        }

        private bool CanBeObject(ParameterDeclaration parameter)
        {
            return parameter.IsObject
                   || Tokens.Variant.Equals(parameter.AsTypeName, StringComparison.InvariantCultureIgnoreCase)
                   && (!parameter.IsArray || parameter.IsParamArray);
        }
    }
}
