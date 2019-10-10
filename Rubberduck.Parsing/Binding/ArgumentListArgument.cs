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
        private readonly Func<Declaration, IBoundExpression> _namedArgumentExpressionCreator;
        private readonly bool _isAddressOfArgument;

        public ArgumentListArgument(IExpressionBinding binding, ParserRuleContext context, ArgumentListArgumentType argumentType, bool isAddressOfArgument = false)
            : this (binding, context, argumentType, calledProcedure => null, isAddressOfArgument)
        {
        }

        public ArgumentListArgument(IExpressionBinding binding, ParserRuleContext context, ArgumentListArgumentType argumentType, Func<Declaration, IBoundExpression> namedArgumentExpressionCreator, bool isAddressOfArgument = false)
        {
            _binding = binding;
            Context = context;
            ArgumentType = argumentType;
            _namedArgumentExpressionCreator = namedArgumentExpressionCreator;
            _isAddressOfArgument = isAddressOfArgument;
            ReferencedParameter = null;
        }

        public ArgumentListArgumentType ArgumentType { get; }
        public IBoundExpression NamedArgumentExpression { get; private set; }
        public IBoundExpression Expression { get; private set; }
        public ParameterDeclaration ReferencedParameter { get; private set; }
        public ParserRuleContext Context { get; }

        public void Resolve(Declaration calledProcedure, int parameterIndex, bool isArrayAccess = false)
        {
            var binding = _binding;
            if (calledProcedure != null)
            {
                NamedArgumentExpression = _namedArgumentExpressionCreator(calledProcedure);
                ReferencedParameter = ResolveReferencedParameter(calledProcedure, parameterIndex);

                if (!_isAddressOfArgument 
                    && !(Context is VBAParser.MissingArgumentContext)
                    && (isArrayAccess 
                        ||  ReferencedParameter != null 
                            && !CanBeObject(ReferencedParameter)))
                {
                    binding = new LetCoercionDefaultBinding(Context, binding);
                }
            }

            Expression = binding.Resolve();
        }

        private ParameterDeclaration ResolveReferencedParameter(Declaration calledProcedure, int parameterIndex)
        {
            if (NamedArgumentExpression != null)
            {
                return NamedArgumentExpression.ReferencedDeclaration as ParameterDeclaration;
            }

            if (parameterIndex >= 0 && calledProcedure is IParameterizedDeclaration parameterizedDeclaration)
            {
                var parameters = parameterizedDeclaration.Parameters.ToList();
                if (parameterIndex >= parameters.Count)
                {
                    return parameters.FirstOrDefault(param => param.IsParamArray);
                }

                return parameters[parameterIndex];
            }

            return null;
        }

        private bool CanBeObject(ParameterDeclaration parameter)
        {
            return parameter.IsObject
                   || Tokens.Variant.Equals(parameter.AsTypeName, StringComparison.InvariantCultureIgnoreCase)
                   && (!parameter.IsArray || parameter.IsParamArray);
        }
    }
}
