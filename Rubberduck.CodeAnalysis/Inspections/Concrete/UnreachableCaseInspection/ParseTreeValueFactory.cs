using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueFactory
    {
        IParseTreeValue CreateExpression(string expression, string typeName);
        IParseTreeValue CreateDeclaredType(string expression, string typeName);
        IParseTreeValue CreateConstant(string expression, string typeName);
        IParseTreeValue Create(string valueToken);
        IParseTreeValue Create(byte value);
        IParseTreeValue Create(int value);
        IParseTreeValue Create(long value);
        IParseTreeValue Create(float value);
        IParseTreeValue Create(double value);
        IParseTreeValue Create(decimal value);
        IParseTreeValue Create(bool value);
        IParseTreeValue CreateDate(string value);
        IParseTreeValue CreateDate(double value);
    }

    public class ParseTreeValueFactory : IParseTreeValueFactory
    {
        public IParseTreeValue CreateConstant(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateConstant(expression, declaredTypeName);
        }

        public IParseTreeValue Create(string valueToken)
        {
            if (TokenTypeResolver.TryDeriveTypeName(valueToken, out (string derivedType, string value) result, out bool derivedFromTypeHint))
            {
                if (derivedFromTypeHint && result.derivedType.Equals(Tokens.String))
                {
                    return ParseTreeValue.CreateExpression(result.value, Tokens.String);
                }
                if (result.derivedType.Equals(Tokens.Date))
                {
                    return CreateDate(valueToken);
                }
                return ParseTreeValue.CreateConstant(result.value, result.derivedType);
            }
            return ParseTreeValue.CreateExpression(valueToken, string.Empty);
        }

        public IParseTreeValue CreateDeclaredType(string expression, string declaredTypeName)
        {
            if (TokenTypeResolver.TryConformTokenToType(expression, declaredTypeName, out string conformedText))
            {
                return ParseTreeValue.CreateConstant(conformedText, declaredTypeName);
            }
            return ParseTreeValue.CreateExpression(expression, declaredTypeName);
        }

        public IParseTreeValue CreateExpression(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateExpression(expression, declaredTypeName);
        }

        public IParseTreeValue Create(byte value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Byte);
        }

        public IParseTreeValue Create(int value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Integer);
        }

        public IParseTreeValue Create(long value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Long);
        }

        public IParseTreeValue Create(float value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Single);
        }

        public IParseTreeValue Create(double value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Double);
        }

        public IParseTreeValue Create(decimal value)
        {
            return ParseTreeValue.CreateConstant(value.ToString(), Tokens.Currency);
        }

        public IParseTreeValue Create(bool value)
        {
            return ParseTreeValue.CreateConstant(value ? Tokens.True : Tokens.False, Tokens.Boolean);
        }

        public IParseTreeValue CreateDate(double value)
        {
            return new ParseTreeValue(value.ToString(), Tokens.Date);
        }

        public IParseTreeValue CreateDate(string value)
        {
            return new ParseTreeValue(value, Tokens.Date);
        }

        private bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");
    }
}
