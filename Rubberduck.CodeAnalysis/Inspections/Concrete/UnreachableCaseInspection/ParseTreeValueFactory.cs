using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueFactory
    {
        IParseTreeValue CreateMismatchExpression(string expression, string typeName);
        IParseTreeValue CreateExpression(string expression, string typeName);
        IParseTreeValue CreateDeclaredType(string expression, string typeName);
        IParseTreeValue CreateValueType(string expression, string typeName);
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
        public IParseTreeValue CreateValueType(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateValueType(expression, declaredTypeName);
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
                return ParseTreeValue.CreateValueType(result.value, result.derivedType);
            }
            return ParseTreeValue.CreateExpression(valueToken, string.Empty);
        }

        public IParseTreeValue CreateDeclaredType(string expression, string declaredTypeName)
        {
            if (TokenTypeResolver.TryConformTokenToType(expression, declaredTypeName, out string conformedText))
            {
                return ParseTreeValue.CreateValueType(conformedText, declaredTypeName);
            }
            return ParseTreeValue.CreateExpression(expression, declaredTypeName);
        }

        public IParseTreeValue CreateMismatchExpression(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateMismatchExpression(expression, declaredTypeName);
        }

        public IParseTreeValue CreateExpression(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateExpression(expression, declaredTypeName);
        }

        public IParseTreeValue Create(byte value)
        {
            return CreateValueType(value.ToString(), Tokens.Byte);
        }

        public IParseTreeValue Create(int value)
        {
            return CreateValueType(value.ToString(), Tokens.Integer);
        }

        public IParseTreeValue Create(long value)
        {
            return CreateValueType(value.ToString(), Tokens.Long);
        }

        public IParseTreeValue Create(float value)
        {
            return CreateValueType(value.ToString(), Tokens.Single);
        }

        public IParseTreeValue Create(double value)
        {
            return CreateValueType(value.ToString(), Tokens.Double);
        }

        public IParseTreeValue Create(decimal value)
        {
            return CreateValueType(value.ToString(), Tokens.Currency);
        }

        public IParseTreeValue Create(bool value)
        {
            return CreateValueType(value ? Tokens.True : Tokens.False, Tokens.Boolean);
        }

        public IParseTreeValue CreateDate(double value)
        {
            return CreateDate(value.ToString());
        }

        public IParseTreeValue CreateDate(string value)
        {
            return new ParseTreeValue(value, Tokens.Date);
        }

        private bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");
    }
}
