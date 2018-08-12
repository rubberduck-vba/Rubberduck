using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PreProcessing;
using System;
using System.Globalization;

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
                return ParseTreeValue.CreateConstant(result.value, result.derivedType);
            }
            return ParseTreeValue.CreateExpression(valueToken, string.Empty);
        }

        public IParseTreeValue CreateDeclaredType(string expression, string declaredTypeName)
        {
            return new ParseTreeValue(expression, declaredTypeName);
        }

        public IParseTreeValue CreateExpression(string expression, string declaredTypeName)
        {
            return ParseTreeValue.CreateExpression(expression, declaredTypeName);
        }

        public IParseTreeValue Create(byte value)
        {
            return CreateConstant(value.ToString(), Tokens.Byte);
        }

        public IParseTreeValue Create(int value)
        {
            return CreateConstant(value.ToString(), Tokens.Integer);
        }

        public IParseTreeValue Create(long value)
        {
            return CreateConstant(value.ToString(), Tokens.Long);
        }

        public IParseTreeValue Create(float value)
        {
            return CreateConstant(value.ToString(), Tokens.Single);
        }

        public IParseTreeValue Create(double value)
        {
            return CreateConstant(value.ToString(), Tokens.Double);
        }

        public IParseTreeValue Create(decimal value)
        {
            return CreateConstant(value.ToString(), Tokens.Currency);
        }

        public IParseTreeValue Create(bool value)
        {
            return CreateConstant(value ? Tokens.True : Tokens.False, Tokens.Boolean);
        }

        public IParseTreeValue CreateDate(double value)
        {
            var dv = new DateValue(DateTime.FromOADate(value));
            var cdv = new ComparableDateValue(dv);
            return CreateConstant(AnnotateAsDateLiteral(cdv.AsDate.ToString(CultureInfo.InvariantCulture)), Tokens.Date);
        }
        public IParseTreeValue CreateDate(string value)
        {
            return CreateConstant(value, Tokens.Date);
        }

        private bool IsStringConstant(string input) => input.StartsWith("\"") && input.EndsWith("\"");

        private static string AnnotateAsDateLiteral(string input)
        {
            var result = input;
            if (!input.StartsWith("#"))
            {
                result = $"#{result}";
            }
            if (!input.EndsWith("#"))
            {
                result = $"{result}#";
            }
            result.Replace(" 00:00:00", "");
            return result;
        }
    }
}
