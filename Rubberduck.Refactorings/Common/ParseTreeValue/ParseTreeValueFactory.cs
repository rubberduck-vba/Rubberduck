using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings;

namespace Rubberduck.Refactoring.ParseTreeValue
{
    public class ParseTreeValueFactory : IParseTreeValueFactory
    {
        public IParseTreeValue CreateValueType(string expression, string declaredTypeName)
        {
            if (expression is null || declaredTypeName is null)
            {
                throw new ArgumentNullException();
            }
            return ParseTreeValue.CreateValueType(new TypeTokenPair(declaredTypeName, expression));
        }

        private static bool HasTypeHint(string token, out string valueTypeHint)
        {
            valueTypeHint = string.Empty;
            return !(token.First() == token.Last() && token.First() == '#') && SymbolList.TypeHintToTypeName.TryGetValue(token.Last().ToString(), out valueTypeHint);
        }

        public IParseTreeValue Create(string valueToken)
        {
            if (valueToken is null || valueToken.Equals(string.Empty))
            {
                throw new ArgumentException();
            }

            if (HasTypeHint(valueToken, out string valueType))
            {
                var vToken = valueToken.Remove(valueToken.Length - 1);
                var conformedTypeTokenPair = TypeTokenPair.ConformToType(valueType, vToken);
                if (conformedTypeTokenPair.HasValue)
                {
                    return ParseTreeValue.CreateValueType(conformedTypeTokenPair);
                }
                return ParseTreeValue.CreateExpression(new TypeTokenPair(valueType, vToken));
            }

            if (TypeTokenPair.TryParse(valueToken, out TypeTokenPair result))
            {
                return ParseTreeValue.CreateValueType(result);
            }
            return ParseTreeValue.CreateExpression(new TypeTokenPair(null, valueToken));
        }

        public IParseTreeValue CreateDeclaredType(string expression, string declaredTypeName)
        {
            if (expression is null || declaredTypeName is null)
            {
                throw new ArgumentNullException();
            }

            if (ParseTreeValue.TryGetNonPrintingControlCharCompareToken(expression, out string comparableToken))
            {
                var charConversion = new TypeTokenPair(Tokens.String, comparableToken);
                return ParseTreeValue.CreateValueType(charConversion);
            }

            var goalTypeTokenPair = new TypeTokenPair(declaredTypeName, null);
            var typeToken = TypeTokenPair.ConformToType(declaredTypeName, expression);
            if (typeToken.HasValue)
            {
                if (LetCoerce.ExceedsValueTypeRange(typeToken.ValueType, typeToken.Token))
                {
                    return ParseTreeValue.CreateOverflowExpression(expression, declaredTypeName);
                }
                return ParseTreeValue.CreateValueType(typeToken);
            }
            return ParseTreeValue.CreateExpression(new TypeTokenPair(declaredTypeName, expression));
        }

        public IParseTreeValue CreateMismatchExpression(string expression, string declaredTypeName)
        {
            if (expression is null || declaredTypeName is null)
            {
                throw new ArgumentNullException();
            }
            return ParseTreeValue.CreateMismatchExpression(expression, declaredTypeName);
        }

        public IParseTreeValue CreateExpression(string expression, string declaredTypeName)
        {
            if (expression is null || declaredTypeName is null)
            {
                throw new ArgumentNullException();
            }
            return ParseTreeValue.CreateExpression(new TypeTokenPair(declaredTypeName, expression));
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
            return new ParseTreeValue(new TypeTokenPair(Tokens.Date, value));
        }
    }
}
