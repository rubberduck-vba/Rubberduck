using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection{

    public static class ExpressionFilterFactory
    {
        private static Dictionary<string, (long typeMin, long typeMax)> IntegralNumberExtents = new Dictionary<string, (long typeMin, long typeMax)>()
        {
            [Tokens.LongLong] = (long.MinValue, long.MaxValue),
            [Tokens.Long] = (Int32.MinValue, Int32.MaxValue),
            [Tokens.Integer] = (Int16.MinValue, Int16.MaxValue),
            [Tokens.Int] = (Int16.MinValue, Int16.MaxValue),
            [Tokens.Byte] = (byte.MinValue, byte.MaxValue)
        };

        public static IExpressionFilter Create(string typeName)
        {
            if (IntegralNumberExtents.Keys.Contains(typeName))
            {
                var integralNumberFilter = new ExpressionFilterIntegral(TryCoerce);
                integralNumberFilter.SetExtents(IntegralNumberExtents[typeName].typeMin, IntegralNumberExtents[typeName].typeMax);
                return integralNumberFilter;
            }
            else if (typeName.Equals(Tokens.Double) || typeName.Equals(Tokens.Single))
            {
                var floatingPointNumberFilter = new ExpressionFilter<double>(TryCoerce, typeName);
                if (typeName.Equals(Tokens.Single))
                {
                    floatingPointNumberFilter.SetExtents(float.MinValue, float.MaxValue);
                }
                return floatingPointNumberFilter;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var fixedPointNumberFilter = new ExpressionFilter<decimal>(TryCoerce, typeName);
                fixedPointNumberFilter.SetExtents(VBACurrency.MinValue, VBACurrency.MaxValue);
                return fixedPointNumberFilter;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                return new ExpressionFilterBoolean(TryCoerce);
            }

            else if (typeName.Equals(Tokens.Date))
            {
                return new ExpressionFilterDate(TryCoerce);
            }

            return new ExpressionFilter<string>(TryCoerce, typeName);
        }

        private static bool TryCoerce(string value, out long result, string typeName = null)
            => TryCoerce(value, out result, long.Parse, typeName ?? Tokens.Long);

        private static bool TryCoerce(string value, out double result, string typeName = null)
            => TryCoerce(value, out result, double.Parse, typeName ?? Tokens.Double);

        private static bool TryCoerce(string value, out decimal result, string typeName = null)
            => TryCoerce(value, out result, decimal.Parse, typeName ?? Tokens.Currency);

        private static bool TryCoerce(string value, out bool result, string typeName = null)
            => TryCoerce(value, out result, bool.Parse, typeName ?? Tokens.Boolean);

        private static bool TryCoerce(string value, out ComparableDateValue result, string typeName = null)
            => TryCoerce(value, out result, ComparableDateValue.Parse, typeName ?? Tokens.Date);

        private static bool TryCoerce<T>(string value, out T result, Func<string,T> parser, string typeName)
        {
            result = default;
            if (LetCoercer.TryCoerceToken((Tokens.String, value), typeName, out string token))
            {
                result = parser(token);
                return true;
            }
            return false;
        }

        private static bool TryCoerce(string value, out string result, string typeName = null)
        {
            result = value;
            return true;
        }
    }
}
