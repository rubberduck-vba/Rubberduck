using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactoring.ParseTreeValue;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{ 
    internal static class ExpressionFilterFactory
    {
        private static readonly Dictionary<string, (long typeMin, long typeMax)> IntegralNumberExtents = new Dictionary<string, (long typeMin, long typeMax)>()
        {
            [Tokens.LongLong] = (long.MinValue, long.MaxValue),
            [Tokens.Long] = (int.MinValue, int.MaxValue),
            [Tokens.Integer] = (short.MinValue, short.MaxValue),
            [Tokens.Int] = (short.MinValue, short.MaxValue),
            [Tokens.Byte] = (byte.MinValue, byte.MaxValue)
        };

        public static IExpressionFilter Create(string valueType)
        {
            if (IntegralNumberExtents.Keys.Contains(valueType))
            {
                var integralNumberFilter = new ExpressionFilterIntegral(valueType, s => long.Parse(s, CultureInfo.InvariantCulture));
                integralNumberFilter.SetExtents(IntegralNumberExtents[valueType].typeMin, IntegralNumberExtents[valueType].typeMax);
                return integralNumberFilter;
            }
            if (valueType.Equals(Tokens.Double) || valueType.Equals(Tokens.Single))
            {
                var floatingPointNumberFilter = new ExpressionFilter<double>(valueType, s => double.Parse(s, CultureInfo.InvariantCulture));
                if (valueType.Equals(Tokens.Single))
                {
                    floatingPointNumberFilter.SetExtents(float.MinValue, float.MaxValue);
                }
                return floatingPointNumberFilter;
            }
            if (valueType.Equals(Tokens.Currency))
            {
                var fixedPointNumberFilter = new ExpressionFilter<decimal>(valueType, VBACurrency.Parse);
                fixedPointNumberFilter.SetExtents(VBACurrency.MinValue, VBACurrency.MaxValue);
                return fixedPointNumberFilter;
            }
            if (valueType.Equals(Tokens.Boolean))
            {
                return new ExpressionFilterBoolean();
            }

            if (valueType.Equals(Tokens.Date))
            {
                return new ExpressionFilterDate();
            }

            return new ExpressionFilter<string>(valueType, a => a);
        }
    }
}
