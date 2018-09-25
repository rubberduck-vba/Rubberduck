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

        public static IExpressionFilter Create(string valueType)
        {
            if (IntegralNumberExtents.Keys.Contains(valueType))
            {
                var integralNumberFilter = new ExpressionFilterIntegral(valueType, long.Parse);
                integralNumberFilter.SetExtents(IntegralNumberExtents[valueType].typeMin, IntegralNumberExtents[valueType].typeMax);
                return integralNumberFilter;
            }
            else if (valueType.Equals(Tokens.Double) || valueType.Equals(Tokens.Single))
            {
                var floatingPointNumberFilter = new ExpressionFilter<double>(valueType, double.Parse);
                if (valueType.Equals(Tokens.Single))
                {
                    floatingPointNumberFilter.SetExtents(float.MinValue, float.MaxValue);
                }
                return floatingPointNumberFilter;
            }
            else if (valueType.Equals(Tokens.Currency))
            {
                var fixedPointNumberFilter = new ExpressionFilter<decimal>(valueType, VBACurrency.Parse);
                fixedPointNumberFilter.SetExtents(VBACurrency.MinValue, VBACurrency.MaxValue);
                return fixedPointNumberFilter;
            }
            else if (valueType.Equals(Tokens.Boolean))
            {
                return new ExpressionFilterBoolean();
            }

            else if (valueType.Equals(Tokens.Date))
            {
                return new ExpressionFilterDate();
            }

            return new ExpressionFilter<string>(valueType, (a) => { return a; });
        }
    }
}
