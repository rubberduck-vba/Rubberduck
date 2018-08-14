using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection{

    public static class ExpressionFilterFactory
    {
        private static decimal CURRENCYMIN = -922337203685477.5808M;
        private static decimal CURRENCYMAX = 922337203685477.5807M;

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
                var integralNumberFilter = new ExpressionFilterIntegral(TokenParser.TryParse);
                integralNumberFilter.SetExtents(IntegralNumberExtents[typeName].typeMin, IntegralNumberExtents[typeName].typeMax);
                return integralNumberFilter;
            }
            else if (typeName.Equals(Tokens.Double) || typeName.Equals(Tokens.Single))
            {
                var floatingPointNumberFilter = new ExpressionFilter<double>(TokenParser.TryParse, typeName);
                if (typeName.Equals(Tokens.Single))
                {
                    floatingPointNumberFilter.SetExtents(float.MinValue, float.MaxValue);
                }
                return floatingPointNumberFilter;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var fixedPointNumberFilter = new ExpressionFilter<decimal>(TokenParser.TryParse, typeName);
                fixedPointNumberFilter.SetExtents(CURRENCYMIN, CURRENCYMAX);
                return fixedPointNumberFilter;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                return new ExpressionFilterBoolean(TokenParser.TryParse);
            }

            else if (typeName.Equals(Tokens.Date))
            {
                return new ExpressionFilterDate(TokenParser.TryParse);
            }

            return new ExpressionFilter<string>(TokenParser.TryParse, typeName);
        }
    }
}
