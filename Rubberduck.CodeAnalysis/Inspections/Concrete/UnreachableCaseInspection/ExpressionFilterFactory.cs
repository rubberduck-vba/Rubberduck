using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection{

    public static class ExpressionFilterFactory
    {
        //The following MIN/MAX values relate to VBA types
        private static class CompareExtents
        {
            public static long LONGMIN = Int32.MinValue; //- 2147486648;
            public static long LONGMAX = Int32.MaxValue; //2147486647
            public static long INTEGERMIN = Int16.MinValue; //- 32768;
            public static long INTEGERMAX = Int16.MaxValue; //32767
            public static long BYTEMIN = byte.MinValue;  //0
            public static long BYTEMAX = byte.MaxValue;    //255
            public static decimal CURRENCYMIN = -922337203685477.5808M;
            public static decimal CURRENCYMAX = 922337203685477.5807M;
            public static double SINGLEMIN = float.MinValue; // -3402823E38;
            public static double SINGLEMAX = float.MaxValue;  //3402823E38;
        }

        private static Dictionary<string, (long typeMin, long typeMax)> IntegralNumberExtents = new Dictionary<string, (long typeMin, long typeMax)>()
        {
            [Tokens.Long] = (CompareExtents.LONGMIN, CompareExtents.LONGMAX),
            [Tokens.LongLong] = (CompareExtents.LONGMIN, CompareExtents.LONGMAX),
            [Tokens.Integer] = (CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Int] = (CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Byte] = (CompareExtents.BYTEMIN, CompareExtents.BYTEMAX)
        };

        public static IExpressionFilter Create(string typeName)
        {
            if (IntegralNumberExtents.Keys.Contains(typeName))
            {
                var integralNumberFilter = new ExpressionFilterIntegral(StringValueConverter.TryConvertString);
                integralNumberFilter.SetExtents(IntegralNumberExtents[typeName].typeMin, IntegralNumberExtents[typeName].typeMax);
                return integralNumberFilter;
            }
            else if (typeName.Equals(Tokens.Double) || typeName.Equals(Tokens.Single))
            {
                var rationalNumberFilter = new ExpressionFilter<double>(StringValueConverter.TryConvertString, typeName);
                if (typeName.Equals(Tokens.Single))
                {
                    rationalNumberFilter.SetExtents(CompareExtents.SINGLEMIN, CompareExtents.SINGLEMAX);
                }
                return rationalNumberFilter;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var decimalTypeFilter = new ExpressionFilter<decimal>(StringValueConverter.TryConvertString, typeName);
                decimalTypeFilter.SetExtents(CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX);
                return decimalTypeFilter;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                return new ExpressionFilterBoolean(StringValueConverter.TryConvertString);
            }

            return new ExpressionFilter<string>(StringValueConverter.TryConvertString, typeName);
        }
    }
}
