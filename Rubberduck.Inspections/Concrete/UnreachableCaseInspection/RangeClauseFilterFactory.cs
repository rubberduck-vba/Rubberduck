using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IRangeClauseFilterFactory
    {
        IRangeClauseFilter Create(string typeNme, IParseTreeValueFactory valueFactory);
    }

    //The following MIN/MAX values relate to VBA types
    internal static class CompareExtents
    {
        public static long LONGMIN = Int32.MinValue; //- 2147483648;
        public static long LONGMAX = Int32.MaxValue; //2147483647
        public static long INTEGERMIN = Int16.MinValue; //- 32768;
        public static long INTEGERMAX = Int16.MaxValue; //32767
        public static long BYTEMIN = byte.MinValue;  //0
        public static long BYTEMAX = byte.MaxValue;    //255
        public static decimal CURRENCYMIN = -922337203685477.5808M;
        public static decimal CURRENCYMAX = 922337203685477.5807M;
        public static double SINGLEMIN = float.MinValue; // -3402823E38;
        public static double SINGLEMAX = float.MaxValue;  //3402823E38;
    }

    public class RangeClauseFilterFactory : IRangeClauseFilterFactory
    {
        public IRangeClauseFilter Create(string typeName, IParseTreeValueFactory valueFactory)
        {
            if (valueFactory is null)
            {
                throw new ArgumentNullException(nameof(valueFactory));
            }

            if (!(IntegralNumberExtents.Keys.Contains(typeName)
                || typeName.Equals(Tokens.Double)
                || typeName.Equals(Tokens.Single)
                || typeName.Equals(Tokens.Currency)
                || typeName.Equals(Tokens.Boolean)
                || typeName.Equals(Tokens.String)))
            {
                throw new ArgumentException($"Unsupported TypeName ({typeName})");
            }

            if (IntegralNumberExtents.Keys.Contains(typeName))
            {
                var integerTypeFilter = new RangeClauseFilter<long>(typeName, valueFactory, this, ParseTreeValue.TryConvertValue);
                var minExtent = valueFactory.Create(IntegralNumberExtents[typeName].MinValue.ToString(), typeName);
                var maxExtent = valueFactory.Create(IntegralNumberExtents[typeName].MaxValue.ToString(), typeName);
                integerTypeFilter.AddExtents(minExtent, maxExtent);
                return integerTypeFilter;
            }

            if (typeName.Equals(Tokens.Double) || typeName.Equals(Tokens.Single))
            {
                var doubleTypeFilter = new RangeClauseFilter<double>(typeName, valueFactory, this, ParseTreeValue.TryConvertValue);
                if (typeName.Equals(Tokens.Single))
                {
                    var minExtent = valueFactory.Create(CompareExtents.SINGLEMIN.ToString(CultureInfo.InvariantCulture), typeName);
                    var maxExtent = valueFactory.Create(CompareExtents.SINGLEMAX.ToString(CultureInfo.InvariantCulture), typeName);
                    doubleTypeFilter.AddExtents(minExtent, maxExtent);
                }
                return doubleTypeFilter;
            }

            if (typeName.Equals(Tokens.Boolean))
            {
                var boolTypeFilter = new RangeClauseFilter<bool>(typeName, valueFactory, this, ParseTreeValue.TryConvertValue);
                return boolTypeFilter;
            }

            if (typeName.Equals(Tokens.Currency))
            {
                var decimalTypeFilter = new RangeClauseFilter<decimal>(typeName, valueFactory, this, ParseTreeValue.TryConvertValue);
                var minExtent = valueFactory.Create(CompareExtents.CURRENCYMIN.ToString(CultureInfo.InvariantCulture), typeName);
                var maxExtent = valueFactory.Create(CompareExtents.CURRENCYMAX.ToString(CultureInfo.InvariantCulture), typeName);
                decimalTypeFilter.AddExtents(minExtent, maxExtent);
                return decimalTypeFilter;
            }

            var filter = new RangeClauseFilter<string>(typeName, valueFactory, this, ParseTreeValue.TryConvertValue);
            return filter;
        }

        internal static Dictionary<string, (long MinValue, long MaxValue)> IntegralNumberExtents = new Dictionary<string, (long MinValue, long MaxValue)>()
        {
            [Tokens.Long] = (CompareExtents.LONGMIN, CompareExtents.LONGMAX),
            [Tokens.Integer] = (CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Int] = (CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Byte] = (CompareExtents.BYTEMIN, CompareExtents.BYTEMAX)
        };
    }
}
