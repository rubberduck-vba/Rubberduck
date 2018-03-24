using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUCIRangeClauseFilterFactory
    {
        IUCIRangeClauseFilter Create(string typeNme, IUCIValueFactory valueFactory, IUCIRangeClauseFilterFactory filterFactory);
    }

    internal static class CompareExtents
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

    public class UCIRangeClauseFilterFactory : IUCIRangeClauseFilterFactory
    {
        public IUCIRangeClauseFilter Create(string typeName, IUCIValueFactory valueFactory, IUCIRangeClauseFilterFactory filterFactory)
        {
            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                var filter = new UCIRangeClauseFilter<long>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertLong);
                var minExtent = valueFactory.Create(IntegerNumberExtents[typeName].Item1.ToString(), typeName);
                var maxExtent = valueFactory.Create(IntegerNumberExtents[typeName].Item2.ToString(), typeName);
                filter.AddExtents(minExtent, maxExtent);
                return filter;
            }
            else if (SingleDataTypeExtents.ContainsKey(typeName))
            {
                var filter = new UCIRangeClauseFilter<double>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertDouble);
                var minExtent = valueFactory.Create(SingleDataTypeExtents[typeName].Item1.ToString(), typeName);
                var maxExtent = valueFactory.Create(SingleDataTypeExtents[typeName].Item2.ToString(), typeName);
                filter.AddExtents(minExtent, maxExtent);
                return filter;
            }
            else if (typeName.Equals(Tokens.Double))
            {
                var filter = new UCIRangeClauseFilter<double>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertDouble);
                return filter;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var filter = new UCIRangeClauseFilter<bool>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertBoolean);
                return filter;
            }
            else if (typeName.Equals(Tokens.String))
            {
                var filter = new UCIRangeClauseFilter<string>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertString);
                return filter;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var filter = new UCIRangeClauseFilter<decimal>(typeName, valueFactory, filterFactory, UCIValueConverter.ConvertDecimal);
                var minExtent = valueFactory.Create(CompareExtents.CURRENCYMIN.ToString(), typeName);
                var maxExtent = valueFactory.Create(CompareExtents.CURRENCYMAX.ToString(), typeName);
                filter.AddExtents(minExtent, maxExtent);
                return filter;
            }
            throw new ArgumentException($"Unsupported TypeName ({typeName})");
        }

        public static Dictionary<string, Tuple<long, long>> IntegerNumberExtents = new Dictionary<string, Tuple<long, long>>()
        {
            [Tokens.Long] = new Tuple<long, long>(CompareExtents.LONGMIN, CompareExtents.LONGMAX),
            [Tokens.Integer] = new Tuple<long, long>(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Int] = new Tuple<long, long>(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Byte] = new Tuple<long, long>(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX)
        };

        public static Dictionary<string, Tuple<double, double>> SingleDataTypeExtents = new Dictionary<string, Tuple<double, double>>()
        {
            [Tokens.Single] = new Tuple<double, double>(CompareExtents.SINGLEMIN, CompareExtents.SINGLEMAX)
        };
    }
}
