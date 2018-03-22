using Antlr4.Runtime;
using Rubberduck.Inspections;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionSummaryClauseFactory
    {
        ISummaryCoverage Create(string typeName, IUnreachableCaseInspectionValueFactory valueFactory);
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

    public class UnreachableCaseInspectionSummaryClauseFactory2 : IUnreachableCaseInspectionSummaryClauseFactory
    {
        public ISummaryCoverage Create(string typeName, IUnreachableCaseInspectionValueFactory valueFactory)
        {
            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                var summary = new SummaryCoverage2<long>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertLong)
                {
                    TypeName = typeName
                };
                var minExtent = valueFactory.Create(IntegerNumberExtents[typeName].Item1.ToString(), typeName);
                var maxExtent = valueFactory.Create(IntegerNumberExtents[typeName].Item2.ToString(), typeName);
                summary.AddExtents(minExtent, maxExtent);
                return summary;
            }
            else if (SingleDataTypeExtents.ContainsKey(typeName))
            {
                var summary = new SummaryCoverage2<double>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertDouble)
                {
                    TypeName = typeName
                };
                var minExtent = valueFactory.Create(SingleDataTypeExtents[typeName].Item1.ToString(), typeName);
                var maxExtent = valueFactory.Create(SingleDataTypeExtents[typeName].Item2.ToString(), typeName);
                summary.AddExtents(minExtent, maxExtent);
                return summary;
            }
            else if (typeName.Equals(Tokens.Double))
            {
                var summary = new SummaryCoverage2<double>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertDouble)
                {
                    TypeName = typeName
                };
                return summary;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var summary = new SummaryCoverage2<bool>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertBoolean)
                {
                    TypeName = typeName
                };
                return summary;
            }
            else if (typeName.Equals(Tokens.String))
            {
                var summary = new SummaryCoverage2<string>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertString)
                {
                    TypeName = typeName
                };
                return summary;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var summary = new SummaryCoverage2<decimal>(new UnreachableCaseInspectionValueFactory(), UCIValueConverter.ConvertDecimal)
                {
                    TypeName = typeName
                };
                var minExtent = valueFactory.Create(CompareExtents.CURRENCYMIN.ToString(), typeName);
                var maxExtent = valueFactory.Create(CompareExtents.CURRENCYMAX.ToString(), typeName);
                summary.AddExtents(minExtent, maxExtent);
                return summary;
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

    public class UnreachableCaseInspectionSummaryClauseFactory : IUnreachableCaseInspectionSummaryClauseFactory
    {
        private static Dictionary<string, Tuple<long, long>> IntegerNumberExtents = new Dictionary<string, Tuple<long, long>>()
        {
            [Tokens.Long] = new Tuple<long, long>(CompareExtents.LONGMIN, CompareExtents.LONGMAX),
            [Tokens.Integer] = new Tuple<long, long>(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Int] = new Tuple<long, long>(CompareExtents.INTEGERMIN, CompareExtents.INTEGERMAX),
            [Tokens.Byte] = new Tuple<long, long>(CompareExtents.BYTEMIN, CompareExtents.BYTEMAX)
        };

        private static Dictionary<string, Tuple<double, double>> RationalNumberExtents = new Dictionary<string, Tuple<double, double>>()
        {
            [Tokens.Double] = new Tuple<double, double>(double.MinValue, double.MaxValue),
            [Tokens.Single] = new Tuple<double, double>(CompareExtents.SINGLEMIN, CompareExtents.SINGLEMAX)
        };

        public ISummaryCoverage Create(string typeName, IUnreachableCaseInspectionValueFactory valueFactory)
        {
            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                var summaryCoverage = new SummaryCoverage<long>(this, valueFactory, UCIValueConverter.ConvertLong)
                {
                    TypeName = typeName,
                };
                summaryCoverage.ApplyExtents(IntegerNumberExtents[typeName].Item1, IntegerNumberExtents[typeName].Item2);
                return summaryCoverage;
                
            }
            else if (RationalNumberExtents.Keys.Contains(typeName))
            {
                var summaryCoverage = new SummaryCoverage<double>(this, valueFactory, UCIValueConverter.ConvertDouble)
                {
                    TypeName = typeName
                };
                summaryCoverage.ApplyExtents(RationalNumberExtents[typeName].Item1, RationalNumberExtents[typeName].Item2);
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var summaryCoverage = new SummaryCoverage<decimal>(this, valueFactory, UCIValueConverter.ConvertDecimal)
                {
                    TypeName = typeName
                };
                summaryCoverage.ApplyExtents(CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX);
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var summaryCoverage = new SummaryCoverage<bool>(this, valueFactory, UCIValueConverter.ConvertBoolean)
                {
                    TypeName = typeName
                };
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.String))
            {
                var summaryCoverage = new SummaryCoverage<string>(this, valueFactory, UCIValueConverter.ConvertString)
                {
                    TypeName = typeName
                };
                return summaryCoverage;
            }
            return null;
        }
    }
}
