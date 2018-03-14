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
    public interface ISummaryCoverageFactory
    {
        ISummaryCoverage Create(string typeName);
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

    public class SummaryCoverageFactory : ISummaryCoverageFactory
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

        public ISummaryCoverage Create(string typeName)
        {
            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                var summaryCoverage = new SummaryCoverage<long>(this, -1/*Convert.ToInt64(true)*/, Convert.ToInt64(false))
                {
                    TypeName = typeName
                };
                summaryCoverage.ApplyExtents(IntegerNumberExtents[typeName].Item1, IntegerNumberExtents[typeName].Item2);
                summaryCoverage.TConverter = UCIValueConverter.ConvertLong;
                return summaryCoverage;
                
            }
            else if (RationalNumberExtents.Keys.Contains(typeName))
            {
                var summaryCoverage = new SummaryCoverage<double>(this, Convert.ToDouble(true), Convert.ToDouble(false))
                {
                    TypeName = typeName
                };
                summaryCoverage.ApplyExtents(RationalNumberExtents[typeName].Item1, RationalNumberExtents[typeName].Item2);
                summaryCoverage.TConverter = UCIValueConverter.ConvertDouble;
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var summaryCoverage = new SummaryCoverage<decimal>(this, Convert.ToDecimal(true), Convert.ToDecimal(false))
                {
                    TypeName = typeName
                };
                summaryCoverage.ApplyExtents(CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX);
                summaryCoverage.TConverter = UCIValueConverter.ConvertDecimal;
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var summaryCoverage = new SummaryCoverage<bool>(this, true, false)
                {
                    TypeName = typeName
                };
                summaryCoverage.TConverter = UCIValueConverter.ConvertBoolean;
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.String))
            {
                //TODO: verify the true/false values are meaningful for SummaryClause<string>
                var summaryCoverage = new SummaryCoverage<string>(this, Tokens.True, Tokens.False)
                {
                    TypeName = typeName
                };
                summaryCoverage.TConverter = UCIValueConverter.ConvertString;
                return summaryCoverage;
            }
            return null;
        }
    }
}
