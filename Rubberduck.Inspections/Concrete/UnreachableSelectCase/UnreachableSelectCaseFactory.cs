using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public interface IUnreachableSelectFactory
    {
        ISummaryCoverage CreateSummaryCoverage(string typeName);
    }

    public class UnreachableSelectCaseFactory
    {
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

        public static ISummaryCoverage CreateSummaryCoverageShell(string typeName)
        {
            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                return CreateSummaryCoverage(typeName, IntegerNumberExtents[typeName].Item1, IntegerNumberExtents[typeName].Item2, Convert.ToInt64(true), Convert.ToInt64(false));
            }
            else if (RationalNumberExtents.Keys.Contains(typeName))
            {
                return CreateSummaryCoverage(typeName, RationalNumberExtents[typeName].Item1, RationalNumberExtents[typeName].Item2, Convert.ToDouble(true), Convert.ToDouble(false));
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                return CreateSummaryCoverage(typeName,CompareExtents.CURRENCYMIN, CompareExtents.CURRENCYMAX, Convert.ToDecimal(true), Convert.ToDecimal(false));
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var summaryCoverage = new SummaryCoverage<bool>()
                {
                    TrueValue = true,
                    FalseValue = false,
                    TypeName = typeName
                };
                return summaryCoverage;
            }
            else if (typeName.Equals(Tokens.String))
            {
                var summaryCoverage = new SummaryCoverage<string>()
                {
                    TrueValue = bool.TrueString,
                    FalseValue = bool.FalseString,
                    TypeName = typeName
                };
                return summaryCoverage;
            }
            return null;
        }

        public static ISummaryCoverage CreateSummaryCoverage(ParserRuleContext selectStmtCtxt, IParseTreeValueResults ptValues,string typeName)
        {
            Debug.Assert(selectStmtCtxt is VBAParser.SelectCaseStmtContext);
            var summaryCoverage = CreateSummaryCoverageShell(typeName);
            summaryCoverage.ParseTreeValueResults = ptValues;

            if (IntegerNumberExtents.Keys.Contains(typeName))
            {
                var concrete = (SummaryCoverage<long>)summaryCoverage;
                concrete.LoadRangeClauseCoverage(selectStmtCtxt, ptValues.ValueResultsAsLong());
                return concrete;

            }
            else if (RationalNumberExtents.Keys.Contains(typeName))
            {
                var concrete = (SummaryCoverage<double>)summaryCoverage;
                concrete.LoadRangeClauseCoverage(selectStmtCtxt, ptValues.ValueResultsAsDouble());
                return concrete;
            }
            else if (typeName.Equals(Tokens.Currency))
            {
                var concrete = (SummaryCoverage<decimal>)summaryCoverage;
                concrete.LoadRangeClauseCoverage(selectStmtCtxt, ptValues.ValueResultsAsDecimal());
                return concrete;
            }
            else if (typeName.Equals(Tokens.Boolean))
            {
                var concrete = (SummaryCoverage<bool>)summaryCoverage;
                concrete.LoadRangeClauseCoverage(selectStmtCtxt, ptValues.ValueResultsAsBoolean());
                return concrete;
            }
            else if (typeName.Equals(Tokens.String))
            {
                var concrete = (SummaryCoverage<string>)summaryCoverage;
                concrete.LoadRangeClauseCoverage(selectStmtCtxt, ptValues.ValueResultsAsString());
                return concrete;
            }
            return null;
        }

        public static IParseTreeVisitor<IParseTreeValueResults> CreateParseTreeVisitor(RubberduckParserState state, string evaluationTypeName = "")
        {
            if (evaluationTypeName.Equals(""))
            {
                return new ContextValueVisitor(state);
            }
            return new ContextValueVisitor(state, evaluationTypeName);
        }

        private static ISummaryCoverage CreateSummaryCoverage<T>(string typeName, T min, T max, T trueVal, T falseVal) where T : IComparable<T>
        {
            var summaryCoverage = new SummaryCoverage<T>();
            summaryCoverage.ApplyExtents(min, max);
            summaryCoverage.TrueValue = trueVal;
            summaryCoverage.FalseValue = falseVal;
            summaryCoverage.TypeName = typeName;
            return summaryCoverage;
        }
    }
}
