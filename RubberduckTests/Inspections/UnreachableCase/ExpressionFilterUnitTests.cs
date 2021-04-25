using NUnit.Framework;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;
using Rubberduck.Refactorings;
using Rubberduck.Refactoring.ParseTreeValue;

namespace RubberduckTests.Inspections.UnreachableCase
{
    /*
        ExpressionFilter is a support class of the UnreachableCaseInspection

        Notes:
        FilterContentToken Parameter encoding:
            Min!5 is the result of adding expression: Is < 5"
            Max!5 is the result of adding expression: Is > 5"
            Range!5:50 adds range 5 To 50"
            Value!5 adds a single value "5"
            RelOp!x < 5 adds RelationalOp "x < 5"
            Or, RelOp!True is interpreted as "Case True" <- simulates a resolved expression like x < 7, where 'x' is a constant == 6

        RangeClauseToken encoding:
            <operand>?<declaredType>_<opSymbol> _<operand>?<declaredType>, <expression>,<filterTypeName>
            If there is no "?<declaredType>", then<operand>'s type is derived by the ParseTreeValue instance.
            The<filterTypeName> is the type that the calculation must yield in order to
            make comparisons.

        When comparing filters for test pass/fail:
            FilterContentTokens are used to load the 'expected' filter.
            RangeClauseTokens are used to load the filter under test (actual)
    */

    [TestFixture]
    public class ExpressionFilterUnitTests
    {
        private const string RANGECLAUSE_DELIMITER = ",";
        private const string VALUE_TYPE_DELIMITER = "?";
        private const string OPERAND_DELIMITER = "_";
        private const string CLAUSETYPE_VALUE_DELIMITER = "!";
        private const string RANGE_STARTEND_DELIMITER = ":";

        private readonly  Lazy<IParseTreeValueFactory> _valueFactory = new Lazy<IParseTreeValueFactory>(() => new ParseTreeValueFactory());
        private IParseTreeValueFactory ValueFactory => _valueFactory.Value;

        [TestCase("Min!-5000", "", "Min(-5000)Max(typeMax)")]
        [TestCase("Min!-5000,Max!5000", "", "Min(-5000)Max(5000)")]
        [TestCase("Max!5,Min!0", "", "Min(0)Max(5)")]
        [TestCase("Min!5", "Max!300", "Min(5)Max(300)")]
        [TestCase("Min!5,Range!45:55", "Max!300", "Min(5)Max(300)Ranges(45:55)")]
        [TestCase("Min!5,Range!45:55", "Max!300,Value!200", "Min(5)Max(300)Ranges(45:55)Values(200)")]
        [TestCase("Min!-2,Range!45:55", "Max!300,Value!200,RelOp!x < 50", "Min(-2)Max(300)Ranges(45:55)Values(200)Predicates(x < 50)")]
        [TestCase("Min!-5000,Max!5000,Range!45:55", "Range!60:65", "Min(-5000)Max(5000)Ranges(45:55,60:65)")]
        [TestCase("Min!-5000,Max!5000,Value!45,Value!46", "Value!60", "Min(-5000)Max(5000)Values(45,46,60)")]
        [TestCase("RelOp!x < 50", "RelOp!x > 75", "Min(typeMin)Max(typeMax)Predicates(x < 50,x > 75)")]
        [TestCase("Min!-5000", "", "Min(-5000)Max(typeMax)")]
        [TestCase("Max!-5000", "", "Min(typeMin)Max(-5000)")]
        [TestCase("RelOp!z < x", "RelOp!y < 35", "Min(typeMin)Max(typeMax)Predicates(y < 35,x > z)")]
        [TestCase("RelOp!x < 55", "RelOp!y < 35", "Min(typeMin)Max(typeMax)Predicates(x < 55,y < 35)")]
        [TestCase("RelOp!x < 55", "RelOp!x < 55", "Min(typeMin)Max(typeMax)Predicates(x < 55)")]
        [Category("Inspections")]
        public void ExpressionFilter_ToString(string firstCase, string secondCase, string expected)
        {
            var filter = FilterContentTokensToFilter(new string[] { firstCase, secondCase }, Tokens.Long);
            Assert.AreEqual(expected, filter.ToString());
        }

        [TestCase("x_<_65,x_<_55", "RelOp!x < 65,RelOp!x < 55", "Long")]
        [TestCase("x_>_55,x_<_65", "Value!-1,RelOp!x > 55,RelOp!x < 65", "Long")]
        [TestCase("x_>_65,x_<_55", "Value!0, RelOp!x > 65,RelOp!x < 55", "Long")]
        [TestCase("x_<>_55,x_<>_60,x_=_95", "Value!-1,RelOp!x <> 55,RelOp!x <> 60,RelOp!x = 95", "Long")]
        [TestCase("0,x_>_55,x_<_65,x_=_70", "RelOp!x > 55,RelOp!x < 65,Value!0,Value!-1", "Long")]
        [TestCase("-1,x_=_55,x_=_65,x_=_0", "RelOp!x = 55,RelOp!x = 65,Value!-1,Value!0", "Long")]
        [TestCase("x_<=_55,x_>_65", "Value!0,RelOp!x <= 55,RelOp!x > 65", "Long")]
        [TestCase("x_<_65.45,x_<_55.97", "RelOp!x < 65.45,RelOp!x < 55.97", "Single")]
        [TestCase("x_>_55.97,x_<_65.45", "Value!-1,RelOp!x > 55.97,RelOp!x < 65.45", "Single")]
        [TestCase("x_>_65.45,x_<_55.97", "Value!0, RelOp!x > 65.45,RelOp!x < 55.97", "Single")]
        [TestCase("x_<>_55.97,x_<>_60,x_=_95", "Value!-1,RelOp!x <> 55.97,RelOp!x <> 60,RelOp!x = 95", "Single")]
        [TestCase("0,x_>_55.97,x_<_65.45,x_=_70", "RelOp!x > 55.97,RelOp!x < 65.45,Value!0,Value!-1", "Single")]
        [TestCase("-1,x_=_55.97,x_=_65.45,x_=_0", "RelOp!x = 55.97,RelOp!x = 65.45,Value!-1,Value!0", "Single")]
        [TestCase("x_<=_55.97,x_>_65.45", "Value!0,RelOp!x <= 55.97,RelOp!x > 65.45", "Single")]
        [Category("Inspections")]
        public void ExpressionFilter_VariableRelationalOps(string firstCase, string expected, string typeName)
        {
            var actualFilter = ExpressionFilterFactory.Create(typeName);
            actualFilter.AddComparablePredicateFilter("x", typeName);

            actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, typeName, actualFilter);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expected }, typeName);

            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase("150?Long_To_50?Long", "Long", "")]
        [TestCase("50?Long_To_50?Long", "Boolean", "Value!True")]
        [TestCase("50?Long_To_50?Long", "Long", "Value!50")]
        [TestCase("50?Long_To_x?Long", "Long", "Range!50:x")]
        [TestCase("50?Long_To_100?Long", "Long", "Range!50:100")]
        [TestCase(@"""Nuts""?String_To_""Soup""?String", "String", @"Range!""Nuts"":""Soup""")]
        [TestCase(@"""Soup""?String_To_""Nuts""?String", "String", "")]
        [TestCase("50.3?Double_To_100.2?Double", "Long", "Range!50:100")]
        [TestCase("50.3?Double_To_100.2?Double", "Double", "Range!50.3:100.2")]
        [TestCase("50_To_100,75_To_125", "Long", "Range!50:100,Range!75:125")]
        [TestCase("50_To_100,175_To_225", "Long", "Range!50:100,Range!175:225")]
        [TestCase("500?Long_To_100?Long", "Long", "")]
        [Category("Inspections")]
        public void ExpressionFilter_AddRangeOfValuesClauses(string firstCase, string filterTypeName, string expectedRangeClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, filterTypeName);
            var expected = FilterContentTokensToFilter(new string[] { expectedRangeClauses }, filterTypeName);
            Assert.AreEqual(expected, actualFilter);
        }

        [TestCase("Is_>_50", "Long", "Max!50")]
        [TestCase("Is_>_50.49", "Long", "Max!50.49")]
        [TestCase("Is_>_50#", "Double", "Max!50")]
        [TestCase("Is_>_True", "Boolean", "RelOp!Is > True")]
        [TestCase("Is_>=_50", "Long", "Max!50,Value!50")]
        [TestCase("Is_>=_50.49", "Double", "Max!50.49,Value!50.49")]
        [TestCase("Is_>=_50#", "Double", "Max!50,Value!50")]
        [TestCase("Is_>=_True", "Boolean", "Value!True,Value!False")]
        [TestCase("Is_<_50", "Long", "Min!50")]
        [TestCase("Is_<_50,Is_<_25", "Long", "Min!50")]
        [TestCase("Is_<_50,Is_<_75", "Long", "Min!75")]
        [TestCase("Is_<_50,Is_<_75,Is_>_300", "Long", "Min!75,Max!300")]
        [TestCase("Is_<=_50", "Long", "Min!50,Value!50")]
        [TestCase("Is_<=_50,Is_>=_51", "Long", "Min!50,Max!51,Value!50,Value!51")]
        [TestCase("Is_=_100", "Long", "Value!100")]
        [TestCase("Is_=_100.49", "Double", "Value!100.49")]
        [TestCase("Is_=_100#", "Double", "Value!100")]
        [TestCase("Is_=_True", "Long", "Value!-1")]
        [TestCase(@"Is_=_""100""", "Long", "Value!100")]
        [TestCase("Is_<>_100", "Long", "Min!100,Max!100")]
        [TestCase("Is_<>_100.49", "Double", "Min!100.49,Max!100.49")]
        [TestCase("Is_<>_100#", "Double", "Min!100,Max!100")]
        [TestCase("Is_<>_True", "Boolean", "RelOp!Is <> True")]
        [TestCase(@"Is_<>_""100""", "Long", "Min!100,Max!100")]
        [TestCase("Is_>_x", "Long", "RelOp!Is > x")]
        [Category("Inspections")]
        public void ExpressionFilter_AddIsClause(string firstCase, string filterTypeName, string expectedRangeClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, filterTypeName);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedRangeClauses }, filterTypeName);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase("Is_<_45.61", "Is_>_45.6", "Double")]
        [TestCase("Is_<_45.61,Is_>_60.5", "39.2_To_66.1", "Double")]
        [TestCase("False_To_False", "50", "Boolean")]
        [TestCase("True_To_False", "", "Boolean")]
        [TestCase("-5000", "False", "Boolean")]
        [TestCase("True", "0", "Boolean")]
        [TestCase("Is_<_5", "Is_>_-5000", "Long")]
        [TestCase("Is_<_40,Is_>_40", "35_To_45", "Long")]
        [TestCase("Is_<_40,Is_>_44", "35_To_45", "Long")]
        [TestCase("Is_<_40,Is_>_40", "40", "Long")]
        [TestCase("Is_>_240,150_To_239", "240, 0,1,2_To_150", "Byte")]
        [TestCase("151_To_255", "150,0,1,2_To_149", "Byte")]
        [TestCase("Is_<_13,Is_>_30,13_To_100", "", "Long")]
        [Category("Inspections")]
        public void ExpressionFilter_FiltersAll(string firstCase, string secondCase, string SelectExpessionTypeName)
        {
            var filter = RangeClauseTokensToFilter(new string[] { firstCase, secondCase }, SelectExpessionTypeName);
            Assert.IsTrue(filter.FiltersAllValues, filter.ToString());
        }

        [TestCase("x_<_3", "Is_<_3", "RelOp!x < 3,Min!3")]
        [TestCase("Not_x", "-x,Not_x", "RelOp!Not x,Value!-x")]
        [TestCase("Is_<_x,Is_<_y", "Is_<_x", "RelOp!Is < x,RelOp!Is < y")]
        [TestCase("-x,-y", "-x", "Value!-x,Value!-y")]
        [TestCase("3_To_55", "x.Item(2)", "Range!3:55,Value!x.Item(2)")]
        [TestCase("3_To_55", "Is_<_6", "Min!6,Range!6:55")]
        [TestCase("3_To_55", "Is_>_6", "Max!6,Range!3:6")]
        [TestCase("Is_<_6", "1_To_5", "Min!6")]
        [TestCase("5,6,7", "Is_>_6", "Max!6,Value!5,Value!6")]
        [TestCase("5,6,7", "Is_<_6", "Min!6,Value!6,Value!7")]
        [TestCase("Is_<_5,Is_>_75", "85", "Min!5,Max!75")]
        [TestCase("Is_<_5,Is_>_75", "0", "Min!5,Max!75")]
        [TestCase("45_To_85", "50", "Range!45:85")]
        [TestCase("5,6,7,8", "6_To_8", "Range!6:8,Value!5")]
        [TestCase("Is_<_400,15_To_160", "500_To_505", "Min!400,Range!500:505")]
        [TestCase("101_To_149", "15_To_160", "Range!15:160")]
        [TestCase("101_To_149", "15_To_148", "Range!15:149")]
        [TestCase("150_To_250,1_To_100,101_To_149", "25_To_249", "Range!1:250")]
        [TestCase("150_To_250,1_To_100,-5_To_-2,101_To_149", "25_To_249", "Range!-5:-2,Range!1:250")]
        [TestCase("5_To_5,x,y", "", "Value!5,Value!x,Value!y")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersInteger(string firstCase, string secondCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase, secondCase }, Tokens.Long);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.Long);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase("101.45_To_149.00007", "101.57_To_110.63", "Range!101.45:149.00007")]
        [TestCase("101.45_To_149.0007", "15.67_To_148.9999", "Range!15.67:149.0007")]
        [TestCase("101.45_To_149.2", "149.2_To_150.5", "Range!101.45:150.5")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersRational(string firstCase, string secondCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase, secondCase }, Tokens.Double);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.Double);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase(@"""Alpha""_To_""Omega"",""Nuts""_To_""Soup""", @"Range!""Alpha"":""Soup""")]
        [TestCase(@"""Alpha"",""Nuts"",""Alpha""", @"Value!""Alpha"", Value!""Nuts""")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersStrings(string firstCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, Tokens.String);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.String);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase("0_To_10", "")]
        [TestCase("10_To_0", "Value!True,Value!False")]
        [TestCase("False_To_True,x_<_3", "RelOp!x < True")]
        [TestCase("True_To_True", "Value!True")]
        [TestCase("True_To_False", "Value!False,Value!True")]
        [TestCase("Is_>_5,x_<_5", "RelOp!Is > True,RelOp!x < 5")]
        [TestCase("Is_<_5,x_<_5", "RelOp!x < 5")]
        [TestCase("-1,0,x_<_3", "Value!True,Value!False")]
        [TestCase("-5_To_15,x_<_3", "Value!True,RelOp!x < True")]
        [TestCase("Is_>_1,x_<_3", "Max!1,RelOp!x < True")]
        [TestCase("Is_<_-2,x_<_3", "RelOp!x < True")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersBoolean(string firstCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, Tokens.Boolean);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.Boolean);
            Assert.AreEqual(expectedFilter, actualFilter);
            if (expectedClauses.Length > 0)
            {
                Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
                Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
            }
        }

        [TestCase("Is_<_#1/1/2001#,#12/1/2000#_To_#1/10/2001#", "Min!#1/1/2001#,Range!#1/1/2001#:#1/10/2001#")]
        [TestCase("Is_<_#1/1/2001#", "Min!#1/1/2001#")]
        [TestCase("#1/1/2001#_To_#1/10/2001#", "Range!#1/1/2001#:#1/10/2001#")]
        [TestCase("#1/1/2001#", "Value!#1/1/2001#")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersDate(string firstCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase }, Tokens.Date);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.Date);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        /*
         * The test cases below cover the truth table
         * for 'Is' clauses present in Boolean Select Case Statements.
         * Cases that always resolve to True/False are store both as Single values.
         * All others (outcome depends on the Select Case value) are 
         * stored as variable Is clauses.
        */

        [TestCase("Is_<_True", "")] //Inherently unreachable
        [TestCase("Is_<=_True", "RelOp!Is <= True")]
        [TestCase("Is_>_True", "RelOp!Is > True")]
        [TestCase("Is_>=_True", "Value!False,Value!True")] //Filters both True and False
        [TestCase("Is_=_True", "RelOp!Is = True")]
        [TestCase("Is_<>_True", "Value!False")]
        [TestCase("Is_>_False", "")] //Inherently unreachable
        [TestCase("Is_>=_False", "RelOp!Is >= False")]
        [TestCase("Is_<_False", "RelOp!Is < False")]
        [TestCase("Is_<=_False", "Value!False,Value!True")] //Filters both True and False
        [TestCase("Is_=_False", "RelOp!Is = False")]
        [TestCase("Is_<>_False", "Value!True")]
        [Category("Inspections")]
        public void ExpressionFilter_BooleanIsClauseTruthTable(string rangeClause, string expected)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { rangeClause }, Tokens.Boolean);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expected }, Tokens.Boolean);
            Assert.AreEqual(expectedFilter, actualFilter);
            if (expected.Length > 0)
            {
                Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
                Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
            }
        }

        [TestCase("True","", "Value!False")]
        [TestCase("False", "", "Value!True")]

        [TestCase("True", "Is_<_True", "Value!False")]
        [TestCase("False", "Is_<_True", "Value!True")]
        [TestCase("True", "Is_<_False", "Value!True,Value!False")]
        [TestCase("False", "Is_<_False", "Value!True,Value!False")]

        [TestCase("True", "Is_<=_True", "Value!True,Value!False")]
        [TestCase("False", "Is_<=_True", "Value!True,Value!False")]
        [TestCase("True", "Is_<=_False", "Value!True,Value!False")]
        [TestCase("False", "Is_<=_False", "Value!True")]

        [TestCase("True", "Is_>_True", "Value!False")]
        [TestCase("False", "Is_>_True", "Value!True")]
        [TestCase("True", "Is_>_False", "Value!False")]
        [TestCase("False", "Is_>_False", "Value!True")]

        [TestCase("True", "Is_>=_True", "Value!False,Value!True")]
        [TestCase("False", "Is_>=_True", "Value!True")]
        [TestCase("True", "Is_>=_False", "Value!False")]
        [TestCase("False", "Is_>=_False", "Value!True")]

        [TestCase("True", "Is_=_False", "Value!False")]
        [TestCase("False", "Is_=_False", "Value!True")]
        [TestCase("True", "Is_=_True", "Value!False,Value!True")]
        [TestCase("False", "Is_=_True", "Value!True,Value!False")]

        [TestCase("True", "Is_<>_False", "Value!False,Value!True")]
        [TestCase("False", "Is_<>_False", "Value!False,Value!True")]
        [TestCase("True", "Is_<>_True", "Value!False")]
        [TestCase("False", "Is_<>_True", "Value!True")]
        [Category("Inspections")]
        public void ExpressionFilter_BooleanIsClauseTruthTableConstSelectExpression(string selectExpressionValue, string rangeClause, string expectedFilterContent)
        {
            var actualFilter = ExpressionFilterFactory.Create(Tokens.Boolean);
            actualFilter.SelectExpressionValue = ValueFactory.Create(selectExpressionValue.Equals(Tokens.True)); //, Tokens.Boolean);

            actualFilter = RangeClauseTokensToFilter(new string[] { rangeClause }, Tokens.Boolean, actualFilter);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedFilterContent }, Tokens.Boolean);
            Assert.AreEqual(expectedFilter, actualFilter);
            if (expectedFilterContent.Length > 0)
            {
                Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
                Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
            }
        }

        [TestCase(@"x_Like_""*"",y_Like_""aa?*""", @"Value!True, RelOp!y Like ""aa?*""")]
        [TestCase(@"x_Like_""*"",x_Like_""aa?*""", @"Value!True,RelOp!x Like ""aa?*""")]
        [TestCase(@"x_Like_""*Bar""", @"RelOp!x Like ""*Bar""")]
        [TestCase(@"x_Like_""[A-Z]*""", @"RelOp!x Like ""[A-Z]*""")]
        [TestCase(@"x_Like_""[A-Z]*"", x_Like_""Fo*oBar""", @"RelOp!x Like ""[A-Z]*"",RelOp!x Like ""Fo*oBar""")]
        [Category("Inspections")]
        public void ExpressionFilter_LikeAddToFilters(string rangeClause, string expectedClause)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { rangeClause }, Tokens.String);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClause}, Tokens.String);
            Assert.AreEqual(expectedFilter, actualFilter);
        }

        [TestCase(@"x_Like_""*Bar""",@"x_Like_""*Bar""", @"RelOp!x Like ""*Bar""")]
        [TestCase(@"x_Like_""*Bar""","True", @"RelOp!x Like ""*Bar"",Value!True")]
        [TestCase(@"x_Like_""*""", "True", "Value!True")]
        [Category("Inspections")]
        public void ExpressionFilter_LikeFiltersDuplicates(string firstCase, string secondCase, string expectedClauses)
        {
            var actualFilter = RangeClauseTokensToFilter(new string[] { firstCase, secondCase }, Tokens.Boolean);
            var expectedFilter = FilterContentTokensToFilter(new string[] { expectedClauses }, Tokens.Boolean);
            Assert.AreEqual(expectedFilter, actualFilter);
            Assert.True(actualFilter.HasFilters, "'Actual' Filter not created");
            Assert.True(expectedFilter.HasFilters, "'Expected' Filter not created");
        }

        [TestCase("Is_<_True", "Boolean")]
        [TestCase("Is_>_False", "Boolean")]
        [TestCase("5_To_3", "Long")]
        [TestCase("False_To_True", "Boolean")]
        [TestCase("#7/4/2018#_To_#7/4/1776#", "Date")]
        [TestCase("44.44_To_36.2", "Double")]
        [TestCase("44.44_To_36.2", "Single")]
        [TestCase("44.4444_To_36.2000", "Currency")]
        [Category("Inspections")]
        public void ExpressionFilter_MalformedRangeOfValues(string rangeClause, string filterTypeName)
        {
            var filter = ExpressionFilterFactory.Create(filterTypeName);
            var expression = RangeClauseTokensToExpressions(new string[] { rangeClause }, filterTypeName).First();
            filter.CheckAndAddExpression(expression);
            Assert.IsTrue(expression.IsInherentlyUnreachable);
        }

        private List<string> RetrieveDelimitedElements(string rangeClauses, string delimiter)
        {
            var isClauseTypeDelimiter = delimiter.Equals(CLAUSETYPE_VALUE_DELIMITER);
            var result = new List<string>();
            if (isClauseTypeDelimiter)
            {

                var index = rangeClauses.IndexOf(CLAUSETYPE_VALUE_DELIMITER);
                if (index < 0)
                {
                    return result;
                }
                var lhs = rangeClauses.Substring(0, index);
                var rhs = rangeClauses.Substring(index + 1);
                result.Add(lhs);
                result.Add(rhs);
            }
            else
            {
                var clauses = rangeClauses.Split(new string[] { delimiter }, StringSplitOptions.None);
                foreach (var clause in clauses)
                {
                    result.Add(clause.Trim());
                }
            }
            return result;
        }

        private (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) GetBinaryOpValues(string delimitedElements)
        {

            var elements = RetrieveDelimitedElements(delimitedElements, OPERAND_DELIMITER);
            if (elements.Count() != 3)
            {
                Assert.Inconclusive("Invalid number of operands passed to 'GetBinaryOpValues(...)'");
            }

            var LHS = CreateInspValueFrom(elements[0]);
            var RHS = CreateInspValueFrom(elements[2]);
            return (LHS, RHS, elements[1]);
        }

        private (IParseTreeValue Operand, string Symbol) GetUnaryOpValues(string delimitedElements)
        {
            const int MAX_ELEMENTS = 2;
            var elements = RetrieveDelimitedElements(delimitedElements, OPERAND_DELIMITER);
            if (elements.Count() > MAX_ELEMENTS)
            {
                Assert.Inconclusive("Invalid number of operands passed to 'GetUnaryOpValues(...)'");
            }

            var operand = elements.Count() == 2 ?
                CreateInspValueFrom(elements[1])
                : CreateInspValueFrom(elements[0]);
            return elements.Count() == MAX_ELEMENTS ? (operand, elements[0]) : (operand, string.Empty);
        }

        private IRangeClauseExpression CreateRangeClauseExpression((IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) expressionElements)
        {
            if (expressionElements.LHS.Token.Equals(Tokens.Is))
            {
                return new IsClauseExpression(expressionElements.RHS, expressionElements.Symbol);
            }
            else if (expressionElements.Symbol.Equals(Tokens.To))
            {
                return new RangeOfValuesExpression((expressionElements.LHS, expressionElements.RHS));
            }
            else if (expressionElements.Symbol.Equals(Tokens.Like))
            {
                return new LikeExpression(expressionElements.LHS, expressionElements.RHS);
            }
            else
            {
                return new BinaryExpression(expressionElements.LHS, expressionElements.RHS, expressionElements.Symbol);
            }
        }

        private IParseTreeValue CreateInspValueFrom(string valAndType)
        {
            if (InspectableDelimited.Any(id => valAndType.Contains(id)))
            {
                var args = RetrieveDelimitedElements(valAndType, VALUE_TYPE_DELIMITER);
                var value = args[0];
                var declaredType = args[1].Equals(string.Empty) ? null : args[1];
                if (declaredType is null)
                {
                    return ValueFactory.Create(value);
                }
                //var ptValue = ValueFactory.Create(value, declaredType);
                var ptValue = ValueFactory.CreateDeclaredType(value, declaredType);
                return ptValue;
            }
            return ValueFactory.Create(valAndType);
        }

        private IExpressionFilter RangeClauseTokensToFilter(IEnumerable<string> rangeClauseTokens, string filterTypeName, IExpressionFilter filter = null)
        {
            if (filter is null)
            {
                filter = ExpressionFilterFactory.Create(filterTypeName);
            }

            var rangeClauses = new List<string>();
            foreach( var token in rangeClauseTokens)
            {
                rangeClauses.AddRange(RetrieveDelimitedElements(token, RANGECLAUSE_DELIMITER));
            }

            var expressions = RangeClauseTokensToExpressions(rangeClauses, filterTypeName);
            foreach (var expression in expressions)
            {
                filter.CheckAndAddExpression(expression);
            }
            return filter;
        }

        private IEnumerable<IRangeClauseExpression> RangeClauseTokensToExpressions(IEnumerable<string> rangeClauseLiterals, string filterTypeName)
        {
            var results = new List<IRangeClauseExpression>();
            foreach (var clause in rangeClauseLiterals.Where(rc => !rc.Equals(string.Empty)))
            {
                IRangeClauseExpression expressionClause = null;
                var operandDelimiters = clause.Where(ch => ch.Equals('_'));
                if (operandDelimiters.Count() == 2)
                {
                    var (lhs, rhs, symbol) = GetBinaryOpValues(clause);
                    expressionClause = CreateRangeClauseExpression((lhs, rhs, symbol));
                }
                else if (operandDelimiters.Count() <= 1)
                {
                    var (operand, symbol) = GetUnaryOpValues(clause);
                    if (symbol.Equals(LogicalOperators.NOT))
                    {
                        expressionClause = new UnaryExpression(operand, symbol);
                    }
                    else
                    {
                        //expressionClause = new ValueExpression(ValueFactory.Create($"{symbol}{operand}", filterTypeName));
                        expressionClause = new ValueExpression(ValueFactory.CreateDeclaredType($"{symbol}{operand}", filterTypeName));
                    }
                }
                else
                {
                    Assert.Inconclusive("unable to parse operands");
                }
                results.Add(expressionClause);
            }
            return results;
        }

        private IExpressionFilter FilterContentTokensToFilter(IEnumerable<string> caseClauses, string typeName, IExpressionFilter filter = null)
        {
            if(filter is null)
            {
                filter = ExpressionFilterFactory.Create(typeName);
            }

            var expressions = new List<IRangeClauseExpression>();
            foreach(var caseClause in caseClauses)
            {
                var rangeClauses = RetrieveDelimitedElements(caseClause, RANGECLAUSE_DELIMITER);
                expressions.AddRange(CreateTestExpressions(rangeClauses));
            }

            foreach (var expr in expressions)
            {
                filter.CheckAndAddExpression(expr);
            }
            return filter;
        }

        private List<IRangeClauseExpression> CreateTestExpressions(IEnumerable<string> annotations)
        {
            var results = new List<IRangeClauseExpression>();
            var clauseItem = string.Empty;
            foreach (var item in annotations)
            {
                clauseItem = item;

                var element = RetrieveDelimitedElements(clauseItem.Trim(), CLAUSETYPE_VALUE_DELIMITER);
                if ( !element.Any() || element[0].Equals(string.Empty) || element.Count() < 2)
                {
                    continue;
                }
                var clauseType = element[0];
                var clauseExpression = element[1];
                var values = RetrieveDelimitedElements(clauseExpression, RANGECLAUSE_DELIMITER);
                foreach (var expr in values)
                {
                    if (clauseType.Equals("Min"))
                    {
                        var uciVal = ValueFactory.Create(clauseExpression);
                        results.Add(new IsClauseExpression(uciVal, RelationalOperators.LT));
                    }
                    else if (clauseType.Equals("Max"))
                    {
                        var uciVal = ValueFactory.Create(clauseExpression);
                        results.Add(new IsClauseExpression(uciVal, RelationalOperators.GT));
                    }
                    else if (clauseType.Equals("Range"))
                    {
                        var startEnd = clauseExpression.Split(new string[] { RANGE_STARTEND_DELIMITER }, StringSplitOptions.None);
                        var rangeStart = ValueFactory.Create(startEnd[0]);
                        var rangeEnd = ValueFactory.Create(startEnd[1]);
                        results.Add(new RangeOfValuesExpression((rangeStart, rangeEnd)));
                    }
                    else if (clauseType.Equals("Value"))
                    {
                        var testVal = ValueFactory.Create(clauseExpression);
                        results.Add(new ValueExpression(testVal));
                    }
                    else if (clauseType.Equals("RelOp"))
                    {
                        string symbol = string.Empty;
                        TryExtractSymbol(item, out symbol);
                        var sides = clauseExpression.Split(new string[] { symbol }, StringSplitOptions.None);

                        if (sides.Count() == 2 && sides.All(sd => !sd.Equals(string.Empty)))
                        {
                            var lhs = ValueFactory.Create(sides[0].Trim());
                            var rhs = ValueFactory.Create(sides[1].Trim());
                            if (lhs.Token.Equals(Tokens.Is))
                            {
                                results.Add(new IsClauseExpression(rhs, symbol));
                            }
                            else if (symbol.Equals(Tokens.Like))
                            {
                                results.Add(new LikeExpression(lhs, rhs));
                            }
                            else
                            {
                                results.Add(new BinaryExpression(lhs, rhs, symbol));
                            }

                        }
                        else if (sides.Count() == 2 && sides.Any(sd => sd.Equals(string.Empty)))
                        {
                            var validValue = sides.First(sd => !sd.Equals(string.Empty));
                            var lhs = ValueFactory.Create(validValue.Trim());
                            results.Add(new UnaryExpression(lhs, symbol));
                        }
                        else
                        {
                            var uciVal = ValueFactory.Create(clauseExpression);
                            results.Add(new UnaryExpression(uciVal, symbol));
                        }
                    }
                    else if (clauseType.Equals("Is"))
                    {
                        string symbol = string.Empty;
                        TryExtractSymbol(item, out symbol);
                        var sides = clauseExpression.Split(new string[] { symbol }, StringSplitOptions.None);

                        if (sides.Count() == 2)
                        {
                            var lhs = ValueFactory.Create(sides[0].Trim());
                            var rhs = ValueFactory.Create(sides[1].Trim());
                            results.Add(new IsClauseExpression(rhs, symbol));
                        }
                    }
                    else
                    {
                        Assert.Fail($"Invalid clauseType ({clauseType}) encountered");
                    }
                }
            }
            return results;
        }

        private bool TryExtractSymbol(string item, out string symbol)
        {
            symbol = string.Empty;
            var matchedSymbols = RelationalOperators.SymbolList.Where(sym => item.Contains($" {sym} "));
            if (!matchedSymbols.Any())
            {
                matchedSymbols = LogicalOperators.SymbolList.Where(sym => item.Contains($" {sym} "));
            }
            if (matchedSymbols.Any())
            {
                symbol = matchedSymbols.First();
                return true;
            }
            // one more look to check for unary expression 'Not x'
            matchedSymbols = RelationalOperators.SymbolList.Where(sym => item.Contains($"{sym} "));
            if (!matchedSymbols.Any())
            {
                matchedSymbols = LogicalOperators.SymbolList.Where(sym => item.Contains($"{sym} "));
            }
            if (matchedSymbols.Any())
            {
                symbol = matchedSymbols.First();
                return true;
            }

            return false;
        }

        private static List<string> InspectableDelimited = new List<string>()
        {
            "?Long",
            "?Integer",
            "?Byte",
            "?Double",
            "?Single",
            "?Currency",
            "?Boolean",
            "?String",
        };
    }
}
