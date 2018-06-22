using NUnit.Framework;
using Rubberduck.Inspections.Concrete.UnreachableCaseInspection;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RubberduckTests.Inspections.UnreachableCase
{
    /*
        RangeClauseFilter is a support class of the UnreachableCaseInspection

        Notes:
        Filter Parameter encoding
        Min!5 is interpreted as to "Case Is < 5"
        Max!5 is interpreted as to "Case Is > 5"
        Range!5:50 is interpreted as to "Case 5 To 50"
        Value!5 is interpreted as to "Case 5"
        RelOp!x < 5 is interpreted as to "Case x < 5"
        Or, RelOp!True is interpreted as "Case True" <- simulates a resolved expression like x < 7, where 'x' is a constant == 6


        Operations encoding:
        <operand>?<declaredType>_<mathSymbol> _<operand>?<declaredType>, <expression>,<selectExpressionType>
        If there is no "?<declaredType>", then<operand>'s type is derived by the ParseTreeValue instance.
        The<selectExpressionType> is the type that the calculation must yield in order to
        make comparisons in the Select Case statement under inspection.
    */

    [TestFixture]
    public class ExpressionFilterUnitTests
    {
        private const string RANGECLAUSE_DELIMITER = ",";
        private const string VALUE_TYPE_DELIMITER = "?";
        private const string OPERAND_DELIMITER = "_";
        private const string CLAUSETYPE_VALUE_DELIMITER = "!";
        private const string RANGE_STARTEND_DELIMITER = ":";

        private IUnreachableCaseInspectionFactoryProvider _factoryProvider;
        private IParseTreeValueFactory _valueFactory;

        private IUnreachableCaseInspectionFactoryProvider FactoryProvider
        {
            get
            {
                if (_factoryProvider is null)
                {
                    _factoryProvider = new UnreachableCaseInspectionFactoryProvider();
                }
                return _factoryProvider;
            }
        }

        private IParseTreeValueFactory ValueFactory
        {
            get
            {
                if (_valueFactory is null)
                {
                    _valueFactory = FactoryProvider.CreateIParseTreeValueFactory();
                }
                return _valueFactory;
            }
        }

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
            var filter = ExpressionFilterFactory.Create(Tokens.Long);
            var expressions = RangeDescriptorsToExpressions(new string[] { firstCase, secondCase }, Tokens.Long);
            foreach (var expr in expressions)
            {
                filter.AddExpression(expr);
            }
            Assert.AreEqual(expected, filter.ToString());
        }

        [TestCase("RelOp!x < 65", "RelOp!x < 55", "RelOp!x < 65,RelOp!x < 55")]
        [TestCase("RelOp!x > 55", "RelOp!x < 65", "Value!-1,RelOp!x > 55,RelOp!x < 65")]
        [TestCase("RelOp!x > 65", "RelOp!x < 55", "Value!0, RelOp!x > 65,RelOp!x < 55")]
        [TestCase("RelOp!x <> 55", "RelOp!x <> 60,RelOp!x = 95", "Value!-1,RelOp!x <> 55,RelOp!x <> 60,RelOp!x = 95")]
        [TestCase("Value!0,RelOp!x > 55", "RelOp!x < 65,RelOp!x = 70", "RelOp!x > 55,RelOp!x < 65,Value!0,Value!-1")]
        [TestCase("Value!-1,RelOp!x = 55", "RelOp!x = 65,RelOp!x = 0", "RelOp!x = 55,RelOp!x = 65,Value!-1,Value!0")]
        [TestCase("RelOp!x <= 55", "RelOp!x > 65", "Value!0,RelOp!x <= 55,RelOp!x > 65")]
        [Category("Inspections")]
        public void ExpressionFilter_VariableRelationalOps(string firstCase, string secondCase, string expected)
        {
            var filter = ExpressionFilterFactory.Create(Tokens.Long);
            filter.AddComparablePredicateFilter("x", Tokens.Long);

            var expressions = RangeDescriptorsToExpressions(new string[] { firstCase, secondCase }, Tokens.Long);
            foreach (var expr in expressions)
            {
                filter.AddExpression(expr);
            }

            var expectedFilter = RangeDescriptorsToFilters(new string[] { expected }, Tokens.Long).First();

            Assert.AreEqual(expectedFilter, filter);
        }

        [TestCase("RelOp!x < 65.45", "RelOp!x < 55.97", "RelOp!x < 65.45,RelOp!x < 55.97")]
        [TestCase("RelOp!x > 55.97", "RelOp!x < 65.45", "Value!-1,RelOp!x > 55.97,RelOp!x < 65.45")]
        [TestCase("RelOp!x > 65.45", "RelOp!x < 55.97", "Value!0, RelOp!x > 65.45,RelOp!x < 55.97")]
        [TestCase("RelOp!x <> 55.97", "RelOp!x <> 60,RelOp!x = 95", "Value!-1,RelOp!x <> 55.97,RelOp!x <> 60,RelOp!x = 95")]
        [TestCase("Value!0,RelOp!x > 55.97", "RelOp!x < 65.45,RelOp!x = 70", "RelOp!x > 55.97,RelOp!x < 65.45,Value!0,Value!-1")]
        [TestCase("Value!-1,RelOp!x = 55.97", "RelOp!x = 65.45,RelOp!x = 0", "RelOp!x = 55.97,RelOp!x = 65.45,Value!-1,Value!0")]
        [TestCase("RelOp!x <= 55.97", "RelOp!x > 65.45", "Value!0,RelOp!x <= 55.97,RelOp!x > 65.45")]
        [Category("Inspections")]
        public void ExpressionFilter_VariableRelationalOpsDouble(string firstCase, string secondCase, string expected)
        {
            var filter = ExpressionFilterFactory.Create(Tokens.Single);
            filter.AddComparablePredicateFilter("x", Tokens.Single);

            var expressions = RangeDescriptorsToExpressions(new string[] { firstCase, secondCase }, Tokens.Single);
            foreach (var expr in expressions)
            {
                filter.AddExpression(expr);
            }

            var expectedFilter = RangeDescriptorsToFilters(new string[] { expected }, Tokens.Single).First();

            Assert.AreEqual(expectedFilter, filter);
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
        public void ExpressionFilter_AddRangeClauses(string firstCase, string selectExpressionTypename, string expectedRangeClauses)
        {
            var filter = ExpressionFilterFactory.Create(selectExpressionTypename);

            var clauses = RetrieveDelimitedElements(firstCase, RANGECLAUSE_DELIMITER);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IParseTreeValue start, out IParseTreeValue end, selectExpressionTypename, out string symbol);
                filter.AddExpression(new RangeOfValuesExpression(start, end));
            }

            var expected = RangeDescriptorsToFilters(new string[] { expectedRangeClauses }, selectExpressionTypename).First();
            Assert.AreEqual(expected, filter);
        }

        [TestCase("Is_>_50", "Long", "Max!50")]
        [TestCase("Is_>_50.49", "Long", "Max!50.49")]
        [TestCase("Is_>_50#", "Double", "Max!50")]
        [TestCase("Is_>_True", "Boolean", "RelOp!Is > True")]
        [TestCase("Is_>=_50", "Long", "Max!50,Value!50")]
        [TestCase("Is_>=_50.49", "Double", "Max!50.49,Value!50.49")]
        [TestCase("Is_>=_50#", "Double", "Max!50,Value!50")]
        [TestCase("Is_>=_True", "Boolean", "Value!True")]
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
        public void ExpressionFilter_AddIsClause(string firstCase, string selectExpressionTypename, string expectedRangeClauses)
        {
            var filter = ExpressionFilterFactory.Create(selectExpressionTypename);

            var clauses = RetrieveDelimitedElements(firstCase, RANGECLAUSE_DELIMITER);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IParseTreeValue start, out IParseTreeValue end, selectExpressionTypename, out string symbol);
                filter.AddExpression(new IsClauseExpression(end, symbol));
            }

            var expected = RangeDescriptorsToFilters(new string[] { expectedRangeClauses }, selectExpressionTypename).First();
            Assert.AreEqual(expected,filter);
        }

        [TestCase("Min!45.61", "Max!45.6", "Double")]
        [TestCase("Min!45.61,Max!60.5", "Range!39.2:65.1", "Double")]
        [TestCase("Range!True:False", "Value!50", "Boolean")]
        [TestCase("Value!-5000", "Value!False", "Boolean")]
        [TestCase("Value!True", "Value!0", "Boolean")]
        [TestCase("Value!500", "Value!0", "Boolean")]
        [TestCase("Min!5", "Max!-5000", "Long")]
        [TestCase("Min!40,Max!40", "Range!35:45", "Long")]
        [TestCase("Min!40,Max!44", "Range!35:45", "Long")]
        [TestCase("Min!40,Max!40", "Value!40", "Long")]
        [TestCase("Max!240,Range!150:239", "Value!240, Value!0,Value!1,Range!2:150", "Byte")]
        [TestCase("Range!151:255", "Value!150, Value!0,Value!1,Range!2:149", "Byte")]
        [TestCase("Min!13,Max!30,Range!12:100", "Value!13,Value!14,Value!15,Value!16,Value!17,Value!18,Range!12:30", "Long")]
        [Category("Inspections")]
        public void ExpressionFilter_FiltersAll(string firstCase, string secondCase, string SelectExpessionTypeName)
        {
            var filter = ExpressionFilterFactory.Create(SelectExpessionTypeName);
            var expressions = RangeDescriptorsToExpressions(new string[] { firstCase, secondCase }, SelectExpessionTypeName);
            foreach(var expr in expressions)
            {
                filter.AddExpression(expr);
            }
            Assert.IsTrue(filter.FiltersAllValues, filter.ToString());
        }

        [TestCase("RelOp!x < 3", "Is!Is < 3", "RelOp!x < 3,Min!3")]
        [TestCase("RelOp!!x", "Value!-x,RelOp!!x", "RelOp!!x,Value!-x")]
        [TestCase("RelOp!!x", "RelOp!Is < x", "RelOp!!x,RelOp!Is < x")]
        [TestCase("RelOp!Is < x,RelOp!Is < y", "RelOp!Is < x", "RelOp!Is < x,RelOp!Is < y")]
        [TestCase("Value!-x,Value!-y", "Value!-x", "Value!-x,Value!-y")]
        [TestCase("Range:3:55", "Value!x.Item(2)", "Range:3:55,Value!x.Item(2)")]
        [TestCase("Range!3:55", "Min!6", "Min!6,Range!6:55")]
        [TestCase("Range!3:55", "Max!6", "Max!6,Range!3:6")]
        [TestCase("Min!6", "Range!1:5", "Min!6")]
        [TestCase("Value!5,Value!6,Value!7", "Max!6", "Max!6,Value!5,Value!6")]
        [TestCase("Value!5,Value!6,Value!7", "Min!6", "Min!6,Value!6,Value!7")]
        [TestCase("Min!5,Max!75", "Value!85", "Min!5,Max!75")]
        [TestCase("Min!5,Max!75", "Value!0", "Min!5,Max!75")]
        [TestCase("Range!45:85", "Value!50", "Range!45:85")]
        [TestCase("Value!5,Value!6,Value!7,Value!8", "Range!6:8", "Range!6:8,Value!5")]
        [TestCase("Min!400,Range!15:160", "Range!500:505", "Min!400,Range!500:505")]
        [TestCase("Range!101:149", "Range!15:160", "Range!15:160")]
        [TestCase("Range!101:149", "Range!15:148", "Range!15:149")]
        [TestCase("Range!150:250,Range!1:100,Range!101:149", "Range!25:249", "Range!1:250")]
        [TestCase("Range!150:250,Range!1:100,Range!-5:-2,Range!101:149", "Range!25:249", "Range!-5:-2,Range!1:250")]
        [TestCase("Range!5:5,Value!x,Value!y", "", "Value!5,Value!x,Value!y")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersIntegers(string existing, string toAdd, string expectedClause)
        {
           (IExpressionFilter expected, IExpressionFilter actual) = TestAddFilters(new string[] { existing, toAdd, expectedClause }, Tokens.Long);
            Assert.IsTrue(actual.HasFilters && expected.HasFilters, "No filter content created");
            Assert.AreEqual(expected, actual);
        }

        [TestCase("Range!101.45:149.00007", "Range!101.57:110.63", "Range!101.45:149.00007")]
        [TestCase("Range!101.45:149.0007", "Range!15.67:148.9999", "Range!15.67:149.0007")]
        [TestCase("Range!101.45:149.2", "Range!149.2:150.5", "Range!101.45:150.5")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersRational(string firstCase, string secondCase, string expectedClauses)
        {
            (IExpressionFilter expected, IExpressionFilter actual) = TestAddFilters(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Double);
            Assert.IsTrue(actual.HasFilters && expected.HasFilters, "No filter content created");
            Assert.AreEqual(expected, actual);
        }

        [TestCase(@"Range!""Alpha"":""Omega""", @"Range!""Nuts"":""Soup""", @"Range!""Alpha"":""Soup""")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersStrings(string firstCase, string secondCase, string expectedClauses)
        {
            (IExpressionFilter expected, IExpressionFilter actual) = TestAddFilters(new string[] { firstCase, secondCase, expectedClauses }, Tokens.String);
            Assert.IsTrue(actual.HasFilters && expected.HasFilters, "No filter content created");
            Assert.AreEqual(expected, actual);
        }

        [TestCase("Range!0:10", "Value!50", "Value!True")]
        [TestCase("Range!False:True", "RelOp!x < 3", "RelOp!x < 3")]
        [TestCase(@"Range!""True:True""", "Value!True", "Value!True")]
        [TestCase(@"Range!""True:False""", "Value!True", "Value!False,Value!True")]
        [TestCase("Min!5", "RelOp!x < 5", "Value!False,RelOp!x < 5")]
        [TestCase("Value!-1,Value!0", "RelOp!x < 3", "Value!True,Value!False")]
        [TestCase("Range!-5:15", "RelOp!x < 3", "Value!True,RelOp!x < True")]
        [TestCase("Min!1", "RelOp!x < 3", "Min!1,RelOp!x < True")]
        [TestCase("Max!-2", "RelOp!x < 3", "Max!-2,RelOp!x < True")]
        [Category("Inspections")]
        public void ExpressionFilter_AddFiltersBoolean(string firstCase, string secondCase, string expectedClauses)
        {
            (IExpressionFilter expected, IExpressionFilter actual) = TestAddFilters(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Boolean);
            Assert.IsTrue(actual.HasFilters && expected.HasFilters, "No filter content created");
            Assert.AreEqual(expected, actual);
        }

        /*
         * The test cases below cover the truth table
         * for 'Is' clauses present in Boolean Select Case Statements.
         * Cases that always resolve to True (or False) are stored as Single values.
         * All others (outcome depends on the Select Case value) are 
         * stored as variable Predicate expressions.
        */

        [TestCase("Is_<_True", "Value!False")] //Always False
        [TestCase("Is_<=_True", "RelOp!Is <= True")]
        [TestCase("Is_>_True", "RelOp!Is > True")]
        [TestCase("Is_>=_True", "Value!True")] //Always True
        [TestCase("Is_=_True", "RelOp!Is = True")]
        [TestCase("Is_<>_True", "RelOp!Is <> True")]
        [TestCase("Is_>_False", "Value!False")] //Alsways False
        [TestCase("Is_>=_False", "RelOp!Is >= False")]
        [TestCase("Is_<_False", "RelOp!Is < False")]
        [TestCase("Is_<=_False", "Value!True")]    //Always True
        [TestCase("Is_=_False", "RelOp!Is = False")]
        [TestCase("Is_<>_False", "RelOp!Is <> False")]
        [Category("Inspections")]
        public void ExpressionFilter_BooleanIsClauseTruthTable(string rangeClause, string expected)
        {
            var filter = ExpressionFilterFactory.Create(Tokens.Boolean);

            var clauses = RetrieveDelimitedElements(rangeClause, RANGECLAUSE_DELIMITER);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IParseTreeValue start, out IParseTreeValue end, Tokens.Boolean, out string symbol);
                filter.AddExpression(new IsClauseExpression(end, symbol));
            }

            clauses = RetrieveDelimitedElements(expected, RANGECLAUSE_DELIMITER);
            var expectedFilter = CreateTestFilter(clauses.ToList(), Tokens.Boolean);
            Assert.AreEqual(expectedFilter, filter);
        }

        [TestCase("x_Like_*Bar", "RelOp!x Like *Bar")]
        [TestCase("x_Like_[A-Z]*", "RelOp!x Like [A-Z]*")]
        [TestCase("x_Like_[A-Z]*, x_Like_Fo*oBar", "RelOp!x Like [A-Z]*,RelOp!x Like Fo*oBar")]
        [Category("Inspections")]
        public void ExpressionFilter_LikeAddToFilter(string rangeClause, string expectedClause)
        {
            var filter = ExpressionFilterFactory.Create(Tokens.String);

            var clauses = RetrieveDelimitedElements(rangeClause, RANGECLAUSE_DELIMITER);
            foreach (var clause in clauses)
            {
                GetBinaryOpValues(clause, out IParseTreeValue start, out IParseTreeValue end, Tokens.String, out string symbol);
                filter.AddExpression(new BinaryExpression(start, end, symbol));
            }

            clauses = RetrieveDelimitedElements(expectedClause, RANGECLAUSE_DELIMITER);
            var expectedFilter = CreateTestFilter(clauses.ToList(), Tokens.String);
            Assert.AreEqual(expectedFilter, filter);
        }

        [TestCase("RelOp!x Like *Bar", "RelOp!x Like *Bar", "RelOp!x Like *Bar")]
        [TestCase("RelOp!x Like *Bar", "Value!True", "RelOp!x Like *Bar,Value!True")]
        [TestCase("RelOp!x Like *", "Value!True", "RelOp!x Like *")]
        [Category("Inspections")]
        public void ExpressionFilter_LikeFiltersDuplicates(string firstCase, string secondCase, string expectedClauses)
        {
            (IExpressionFilter expected, IExpressionFilter actual) = TestAddFilters(new string[] { firstCase, secondCase, expectedClauses }, Tokens.Boolean);
            Assert.IsTrue(actual.HasFilters && expected.HasFilters, "No filter content created");
            Assert.AreEqual(expected, actual);
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
//TODO: Take a hard look at the helpers...overlap? redundancy?
        private void GetBinaryOpValues(string operands, out IParseTreeValue LHS, out IParseTreeValue RHS, string selectExpressionType, out string opSymbol)
        {
            var operandItems = RetrieveDelimitedElements(operands, OPERAND_DELIMITER);

            LHS = CreateInspValueFrom(operandItems[0], conformTo: selectExpressionType);
            opSymbol = operandItems[1];
            RHS = CreateInspValueFrom(operandItems[2], conformTo: selectExpressionType);
        }

        private IParseTreeValue CreateInspValueFrom(string valAndType, string conformTo = null)
        {
            if (valAndType.Contains(VALUE_TYPE_DELIMITER))
            {
                var args = RetrieveDelimitedElements(valAndType, VALUE_TYPE_DELIMITER);
                var value = args[0];
                string declaredType = args[1].Equals(string.Empty) ? null : args[1];
                if (conformTo is null)
                {
                    if (declaredType is null)
                    {
                        return ValueFactory.Create(value);
                    }
                    var ptValue = ValueFactory.Create(value, declaredType);
                    ptValue.ParsesToConstantValue = true;
                    return ptValue;
                }
                else
                {
                    if (declaredType is null)
                    {
                        return ValueFactory.Create(value, conformToTypeName: conformTo);
                    }
                    var ptValue = ValueFactory.Create(value, declaredType, conformTo);
                    return ptValue;
                }
            }
            return conformTo is null ? ValueFactory.Create(valAndType)
                : ValueFactory.Create(valAndType, conformToTypeName: conformTo);
        }

        private (IExpressionFilter expected, IExpressionFilter actual) TestAddFilters(string[] inputs, string typeName)
        {
            Assert.IsTrue(inputs.Count() >= 2, "At least two rangeClase input strings are neede for this test");

            IExpressionFilter filter = ExpressionFilterFactory.Create(typeName);
            var expressions = RangeDescriptorsToExpressions(inputs, typeName);
            for (var idx = 0; idx <= expressions.Count() - 2; idx++)
            {
                var expr = expressions[idx];
                filter.AddExpression(expr);
            }

            var sumClauses = RangeDescriptorsToFilters(inputs, typeName);
            var expected = sumClauses[sumClauses.Count - 1];
            return (expected, filter);
        }

        private List<IExpressionFilter> RangeDescriptorsToFilters(string[] input, string typeName)
        {
            var caseToRanges = CasesToRanges(input);
            var filters = new List<IExpressionFilter>();
            foreach (var id in caseToRanges)
            {
                var newFilter = CreateTestFilter(id.Value, typeName);
                filters.Add(newFilter);
            }
            return filters;
        }

        private List<IRangeClauseExpression> RangeDescriptorsToExpressions(string[] input, string typeName)
        {
            var caseToRanges = CasesToRanges(input);
            var results = new List<IRangeClauseExpression>();
            foreach (var id in caseToRanges)
            {
                var expressions = CreateTestExpressions(id.Value, typeName);
                results.AddRange(expressions);
            }
            return results;
        }

        private Dictionary<string, List<string>> CasesToRanges(string[] caseClauses)
        {
            var caseToRanges = new Dictionary<string, List<string>>();
            var idx = 0;
            foreach (var cc in caseClauses)
            {
                idx++;
                caseToRanges.Add($"{idx}{cc}", new List<string>());
                var rgs = RetrieveDelimitedElements(cc, RANGECLAUSE_DELIMITER);
                foreach (var rg in rgs)
                {
                    caseToRanges[$"{idx}{cc}"].Add(rg.Trim());
                }
            }
            return caseToRanges;
        }

        private IExpressionFilter CreateTestFilter(List<string> annotations, string conformToType = null)
        {
            var result = ExpressionFilterFactory.Create(conformToType);
            var expressions = CreateTestExpressions(annotations, conformToType);
            foreach (var expression in expressions)
            {
                result.AddExpression(expression);
            }
            return result;
        }

        private List<IRangeClauseExpression> CreateTestExpressions(List<string> annotations, string conformToType = null)
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
                        var uciVal = ValueFactory.Create(clauseExpression, conformToTypeName: conformToType);
                        results.Add(new IsClauseExpression(uciVal, LogicSymbols.LT));
                    }
                    else if (clauseType.Equals("Max"))
                    {
                        var uciVal = ValueFactory.Create(clauseExpression, conformToTypeName: conformToType);
                        results.Add(new IsClauseExpression(uciVal, LogicSymbols.GT));
                    }
                    else if (clauseType.Equals("Range"))
                    {
                        var startEnd = clauseExpression.Split(new string[] { RANGE_STARTEND_DELIMITER }, StringSplitOptions.None);
                        var testValStart = ValueFactory.Create(startEnd[0], conformToTypeName: conformToType);
                        var testValEnd = ValueFactory.Create(startEnd[1], conformToTypeName: conformToType);
                        results.Add(new RangeOfValuesExpression(testValStart, testValEnd));
                    }
                    else if (clauseType.Equals("Value"))
                    {
                        var testVal = ValueFactory.Create(clauseExpression, conformToTypeName: conformToType);
                        results.Add(new ValueExpression(testVal));
                    }
                    else if (clauseType.Equals("RelOp"))
                    {
                        string symbol = string.Empty;
                        TryExtractSymbol(item, out symbol);
                        var sides = clauseExpression.Split(new string[] { symbol }, StringSplitOptions.None);

                        if (sides.Count() == 2)
                        {
                            var lhs = ValueFactory.Create(sides[0].Trim(), conformToTypeName: conformToType);
                            var rhs = ValueFactory.Create(sides[1].Trim(), conformToTypeName: conformToType);
                            if (lhs.ValueText.Equals(Tokens.Is))
                            {
                                results.Add(new IsClauseExpression(rhs, symbol));
                            }
                            else
                            {
                                results.Add(new BinaryExpression(lhs, rhs, symbol));
                            }

                        }
                        else
                        {
                            var uciVal = ValueFactory.Create(clauseExpression, conformToTypeName: conformToType);
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
                            var lhs = ValueFactory.Create(sides[0].Trim(), conformToTypeName: conformToType);
                            var rhs = ValueFactory.Create(sides[1].Trim(), conformToTypeName: conformToType);
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
            foreach (var logicSymbol in LogicSymbols.LogicSymbolList)
            {
                if (item.Contains($" {logicSymbol} "))
                {
                    symbol = logicSymbol;
                    return true;
                }
            }
            return false;
        }
    }
}
