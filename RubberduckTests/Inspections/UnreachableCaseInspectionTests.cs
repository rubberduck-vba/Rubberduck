using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnreachableCaseInspectionTests
    {

        [TestCase(@"""105""", "105")]
        [TestCase("105", "105")]
        [TestCase("45.2", "45")]
        [TestCase("True", "-1")]
        [TestCase("False", "0")]
        [TestCase("32.000023@", "32")]
        [TestCase("32.000023!", "32")]
        [TestCase("32.500023#", "33")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_ValueConversions(string firstCase, string value)
        {
            long num = 0;
            var cNum = new UnreachableCaseInspectionValue(firstCase);
            if (cNum.AsLong().HasValue)
            {
                num = cNum.AsLong().Value;
            }
            else
            {
                Assert.IsTrue(false, "Unable to convert to Long");
            }
            Assert.AreEqual(num, long.Parse(value));
        }

        [TestCase("What@", "What")]
        [TestCase("What!", "What")]
        [TestCase("What#", "What")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_NonNumberConversions(string firstCase, string value)
        {
            var cNum = new UnreachableCaseInspectionValue(firstCase);
            Assert.IsFalse(cNum.AsLong().HasValue);
            Assert.AreEqual(cNum.AsString(),value);
        }

        [TestCase("z * b", "5 To 10", ExpectedResult = "Long")]
        [TestCase("z * c", "5 To 10", ExpectedResult = "Double")]
        [TestCase("CStr(z) & d", "5 To 10", ExpectedResult = "Long")]
        [TestCase("z & d", "d To dd", ExpectedResult = "String")]
        [TestCase(@"z & ""45""", "5 To 10", ExpectedResult = "Long")]
        [TestCase("VBA.Rnd() > 0.5", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("True", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("z And True", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("z And j > 0.00", "5 To 10", ExpectedResult = "Boolean")]
        [TestCase("TestValueLong()", "5 To 10", ExpectedResult = "Long")]
        [TestCase("v", "b To bb", ExpectedResult = "Long")]
        [TestCase("v", "5 To 100", ExpectedResult = "Long")]
        [TestCase("v", "5 To 45.6", ExpectedResult = "Double")]
        [TestCase("v", "CLng(j) * bb * VBA.Rnd() + b ^ 4.5", ExpectedResult = "")]
        [TestCase("v", @"Len(""Whaaaat"") * bb * VBA.Rnd() + b ^ 4.5", ExpectedResult = "")]
        [TestCase("hint&", "5.0 To 45.6", ExpectedResult = "Long")]
        [TestCase("Sunday", "5.0 To 45.6", ExpectedResult = "Long")]
        [Category("Inspections")]
        public string UnreachableCaseInspection_DetermineSelectCaseType(string selectExpr, string firstCaseExpr)
        {
            string inputCode =
@"
        private Enum Weekday
            Sunday = 1
            Monday = 2
            Tuesday = 3
            Wednesday = 4
            Thursday = 5
            Friday = 6
            Saturday = 7
        End Enum
        Private const b As Long = 55
        Private const bb As Long = 100
        Private const c As Double = 0.0023
        Private const cc As Double = 0.509
        Private const d As String = ""Bar""
        Private const dd As String = ""Foo""

        Private Function TestValueLong() As Long
            TestValueLong = 5
        End Function

        Sub Foo(z As Long, j As Double, m As String, v as Variant, w as Boolean)

        Dim hint&
        hint& = 25

        Select Case <selectExpr>
          Case <firstCaseExpr>
            'OK
          Case Else
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<selectExpr>", selectExpr);
            inputCode = inputCode.Replace("<firstCaseExpr>", firstCaseExpr);
            UnreachableCaseInspection inspection;
            var TestTuple = GetInspectionAndParseTree(inputCode);
            inspection = TestTuple.Item1;
            var ctxt = GetDescendent<VBAParser.SelectCaseStmtContext>(TestTuple.Item2);
            var listener = inspection.Listener;
            var result = inspection.GetSelectCaseEvaluationType(ctxt);
            return result;
        }

        [Test]
        [Category("Inspections")]
        public void UnreachableCaseInspection_StringsInTheMixConvertable()
        {
            const string inputCode =
@"

        Sub Foo(z As Long, s As String)

        Select Case z + CLng(s)
            Case ""105""
            'OK
            Case 55
            'Unreachable
            Case 55
            'Unreachable
            Case Else
            'OK
        End Select

        End Sub";

            UnreachableCaseInspection inspection;
            var TestTuple = GetInspectionAndParseTree(inputCode);
            inspection = TestTuple.Item1;
            var ctxt = GetDescendent<VBAParser.SelectCaseStmtContext>(TestTuple.Item2);
            var listener = inspection.Listener;
            var result = inspection.GetSelectCaseEvaluationType(ctxt);
            Assert.AreEqual(Tokens.Long, result);
        }

        [TestCase(@"""105""")]
        [TestCase("105")]
        [TestCase("45.2")]
        [TestCase("True")]
        [TestCase("32.000023@")]
        [TestCase("32.000023!")]
        [TestCase("32.000023#")]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageFromRangeClause(string firstCase)
        {
            string inputCode =
@"
        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            long num = 0;
            if (UnreachableCaseInspectionValue.IsStringConstant(firstCase))
            {
                num = Convert.ToInt64(firstCase.Replace("\"", ""));
            }
            else
            {
                var cNum = new UnreachableCaseInspectionValue(firstCase);
                if (cNum.AsLong().HasValue)
                {
                    num = cNum.AsLong().Value;
                }
                else
                {
                    Assert.IsTrue(false, "Unable to convert to Long");
                }
            }
            inputCode = inputCode.Replace("<firstCase>", firstCase);
            UnreachableCaseInspection inspection;
            var TestTuple = GetInspectionAndParseTree(inputCode);
            inspection = TestTuple.Item1;
            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(TestTuple.Item2);
            var listener = inspection.Listener;
            var result = inspection.GetCoverage(ctxt, Tokens.Long);
            Assert.AreEqual(true, result.Summary.SingleValues.Any());
            Assert.AreEqual(num, result.Summary.SingleValues.First().AsLong());
        }

        [TestCase("50 To 100", 50, 100)]
        [TestCase("fromVal To toVal", 50, 100)]
        [TestCase("50.25 To 100.49", 50, 100)]
        [TestCase("True To False", -1, 0)]
        [TestCase("False To True", -1, 0)]
        [TestCase(@"""50"" To ""100""", 50, 100)]
        [TestCase("100 To 50", 50, 100)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageRanges(string firstCase, long start, long end)
        {
            string inputCode =
@"
        Private Const fromVal As long = 50
        Private Const toVal As Long = 100

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.Ranges.Any());
            var startVal = result.Summary.Ranges.First().Item1.AsLong().Value;
            var endVal = result.Summary.Ranges.First().Item2.AsLong().Value;
            Assert.AreEqual(start, startVal);
            Assert.AreEqual(end, endVal);
        }

        [TestCase("Is < 100", 100)]
        [TestCase("Is < 100.49", 100)]
        [TestCase("Is < 100#", 100)]
        [TestCase("Is < True", -1)]
        [TestCase(@"Is < ""100""", 100)]
        [TestCase("Is < toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsLTClause(string firstCase, long isLTMax)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsLT != null);
            var IsLTMax = result.Summary.IsLT.AsLong().Value;
            Assert.AreEqual(isLTMax, IsLTMax);
        }

        [TestCase("Is < x", 0)]
        [TestCase("z < x", 0)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsLTClauseVariable(string firstCase, long isLTMax)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long, x As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsLT == null);
        }

        [TestCase("Is < 45", "Is < 100", 100)]
        [TestCase("Is < 100", "Is < 45", 100)]
        [TestCase("Is > 45", "Is > 100", 45)]
        [TestCase("Is > 100", "Is > 45", 45)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsLTClauses(string firstCase, string secondCase, long isMaxMin)
        {
            string inputCode =
@"
        Private Const fromVal As long = 50
        Private Const toVal As Long = 100

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
            Case <secondCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            inputCode = inputCode.Replace("<secondCase>", secondCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var selectCaseStmtctxt = GetDescendent<VBAParser.SelectCaseStmtContext>(parseTree);
            var ctxts = selectCaseStmtctxt.GetDescendents<VBAParser.RangeClauseContext>();
            SummaryCoverage result = null;
            foreach( var ctxt in ctxts)
            {
                result = inspection.GetCoverage(ctxt, Tokens.Long, result);
            }

            if(result.Summary.IsLT != null)
            {
                var IsLTMax = result.Summary.IsLT.AsLong().Value;
                Assert.AreEqual(isMaxMin, IsLTMax);
            }
            else
            {
                var IsGTMin = result.Summary.IsGT.AsLong().Value;
                Assert.AreEqual(isMaxMin, IsGTMin);
            }
        }

        [TestCase("Is <= 100", 100)]
        [TestCase("Is <= 100.49", 100)]
        [TestCase("Is <= 100#", 100)]
        [TestCase("Is <= True", -1)]
        [TestCase(@"Is <= ""100""", 100)]
        [TestCase("Is <= toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsLTEClause(string firstCase, long isLTMax)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsLT != null);
            var IsLTMax = result.Summary.IsLT.AsLong().Value;
            Assert.AreEqual(isLTMax, IsLTMax);
            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(isLTMax, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("Is > 100", 100)]
        [TestCase("Is > 100.49", 100)]
        [TestCase("Is > 100#", 100)]
        [TestCase("Is > True", -1)]
        [TestCase(@"Is > ""100""", 100)]
        [TestCase("Is > toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsGTClause(string firstCase, long isGTMin)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsGT != null);
            var IsGTMin = result.Summary.IsGT.AsLong().Value;
            Assert.AreEqual(isGTMin, IsGTMin);
        }

        [TestCase("Is >= 100", 100)]
        [TestCase("Is >= 100.49", 100)]
        [TestCase("Is >= 100#", 100)]
        [TestCase("Is >= True", -1)]
        [TestCase(@"Is >= ""100""", 100)]
        [TestCase("Is >= toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsGTEClause(string firstCase, long isGTMin)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsGT != null);
            var IsGTMin = result.Summary.IsGT.AsLong().Value;
            Assert.AreEqual(isGTMin, IsGTMin);
            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(isGTMin, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("Is = 100", 100)]
        [TestCase("Is = 100.49", 100)]
        [TestCase("Is = 100#", 100)]
        [TestCase("Is = True", -1)]
        [TestCase(@"Is = ""100""", 100)]
        [TestCase("Is = toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsEQClause(string firstCase, long isGTMin)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(isGTMin, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("Is <> 100", 100)]
        [TestCase("Is <> 100.49", 100)]
        [TestCase("Is <> 100#", 100)]
        [TestCase("Is <> True", -1)]
        [TestCase(@"Is <> ""100""", 100)]
        [TestCase("Is <> toVal", 1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageIsNEQClause(string firstCase, long isNEQ)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.IsGT != null);
            var IsGTMin = result.Summary.IsGT.AsLong().Value;
            Assert.AreEqual(isNEQ, IsGTMin);
            Assert.AreEqual(true, result.Summary.IsLT != null);
            var IsLTMax = result.Summary.IsLT.AsLong().Value;
            Assert.AreEqual(isNEQ, IsLTMax);
        }

        [TestCase("z < 100", 100, true)]
        [TestCase("z <= 101", 101, true)]
        [TestCase("100 > z", 100, true)]
        [TestCase("101 >= z", 101, true)]
        [TestCase("z > 300", 300, false)]
        [TestCase("300 < z", 300, false)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageRelationalOp(string firstCase, long valToCheck, bool checkLT)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            if (checkLT)
            {
                Assert.AreEqual(true, result.Summary.IsLT != null);
                var IsLTMax = result.Summary.IsLT.AsLong().Value;
                Assert.AreEqual(valToCheck, IsLTMax);
            }
            else
            {
                Assert.AreEqual(true, result.Summary.IsGT != null);
                var IsGTMin = result.Summary.IsGT.AsLong().Value;
                Assert.AreEqual(valToCheck, IsGTMin);
            }
        }

        [TestCase("50 * 5", 250)]
        [TestCase("8 / 2", 4)]
        [TestCase("toVal / fromVal", 2)]
        [TestCase("toVal + fromVal", 1500)]
        [TestCase("fromVal - toVal", -500)]
        [TestCase("toVal * True + fromVal / 2", -750)]
        [TestCase("2 ^ 3", 8)]
        [TestCase("9 Mod 4", 1)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageBinaryMathOps(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(target, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("fromVal > 5 And toVal > 20", -1)]
        [TestCase("fromVal > 500000 Or toVal > 20000000", 0)]
        [TestCase("True Xor True", 0)]
        [TestCase("Not fromVal", 0)]
        //[TestCase("10 Eqv 8", -3)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageLogicOps(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(target, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("(fromVal - toVal) * 2", -1000)]
        [TestCase("(((((fromVal) - (toVal)) * (2))))", -1000)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageParentheses(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(target, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        [TestCase("-fromVal", -500)]
        [Category("Inspections")]
        public void UnreachableCaseInspection_SummaryCoverageUnaryMinus(string firstCase, long target)
        {
            string inputCode =
@"
        Private Const fromVal As long = 500
        Private Const toVal As Long = 1000

        Sub Foo(z As Long)

        Select Case z
            Case <firstCase>
            'OK
        End Select

        End Sub";

            inputCode = inputCode.Replace("<firstCase>", firstCase);
            var testTuple = GetInspectionAndParseTree(inputCode);
            var inspection = testTuple.Item1;
            var parseTree = testTuple.Item2;

            var ctxt = GetDescendent<VBAParser.RangeClauseContext>(parseTree);
            var result = inspection.GetCoverage(ctxt, Tokens.Long);

            Assert.AreEqual(true, result.Summary.SingleValues.Any(), "SingleValue not updated");
            Assert.AreEqual(target, result.Summary.SingleValues.First().AsLong(), "SingleValue has incorrect Value");
        }

        private Tuple<UnreachableCaseInspection, IParseTree> GetInspectionAndParseTree(string inputCode)
        {
            UnreachableCaseInspection inspection;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            string result = string.Empty;
            IParseTree parseTree;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                inspection = new UnreachableCaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                parseTree = state.ParseTrees.First().Value;
            }
            var ctxt = GetDescendent<VBAParser.SelectCaseStmtContext>(parseTree);
            var listener = inspection.Listener;
            return new Tuple<UnreachableCaseInspection, IParseTree>(inspection, parseTree);
        }

        private T GetDescendent<T>(IParseTree pt)
        {
            for(var idx = 0; idx < pt.ChildCount; idx++)
            {
                var child = pt.GetChild(idx);
                if(child is T)
                {
                    return (T)child;
                }
                if(child.ChildCount > 0)
                {
                    var result =  GetDescendent<T>(child);
                    if(result is T)
                    {
                        return result;
                    }
                }
            }
            return default;
        }

        //        [TestCase("String", @"""Foo""", @"""Bar""")]
        //        [TestCase("Long", "450000", "850000")]
        //        [TestCase("Integer", "4500", "8500")]
        //        [TestCase("Byte", "3", "254")]
        //        [TestCase("Double", "45000.345", "55000.25")]
        //        [TestCase("Single", "45.345", "55.25")]
        //        [TestCase("Currency", "4.34578", "5.25869")]
        //        [TestCase("Boolean", "True", "False")]
        //        [TestCase("Boolean", "55", "0")]
        //        //Negative values
        //        [TestCase("Long", "-450000", "850000")]
        //        [TestCase("Integer", "-4500", "8500")]
        //        [TestCase("Double", "-45000.345", "55000.25")]
        //        [TestCase("Single", "-45.345", "55.25")]
        //        [TestCase("Currency", "-4.34578", "5.25869")]
        //        [TestCase("Boolean", "-55", "0")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SingleUnreachableAllTypes(string type, string value1, string value2)
        //        {
        //            string inputCode =
        //@"Sub Test(x As <Type>)

        //Const firstVal As <Type> = <Value1>
        //Const secondVal As <Type> = <Value2>

        //Select Case x
        //    Case firstVal, secondVal
        //    'OK
        //    Case firstVal
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<Type>", type);
        //            inputCode = inputCode.Replace("<Value1>", value1);
        //            inputCode = inputCode.Replace("<Value2>", value2);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [TestCase("Long", "2147486648#", "-2147486649#")]
        //        [TestCase("Integer", "40000", "-50000")]
        //        [TestCase("Byte", "256", "-1")]
        //        [TestCase("Currency", "922337203685490.5808", "-922337203685477.5809")]
        //        [TestCase("Single", "3402824E38", "-3402824E38")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_ExceedsLimits(string type, string value1, string value2)
        //        {
        //            string inputCode =
        //@"Sub Foo(x As <Type>)

        //Const firstVal As <Type> = <Value1>
        //Const secondVal As <Type> = <Value2>

        //Select Case x
        //    Case firstVal
        //    'Unreachable
        //    Case secondVal
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<Type>", type);
        //            inputCode = inputCode.Replace("<Value1>", value1);
        //            inputCode = inputCode.Replace("<Value2>", value2);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [TestCase("x Or x < 5")]
        //        [TestCase("x = 1 Xor x < 5")]
        //        [TestCase("x And x < 5")]
        //        [TestCase("x Eqv 1")]
        //        [TestCase("Not x")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LogicalOpSelectCase(string booleanOp)
        //        {
        //            string inputCode =
        //@"Sub Foo(x As Long)
        //Select Case <boolOp>
        //    Case True
        //    'OK
        //    Case False 
        //    'OK
        //    Case -5
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<boolOp>", booleanOp);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [TestCase("Is > 8", "12", "9")]
        //        [TestCase("Is >= 8", "12", "8")]
        //        [TestCase("Is < 8", "-56", "7")]
        //        [TestCase("Is <= 8", "-56", "8")]
        //        [TestCase("Is <> 8", "-56", "5000")]
        //        [TestCase("Is = 8", "16 / 2", "4 * 2")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStmt(string isStmt, string unreachableValue1, string unreachableValue2)
        //        {
        //            string inputCode =
        //@"Sub Foo(z As Long)

        //Select Case z
        //    Case <IsStmt>
        //    'OK
        //    Case <Unreachable1>
        //    'Unreachable
        //    Case <Unreachable2>
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<IsStmt>", isStmt);
        //            inputCode = inputCode.Replace("<Unreachable1>", unreachableValue1);
        //            inputCode = inputCode.Replace("<Unreachable2>", unreachableValue2);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [TestCase("Dim Hint$\r\nSelect Case Hint$", @"""Here"" To ""Eternity""", @"""Forever""")] //String
        //        [TestCase("Dim Hint#\r\nHint#= 1.0\r\nSelect Case Hint#", "10.00 To 30.00", "20.00")] //Double
        //        [TestCase("Dim Hint!\r\nHint! = 1.0\r\nSelect Case Hint!", "10.00 To 30.00", "20.00")] //Single
        //        [TestCase("Dim Hint%\r\nHint% = 1\r\nSelect Case Hint%", "10 To 30", "20")] //Integer
        //        [TestCase("Dim Hint&\r\nHint& = 1\r\nSelect Case Hint&", "1000 To 3000", "2000")] //Long
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_TypeHint(string typeHintExpr, string firstCase, string secondCase)
        //        {
        //            string inputCode =
        //@"
        //Sub Foo()

        //<typeHintExprAndSelectCase>
        //    Case <firstCaseVal>
        //    'OK
        //    Case <secondCaseVal>
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<typeHintExprAndSelectCase>", typeHintExpr);
        //            inputCode = inputCode.Replace("<firstCaseVal>", firstCase);
        //            inputCode = inputCode.Replace("<secondCaseVal>", secondCase);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [TestCase("Long", "Is < 5", "x > -5000")]
        //        [TestCase("Long", "Is <> 4", "4")]
        //        [TestCase("Long", "Is <> -4", "4 - 8")]
        //        [TestCase("Long", "x > -5000", "Is < 1")]
        //        [TestCase("Long", "-5000 < x", "Is < 1")]
        //        [TestCase("Integer", "x <> 40", "35 To 45")]
        //        [TestCase("Double", "x > -5000.0", "Is < 1.7")]
        //        [TestCase("Single", "x > -5000.0", "Is < 1.7")]
        //        [TestCase("Currency", "x > -5000.0", "Is < 1.7")]
        //        [TestCase("Boolean", "-5000", "False")]
        //        [TestCase("Boolean", "True", "0")]
        //        [TestCase("Boolean", "50", "0")]
        //        [TestCase("Boolean", "Is > -1", "-10")]
        //        [TestCase("Boolean", "Is < -100", "Is > -10")]
        //        [TestCase("Boolean", "Is < 0", "0")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CoversAll(string type, string firstCase, string secondCase)
        //        {
        //            string inputCode =
        //@"Sub Foo(x As <Type>)

        //Select Case x
        //    Case <firstCase>
        //    'OK
        //    Case <secondCase>
        //    'Unreachable
        //    Case 45 * 12
        //    'Unreachable
        //    Case 500 To 700
        //    'Unreachable
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<Type>", type);
        //            inputCode = inputCode.Replace("<firstCase>", firstCase);
        //            inputCode = inputCode.Replace("<secondCase>", secondCase);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2, caseElse: 1);
        //        }

        //        [TestCase("0 To 10")]
        //        [TestCase("Is < 1")]
        //        [TestCase("-10 To 5")]
        //        [TestCase("5 To -10")]
        //        [TestCase("True To False")]
        //        [TestCase("False To True")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_BooleanSingleStmtCoversAll(string firstCase)
        //        {
        //            string inputCode =
        //@"Sub Foo(x As Boolean)

        //Select Case x
        //    Case <firstCase>
        //    'OK
        //    Case False
        //    'unreachable
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<firstCase>", firstCase);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        //        }

        //        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 2 = 49, x ^ 3 = 8")]
        //        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30")]
        //        [TestCase("x ^ 2 = 49, (CLng(VBA.Rnd() * 100) * x) < 30", "(CLng(VBA.Rnd() * 100) * x) < 30, x ^ 2 = 49")]
        //        [TestCase("x ^ 2 = 49, x ^ 3 = 8", "x ^ 3 = 8")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NoInspectionTextCompareOnly(string complexClause1, string complexClause2)
        //        {
        //            string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case <complexClause1>
        //    'OK
        //    Case <complexClause2>
        //    'Unreachable - detected by text compare of range clause(s)
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<complexClause1>", complexClause1);
        //            inputCode = inputCode.Replace("<complexClause2>", complexClause2);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [TestCase("Long", "5000 - 1000", "4000")]
        //        [TestCase("Double", "50.00 - 10.00", "40.00")]
        //        [TestCase("Currency", "50.00 - 10.00", "40.00")]
        //        [TestCase("Single", "50.00 - 10.00", "40.00")]
        //        [TestCase("Long", "5000 + 1000", "6000")]
        //        [TestCase("Double", "50.00 + 10.00", "60.00")]
        //        [TestCase("Single", "50.00 + 10.00", "60.00")]
        //        [TestCase("Long", "50 * 10", "500")]
        //        [TestCase("Double", "50.00 * 10.00", "500.00")]
        //        [TestCase("Single", "50.00 * 10.00", "500.00")]
        //        [TestCase("Long", "5000 / 1000", "5")]
        //        [TestCase("Double", "50.00 / 10.00", "5.0")]
        //        [TestCase("Currency", "50.00 / 10.00", "5.0")]
        //        [TestCase("Single", "50.00 / 10.00", "5.0")]
        //        [TestCase("Single", "52.00 Mod 10.00", "2.0")]
        //        [TestCase("Single", "2.00 ^ 3.00", "8.0")]
        //        [TestCase("Integer", "58 Mod 4", "2")]
        //        [TestCase("Integer", "2 ^ 3", "8")]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseClauseHasBinaryMathOp(string type, string mathOp, string unreachable)
        //        {
        //            string inputCode =
        //@"
        //Sub Foo(z As <Type>)

        //Select Case z
        //    Case <mathOp>
        //    'OK
        //    Case <unreachable>
        //    'Unreachable
        //End Select

        //End Sub";
        //            inputCode = inputCode.Replace("<Type>", type);
        //            inputCode = inputCode.Replace("<mathOp>", mathOp);
        //            inputCode = inputCode.Replace("<unreachable>", unreachable);
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_PowOpEvaluationAlgebraNoDetection()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case x ^ 2 = 49
        //    'OK
        //    Case x = 7
        //    'Unreachable, but not detected - math/algebra on the Select Case variable yet to be supported
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NumberRangeConstants()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Const JAN As Long = 1
        //Const DEC As Long = 12
        //Const AUG As Long = 8

        //Select Case JAN * x
        //    Case JAN To DEC
        //    'OK
        //    Case AUG
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NumberRangeMixedTypes()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Select Case x
        //    Case 1 To ""Forever""
        //    'Mismatch - unreachable
        //    Case 1 To 50
        //    'OK
        //    Case 45
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, mismatch: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NumberRangeCummulativeCoverage()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Select Case x
        //    Case 150 To 250
        //    'OK
        //    Case 1 To 100
        //    'OK
        //    Case 101 To 149
        //    'OK
        //    Case 25 To 249 
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NumberRangeHighToLow()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Select Case x
        //    Case 100 To 1
        //    'OK
        //    Case 50
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseElseIsClausePlusRange()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Select Case x
        //    Case Is > 200
        //    'OK
        //    Case 50 To 200
        //    'OK
        //    Case Is < 50
        //    'OK
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseElseIsClausePlusRangeAndSingles()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x as Long)

        //Select Case x
        //    Case 53,54
        //    'OK
        //    Case Is > 200
        //    'OK
        //    Case 55 To 200
        //    'OK
        //    Case Is < 50
        //    'OK
        //    Case 50,51,52
        //    'OK
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NestedSelectCase()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long, z As Long) 

        //Select Case x
        //    Case 1 To 10
        //    'OK
        //    Case 9
        //    'Unreachable
        //    Case 11
        //    Select Case  z
        //        Case 5 To 25
        //        'OK
        //        Case 6
        //        'Unreachable
        //        Case 8
        //        'Unreachable
        //        Case 15
        //        'Unreachable
        //    End Select
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 4);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NestedSelectCases()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As String, z As String )

        //Select Case x
        //    Case ""Foo"", ""Bar"", ""Goo""
        //    'OK
        //    Case ""Foo""
        //    'Unreachable
        //    Case ""Food""
        //    Select Case  z
        //        Case ""Food"", ""Bard"",""Good""
        //        'OK
        //        Case ""Bar""
        //        'OK
        //        Case ""Foo""
        //        'OK
        //        Case ""Goo""
        //        'OK
        //    End Select
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_NestedSelectCaseSUnreachable()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As String, z As String)

        //'Const x As String = ""Foo""
        //'Const z As String = ""Bar""

        //Select Case x
        //    Case ""Foo"", ""Bar""
        //    'OK
        //    Case ""Foo""
        //    'Unreachable
        //    Case ""Food""
        //    Select Case  z
        //        Case ""Foo"", ""Bar"",""Goo""
        //        'OK
        //        Case ""Bar""
        //        'Unreachable
        //        Case ""Foo""
        //        'Unreachable
        //        Case ""Goo""
        //        'Unreachable
        //    End Select
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 4);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SimpleLongCollisionConstantEvaluation()
        //        {
        //            const string inputCode =
        //@"

        //private const BASE As Long = 10
        //private const MAX As Long = BASE ^ 2

        //Sub Foo(x As Long)

        //Select Case x
        //    Case 100
        //    'OK
        //    Case MAX 
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_MixedSelectCaseTypes()
        //        {
        //            const string inputCode =
        //@"

        //private const MAXValue As Long = 5
        //private const TwentyFiveCents As Double = .25
        //private const MINCoins As Long = 4

        //Sub Foo(numQuarters As Byte)

        //Select Case numQuarters * TwentyFiveCents
        //    Case 1.25 To 10.00
        //    'OK
        //    Case MAXValue 
        //    'Unreachable
        //    Case MINCoins * TwentyFiveCents
        //    'OK
        //    Case MINCoins * 2
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_ExceedsIntegerButIncludesAccessibleValues()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Integer)

        //Select Case x
        //    Case 10,11,12
        //    'OK
        //    Case 15, 40000
        //    'Exceeds Integer value - but other value makes case reachable....no Error
        //    Case x < 4
        //    'OK
        //    Case -50000
        //    'Exceeds Integer values
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IntegerWithDoubleValue()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Integer)

        //Select Case x
        //    Case Is < -50
        //    'OK
        //    Case 214.0
        //    'OK - ish
        //    Case -214#
        //    'unreachable
        //    Case 98
        //    'OK
        //    Case 5 To 25, 50, 80
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_VariantSelectCase()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Variant)

        //Select Case x
        //    Case .4 To .9
        //    'OK
        //    Case 0.23
        //    'OK
        //    Case 0.55
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_VariantSelectCaseInferFromConstant()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Variant)

        //private Const TheValue As Double = 45.678
        //private Const TheUnreachableValue As Long = 25

        //Select Case x
        //    Case TheValue * 2
        //    'OK
        //    Case 0 To TheValue
        //    'OK
        //    Case TheUnreachableValue
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_VariantSelectCaseInferFromConstant2()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Variant)

        //private Const TheValue As Double = 45.678
        //private Const TheUnreachableValue As Long = 77

        //Select Case x
        //    Case x > TheValue
        //    'OK
        //    Case 0 To TheValue - 20
        //    'OK
        //    Case TheUnreachableValue
        //    'Unreachable
        //    Case 55
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_BuiltInSelectCase()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Variant)

        //Select Case VBA.Rnd()
        //    Case .4 To .9
        //    'OK
        //    Case 0.23
        //    'OK
        //    Case 0.55
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_BooleanNEQ()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Boolean)

        //Select Case x
        //    Case True
        //    'OK
        //    Case x <> False
        //    'Unreachable
        //    Case 95
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LongCollisionIndeterminateCase()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Long, y As Double)

        //Select Case x
        //    Case x > -3000
        //    'OK
        //    Case x < y
        //    'OK - indeterminant
        //    Case 95
        //    'Unreachable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LongCollisionMultipleVariablesSameType()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Long, y As Long)

        //Select Case x * y
        //    Case x > -3000
        //    'OK
        //    Case y > -3000
        //    'OK
        //    Case x < y
        //    'OK - indeterminant
        //    Case 95
        //    'OK - this gives a false positive when evaluated as if 'x' or 'y' is the only select case variable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LongCollisionMultipleVariablesDifferentType()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Long, y As Double)

        //Select Case x * y
        //    Case x > -3000
        //    'OK
        //    Case y > -3000
        //    'OK
        //    Case x < y
        //    'OK - indeterminant
        //    Case 95
        //    'OK - this gives a false positive when evaluated as if 'x' or 'y' is the only select case variable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LongCollisionVariableAndConstantDifferentType()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Long)

        //private const y As Double = 0.5

        //Select Case x * y
        //    Case x > -3000
        //    'OK
        //    Case y > -3000
        //    'Unreachable
        //    Case x < y
        //    'OK - indeterminant
        //    Case 95
        //    'OK - this gives a false positive when evaluated as if 'x' is the only select case variable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_LongCollisionUnaryMathOperation()
        //        {
        //            const string inputCode =
        //@"Sub Foo( x As Long, y As Double)

        //Select Case -x  'math on the Select Case variable disqualifies inspection
        //    Case x > -3000
        //    'OK
        //    Case y > -3000
        //    'OK
        //    Case x < y
        //    'OK - indeterminant
        //    Case 95
        //    'unreachable - not evaluated
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_BooleanExpressionUnreachableCaseElseInvertBooleanRange()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Boolean)

        //Select Case VBA.Rnd() > 0.5
        //    Case False To True 
        //    'OK
        //    Case True
        //    'Unreachable
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_StringWhereLongShouldBe()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case 1 To 49
        //    'OK
        //    Case 50
        //    'OK
        //    Case ""Test""
        //    'Unreachable
        //    Case ""85""
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_MixedTypes()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case 1 To 49
        //    'OK
        //    Case ""Test"", 100, ""92""
        //    'OK - ""Test"" will not be evaluated
        //    Case ""85""
        //    'OK
        //    Case 2
        //    'Unreachable
        //    Case 92
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_StringWhereLongShouldBeIncludeLongAsString()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case 1 To 49
        //    'OK
        //    Case ""51""
        //    'OK
        //    Case ""Hello World""
        //    'Unreachable
        //    Case 50
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, mismatch: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_MultipleRanges()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case 1 To 4, 7 To 9, 11, 13, 15 To 20
        //    'OK
        //    Case 8
        //    'Unreachable
        //    Case 11
        //    'Unreachable
        //    Case 17
        //    'Unreachable
        //    Case 21
        //    'Reachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 3);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CascadingIsStatements()
        //        {
        //            const string inputCode =
        //@"Sub Foo(LNumber As Long)

        //Select Case LNumber
        //    Case Is < 100
        //        'OK
        //    Case Is < 200
        //        'OK
        //    Case Is < 300
        //        'OK
        //    Case Else
        //        'OK
        //    End Select
        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CascadingIsStatementsGT()
        //        {
        //            const string inputCode =
        //@"Sub Foo(LNumber As Long)

        //Select Case LNumber
        //    Case Is > 300
        //    'OK
        //    Case Is > 200
        //    'OK  
        //    Case Is > 100
        //    'OK  
        //    Case Else
        //    'OK
        //    End Select
        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStatementUnreachableGT()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case Is > 100
        //        'OK  
        //    Case Is > 200
        //        'unreachable  
        //    Case Is > 300
        //        'unreachable
        //    Case Else
        //        'OK
        //    End Select
        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStatementUnreachableLT()
        //        {
        //            const string inputCode =
        //@"Sub Foo(x As Long)

        //Select Case x
        //    Case Is < 300
        //        'OK  
        //    Case Is < 200
        //        'unreachable  
        //    Case Is < 100
        //        'unreachable
        //    Case Else
        //        'OK
        //    End Select
        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStmtToIsStmtCaseElseUnreachableUsingIs()
        //        {
        //            const string inputCode =
        //@"Sub Foo(z As Long)

        //Select Case z
        //    Case z <> 5 
        //    'OK
        //    Case Is = 5
        //    'OK
        //    Case 400
        //    'Unreachable
        //    Case Else
        //    'Unreachable
        //End Select
        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1,  caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseClauseHasParens()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Long)

        //private const maxValue As Long = 5000
        //private const subtract As Long = 2000

        //Select Case z
        //    Case (maxValue - subtract) * 10
        //    'OK
        //    Case 30000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseClauseHasMultipleParens()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Long)

        //private const maxValue As Long = 5000
        //private const subtractValue As Long = 2000

        //Select Case z
        //    Case (maxValue - subtractValue) * (55 - 35) / 10
        //    'OK
        //    Case 6000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test] 
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SelectCaseHasMultOpWithFunction()
        //        {
        //            const string inputCode =
        //@"
        //Function Bar() As Long
        //    Bar = 5
        //End Function

        //Sub Foo(z As Long)

        //Select Case Bar() * z
        //    Case Is > 5000
        //    'OK
        //    Case 5000
        //    'OK
        //    Case 5001
        //    'Unreachable
        //    Case 10000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseClauseHasMultOpInParens()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Long)

        //private const maxValue As Long = 5000

        //Select Case (((z)))
        //    Case ((2 * maxValue))
        //    'OK
        //    Case 10000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseClauseHasMultOp2Constants()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Long)

        //private const maxValue As Long = 5000
        //private const minMultiplier As Long = 2

        //Select Case z
        //    Case maxValue / minMultiplier
        //    'OK
        //    Case 2500
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_EnumerationNumberRangeNoDetection()
        //        {
        //            const string inputCode =
        //@"
        //private Enum Weekday
        //    Sunday = 1
        //    Monday = 2
        //    Tuesday = 3
        //    Wednesday = 4
        //    Thursday = 5
        //    Friday = 6
        //    Saturday = 7
        //    End Enum

        //Sub Foo(z As Weekday)

        //Select Case z
        //    Case Weekday.Monday To Weekday.Saturday
        //    'OK
        //    Case z = Weekday.Tuesday
        //    'Unreachable
        //    Case Weekday.Wednesday
        //    'Unreachable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_EnumerationNumberRangeNonConstant()
        //        {
        //            const string inputCode =
        //@"
        //private Enum BitCountMaxValues
        //    max1Bit = 2 ^ 0
        //    max2Bits = 2 ^ 1 + max1Bit
        //    max3Bits = 2 ^ 2 + max2Bits
        //    max4Bits = 2 ^ 3 + max3Bits
        //End Enum

        //Sub Foo(z As BitCountMaxValues)

        //Select Case z
        //    Case 7
        //    'OK
        //    Case BitCountMaxValues.max3Bits
        //    'Unreachable
        //    Case BitCountMaxValues.max4Bits
        //    'OK
        //    Case 15
        //    'Unreachable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_EnumerationLongCollision()
        //        {
        //            const string inputCode =
        //@"
        //private Enum BitCountMaxValues
        //    max1Bit = 2 ^ 0
        //    max2Bits = 2 ^ 1 + max1Bit
        //    max3Bits = 2 ^ 2 + max2Bits
        //    max4Bits = 2 ^ 3 + max3Bits
        //End Enum

        //Sub Foo(z As BitCountMaxValues)

        //Select Case z
        //    Case BitCountMaxValues.max3Bits
        //    'OK
        //    Case 7
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_EnumerationNumberRangeConflicts()
        //        {
        //            const string inputCode =
        //@"
        //        private Enum Fruit
        //            Apple = 10
        //            Pear = 20
        //            Orange = 30
        //            End Enum

        //        Sub Foo(z As Fruit)

        //        Select Case z
        //            Case Apple
        //            'OK
        //            Case Pear 
        //            'OK     
        //            Case Orange        
        //            'OK
        //            Case Else
        //            'OK - avoid flagging CaseElse for enums so guard clauses such as below are retained
        //            Err.Raise 5, ""MyFunction"", ""Invalid value given for the enumeration.""
        //        End Select

        //        End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0, caseElse: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_EnumerationNumberCaseElse()
        //        {
        //            const string inputCode =
        //@"
        //        private Enum Fruit
        //            Apple = 10
        //            Pear = 20
        //            Orange = 30
        //            End Enum

        //        Sub Foo(z As Fruit)

        //        Select Case z
        //            Case z <> Apple
        //            'OK
        //            Case Apple 
        //            'OK     
        //            Case Else
        //            'unreachable - Guard clause will always be skipped
        //            Err.Raise 5, ""MyFunction"", ""Invalid value given for the enumeration.""
        //        End Select

        //        End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseElseByte()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Byte)

        //Select Case z
        //    Case z >= 2
        //    'OK
        //    Case 0,1
        //    'OK
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_CaseElseByteMultipleCases()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Byte)

        //Select Case z
        //    Case z >= 240
        //    'OK
        //    Case 0,1
        //    'OK
        //    Case Is < 100
        //    'OK
        //    Case 150 To 240
        //    'OK
        //    Case 100 To 228
        //    'OK
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_RangeCollisionsAggregateClauses()
        //        {
        //            const string inputCode =
        //@"
        //Sub Foo(z As Long)

        //Select Case z
        //    Case z > 30
        //    'OK
        //    Case 14,15,16,17,18,19 To 30
        //    'OK
        //    Case 30 To 100
        //    'Unreachable
        //    Case Is <= 13
        //    'OK   
        //    Case Else
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 1, caseElse: 1);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SelectCaseUsesConstantReferenceExpr()
        //        {
        //            const string inputCode =
        //@"
        //private Const maxValue As Long = 5000

        //Sub Foo(z As Long)

        //Select Case ( z * 3 ) - 2   'math on the Select Case variable disqualifies inspection
        //    Case z > maxValue
        //    'OK
        //    Case 15
        //    'OK
        //    Case 6000
        //    'Unreachable - not evaluated
        //    Case 8500
        //    'Unreachable - not evaluated
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 0);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SelectCaseUsesConstantIndeterminantExpression()
        //        {
        //            const string inputCode =
        //@"
        //private Const maxValue As Long = 5000

        //Sub Foo(z As Long)

        //Select Case z
        //    Case z > maxValue / 2
        //    'OK
        //    Case z > maxValue
        //    'Unreachable
        //    Case 15
        //    'OK
        //    Case 8500
        //    'Unreachable
        //    Case Else
        //    'OK
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SelectCaseIsFunction()
        //        {
        //            const string inputCode =
        //@"
        //Function Bar() As Long
        //    Bar = 5
        //End Function

        //Sub Foo()

        //Select Case Bar()
        //    Case Is > 5000
        //    'OK
        //    Case 5000
        //    'OK
        //    Case 5001
        //    'Unreachable
        //    Case 10000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_SelectCaseIsFunctionWithParams()
        //        {
        //            const string inputCode =
        //@"
        //Function Bar(x As Long, y As Double) As Long
        //    Bar = 5
        //End Function

        //Sub Foo(firstVar As Long, secondVar As Double)

        //Select Case Bar( firstVar, secondVar )
        //    Case Is > 5000
        //    'OK
        //    Case 5000
        //    'OK
        //    Case 5001
        //    'Unreachable
        //    Case 10000
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStmtAndNegativeRange()
        //        {
        //            const string inputCode =
        //@"Sub Foo(z As Long)

        //Select Case z
        //    Case Is < 8
        //    'OK
        //    Case -10 To -3
        //    'Unreachable
        //    Case 0
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        //        [Test]
        //        [Category("Inspections")]
        //        public void UnreachableCaseInspection_IsStmtAndNegativeRangeWithConstants()
        //        {
        //            const string inputCode =
        //@"
        //private const START As Long = 10
        //private const FINISH As Long = 3

        //Sub Foo(z As Long)
        //Select Case z
        //    Case Is < 8
        //    'OK
        //    Case -(START * 4) To -(FINISH * 2) 
        //    'Unreachable
        //    Case 0
        //    'Unreachable
        //End Select

        //End Sub";
        //            CheckActualResultsEqualsExpected(inputCode, unreachable: 2);
        //        }

        private void CheckActualResultsEqualsExpected(string inputCode, int unreachable = 0, int mismatch = 0, int caseElse = 0)
        {
            var expected = new Dictionary<string, int>
            {
                { InspectionsUI.UnreachableCaseInspection_Unreachable, unreachable },
                { InspectionsUI.UnreachableCaseInspection_TypeMismatch, mismatch },
                { InspectionsUI.UnreachableCaseInspection_CaseElse, caseElse },
            };

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            IEnumerable<Rubberduck.Parsing.Inspections.Abstract.IInspectionResult> actualResults;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnreachableCaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }
            var actualUnreachable = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_Unreachable));
            var actualMismatches = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_TypeMismatch));
            var actualUnreachableCaseElses = actualResults.Where(ar => ar.Description.Equals(InspectionsUI.UnreachableCaseInspection_CaseElse));

            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_Unreachable], actualUnreachable.Count(), "Unreachable result");
            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_TypeMismatch], actualMismatches.Count(), "Mismatch result");
            Assert.AreEqual(expected[InspectionsUI.UnreachableCaseInspection_CaseElse], actualUnreachableCaseElses.Count(), "CaseElse result");
        }
    }
}
