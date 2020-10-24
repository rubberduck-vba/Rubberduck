using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime.Tree;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections.ThunderCode
{
    [TestFixture]
    [Category("Inspections")]
    [Category("ThunderCode")]
    public class ThunderCodeInspectionTests
    {
        [Test]
        [Category("Inspections")]
        [TestCase(1, "Public Sub foo\u00A0bar()" + @"
End Sub")]
        [TestCase(0, @"Public Sub foo()
End Sub")]
        public void NonBreakingSpaceIdentifier_ReturnsResult(int expectedCount, string inputCode)
        {
            var func = new Func<RubberduckParserState, IInspection>(state =>
                new NonBreakingSpaceIdentifierInspection(state));
            ThunderCatsGo(func, inputCode, expectedCount);
        }

        [Test]
        [Category("Inspections")]
        [TestCase(0, @"Public Sub foo bar()
End Sub")]
        public void NonBreakingSpaceIdentifier_IllegalInputCausesParserError(int expectedCount, string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                parser.Parse(new CancellationTokenSource());
                var actualStatus = state.Status;
                Assert.AreEqual(ParserState.Error, actualStatus);
            }
        }

        [Test]
        [TestCase(2, @"Do")]
        [TestCase(2, @"Loop")]
        [TestCase(2, @"For")]
        [TestCase(2, @"Next")]
        [TestCase(0, @"Foo")]
        [TestCase(0, @"Bar")]
        [TestCase(0, @"ForNext")]
        [TestCase(0, @"Fur")]
        public void KeywordsUsedAsMember_ReturnsResult(int expectedCount, string inputVariable)
        {
            var inputCode = $@"Private Type HeeHaw
  {inputVariable} As Variant
End Type

Private Enum HawHaw
  [{inputVariable}] = 1
End Enum";

            var func = new Func<RubberduckParserState, IInspection>(state =>
                new KeywordsUsedAsMemberInspection(state));
            ThunderCatsGo(func, inputCode, expectedCount);
        }

        [Test]
        [TestCase( @"[Do]")]
        [TestCase(@"Do ")]
        [TestCase(@" Do")]
        [TestCase(@"")]
        [TestCase(@"[]")]
        public void KeywordsUsedAsMember_CorrectlyDistinguishesBracketedIdentifiers(string inputVariable)
        {
            var inputCode = $@"
Private Enum HawHaw
  [{inputVariable}] = 1
End Enum";

            var func = new Func<RubberduckParserState, IInspection>(state =>
                new KeywordsUsedAsMemberInspection(state));
            ThunderCatsGo(func, inputCode, 0);
        }



        // NOTE: the inspection only covers trivial cases and is not exhaustive
        // For that reason, some of test cases which the evil continuations exists
        // may still pass without any results. To cover them all would likely be too
        // expensive.
        [Test]
        [TestCase(1, @"Private Sub Foo()
End _
Sub")]
        [TestCase(0, @"Private _
Sub Foo
End Sub")]
        [TestCase(1, @"Private Function Foo() As Boolean
End _
 _
Function")]
        [TestCase(0, @"Private _
 _
Function Foo() As Boolean
End Function")]
        [TestCase(1, @"Private Property _
Get Foo() As String
End Property")]
        [TestCase(0, @"Private Property Get Foo() As String
Foo = _
 23
End Property")]
        [TestCase(1, @"Private Property Get Foo() As String
Foo = 23
End _
Property")]
        [TestCase(0, @"Private Property Let Foo(NewValue As String)
  Bar = NewValue
End Property")]
        [TestCase(1, @"Private Property Let Foo(NewValue As String)
  Bar = NewValue
End _
Property")]
        [TestCase(2, @"Private Property _
Let Foo(NewValue As String)
  Bar = NewValue
End _
Property")]
        [TestCase(1, @"Option _
Explicit")]
        [TestCase(0, @"Option Explicit")]
        [TestCase(1, @"Option Explicit
Option _
Base 1")]
        [TestCase(0, @"Option Explicit
Option Base 1")]
        [TestCase(1, @"Public Type foo
  bar As Variant
End _
Type")]
        [TestCase(0, @"Public Type foo
  bar _ 
  As _
  Variant
End Type")]
        [TestCase(1, @"Public Enum foo
  bar = 1
End _
Enum", Ignore = "Enum does not allow line continuations")]
        [TestCase(0, @"Public Enum bar
  foo =  1
End Enum")]
        [TestCase(1, @"Private Sub Foo()
  On _
Error GoTo 0
End Sub")]
        [TestCase(0, @"Private Sub Foo()
  On Error GoTo 0
End Sub")]
        [TestCase(0, @"Private Sub Foo()
  On Local Error GoTo 0
End Sub")]
        [TestCase(1, @"Private Sub Foo()
  On _ 
Local Error GoTo 0
End Sub")]
        [TestCase(0, @"Private Sub Foo()
  For i = 0 to 3
    Exit For
  Next
End Sub")]
        [TestCase(1, @"Private Sub Foo()
  For i = 0 to 3
    Exit _ 
For
  Next
End Sub")]
        public void EvilLineContinuation_ReturnResults(int expectedCount, string inputCode)
        {
            var func = new Func<RubberduckParserState, IInspection>(state =>
                new LineContinuationBetweenKeywordsInspection(state));
            ThunderCatsGo(func, inputCode, expectedCount);
        }

        [Test]
        [TestCase(0,@"Public Sub Gogo()
GoTo 1
1:
End Sub")]
        [TestCase(0, @"Public Sub Gogo()
GoTo 1
1
End Sub")]
        [TestCase(1, @"Public Sub Gogo()
-1:
End Sub")]
        [TestCase(1, @"Public Sub Gogo()
-1
End Sub")]
        [TestCase(1, @"Public Sub Gogo()
GoTo 1
1
-1:
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
GoTo -1
1
-1:
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
GoTo -1
1:
-1
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
GoTo -5
1
-5:
End Sub")]
        [TestCase(1, @"Public Sub Gogo()
On Error GoTo 1
1
-1:
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
On Error GoTo -1
1
-1:
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
On Error GoTo -1
1:
-1
End Sub")]
        [TestCase(0, @"Public Sub Gogo()
On Error GoTo -1
1
End Sub")]
        [TestCase(1, @"Public Sub Gogo()
On Error GoTo -2
1
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
On Error GoTo -5
1
-5:
End Sub")]
        [TestCase(2, @"Public Sub Gogo()
On Error GoTo -5
1:
-5
End Sub")]
        public void NegativeLineNumberLabel_ReturnResults(int expectedCount, string inputCode)
        {
            var func = new Func<RubberduckParserState, IInspection>(state =>
                new NegativeLineNumberInspection(state));
            ThunderCatsGo(func, inputCode, expectedCount);
        }

        [Test]
        [TestCase(0, @"Public Sub Derp()
  On Error GoTo 1
  Exit Sub
1
  Debug.Print ""derp""
End Sub")]
        [TestCase(0, @"Public Sub Derp()
  On Error GoTo 1
  Exit Sub
1:
  Debug.Print ""derp""
End Sub")]
        [TestCase(0, @"Public Sub Derp()
  On Error GoTo 0
  Exit Sub
  Debug.Print ""derp""
End Sub")]
        [TestCase(1, @"Public Sub Derp()
  On Error GoTo -1
  Exit Sub
  Debug.Print ""derp""
End Sub")]
        [TestCase(0, @"Public Sub Derp()
  On Error GoTo -2
  Exit Sub
-2
  Debug.Print ""derp""
End Sub")]
        [TestCase(0, @"Public Sub Derp()
  On Error GoTo -2
  Exit Sub
-2:
  Debug.Print ""derp""
End Sub")]
        public void OnErrorGoToMinusOne_ReturnResults(int expectedCount, string inputCode)
        {
            var func = new Func<RubberduckParserState, IInspection>(state =>
                new OnErrorGoToMinusOneInspection(state));
            ThunderCatsGo(func, inputCode, expectedCount);
        }

        private static void ThunderCatsGo(Func<RubberduckParserState, IInspection> inspectionFunction, string inputCode, int expectedCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = inspectionFunction(state);
                var actualResults = InspectionResults(inspection, state);

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection parseTreeInspection)
            {
                WalkTrees(parseTreeInspection, state);
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }

        private static void WalkTrees(IParseTreeInspection inspection, RubberduckParserState state)
        {
            var codeKind = inspection.TargetKindOfCode;
            var listener = inspection.Listener;

            List<KeyValuePair<QualifiedModuleName, IParseTree>> trees;
            switch (codeKind)
            {
                case CodeKind.AttributesCode:
                    trees = state.AttributeParseTrees;
                    break;
                case CodeKind.CodePaneCode:
                    trees = state.ParseTrees;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(codeKind), codeKind, null);
            }

            foreach (var (module, tree) in trees)
            {
                listener.CurrentModuleName = module;
                ParseTreeWalker.Default.Walk(listener, tree);
            }
        }
    }
}
