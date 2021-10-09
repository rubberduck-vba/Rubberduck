using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IIfSideEffectInspectionTests : InspectionTestsBase
    {
        [TestCase("Func1()", "Func2()", 2)]
        [TestCase("Func1()", "member2", 1)]
        [TestCase("member1", "Func2()", 1)]
        [TestCase("member1", "member2", 0)]
        [TestCase("Prop1()", "Prop2()", 2)]
        [TestCase("Prop1()", "member2", 1)]
        [TestCase("member1", "Prop2()", 1)]
        [Category("Inspections")]
        [Category("IIfSideEffect")]
        public void TypicalScenario(string secondArg, string thirdArg, long expected)
        {
            var inputcode =
$@"
Private member1 As Long
Private member2 As Long

Sub Foo(ByVal flag As Boolean)
    Dim result As Long
    result = IIf(flag, {secondArg}, {thirdArg})
End Sub

Private Function Func1() As Long
End Function

Private Function Func2() As Long
End Function

Private Property Get Prop1() As Long
End Property

Private Property Get Prop2() As Long
End Property
";
            Assert.AreEqual(expected, IIfInspectionResults(inputcode).Count());
        }

        [Test]
        [Category("Inspections")]
        [Category("IIfSideEffect")]
        public void FunctionIsUsedForFirstArgument()
        {
            var inputcode =
$@"

Private trueField As Long
Private falseField As Long

Sub Foo()
    Dim result As Long
    result = IIf(Func1(), trueField, falseField)
End Sub

Private Function Func1() As Boolean
End Function
";
            Assert.AreEqual(0, IIfInspectionResults(inputcode).Count());
        }

        [TestCase("IIf(TruePart:=Func1, Expression:=flag, FalsePart:=default)", 1)]
        [TestCase("IIf(flag, FalsePart:=default, TruePart:=Func1)", 1)]
        [Category("Inspections")]
        [Category("IIfSideEffect")]
        public void NamedArgumentsUsed(string iifExpression, long expected)
        {
            var inputcode =
$@"

Sub Foo(ByVal flag As Boolean, Optional ByVal default As Long = 0)
    Dim result As Long
    result = {iifExpression}
End Sub

Private Function Func1() As Boolean
End Function
";
            Assert.AreEqual(expected, IIfInspectionResults(inputcode).Count());
        }

        [TestCase(@"IIf(flag, CLng(CStr(Func1() + 10) & ""000""), default)", 3)]
        [TestCase(@"IIf(flag, default, CLng(CStr(Func1() + 10) & ""000""))", 3)]
        [TestCase(@"IIf(flag, default, CLng(Trim(CStr(Func1() + 10)) & ""000""))", 3)] //Trim is not flagged
        [TestCase(@"IIf(CLng(CStr(Func1() + 10) & ""000""), 2000, default, default)", 0)]
        [Category("Inspections")]
        [Category("IIfSideEffect")]
        public void NestedUseOfFunctions(string iifExpression, long expected)
        {
            var inputcode =
$@"

Sub Foo(ByVal flag As Boolean, Optional ByVal default As Long = 0)
    Dim result As Long
    result = {iifExpression}
End Sub

Private Function Func1() As Long
End Function
";
            Assert.AreEqual(expected, IIfInspectionResults(inputcode).Count());
        }

        [TestCase("IIf(IIf(Func1(), Func2, 185), 45, 65)", 1)]
        [TestCase("IIf(IIf(Func1(), 100, 185), 45, 65)", 0)]
        [TestCase("IIf(IIf(Func1() > Func2(), 100, 185), 45, 65)", 0)]
        [Category("Inspections")]
        [Category("IIfSideEffect")]
        public void NestedIIfUseOfFunctions(string iifExpression, long expected)
        {
            var inputcode =
$@"
Sub Foo(ByVal flag As Boolean, Optional ByVal default As Long = 0)
    Dim result As Long
    result = {iifExpression}
End Sub

Private Function Func1() As Long
End Function

Private Function Func2() As Long
End Function
";
            Assert.AreEqual(expected, IIfInspectionResults(inputcode).Count());
        }

        private IEnumerable<IInspectionResult> IIfInspectionResults(string inputCode)
        {
            return InspectionResultsForModules((MockVbeBuilder.TestModuleName, inputCode, ComponentType.StandardModule), ReferenceLibrary.VBA);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new IIfSideEffectInspection(state);
        }
    }
}
