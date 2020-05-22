using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyMethodInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void EmptyMethodBlock_InspectionName()
        {
            var inspection = new EmptyMethodInspection(null);

            Assert.AreEqual(nameof(EmptyMethodInspection), inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_DoesNotFireOnImplementedMethod()
        {
            const string inputCode =
                @"Sub Foo()
    MsgBox ""?""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_FiresOnNotImplementedMethod()
        {
            const string inputCode =
                @"Sub Foo()
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethodInterfaceImplementation_ReturnsResult()
        {
            const string interfaceCode =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string concreteCode =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore EmptyMethod
Sub Foo()
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Label:")]
        [TestCase("Const Bar = 42")]
        [TestCase("Dim bar As Long")]
        [TestCase("Const Bar = 42: Dim baz As Long")]
        [TestCase("Const Bar = 42\nDim baz As Long")]
        [TestCase("Label: Const Bar = 42: Dim baz As Long")]
        [TestCase("Label:\nConst Bar = 42\nDim baz As Long")]
        [TestCase("Rem This is a rem comment")]
        [TestCase("'This is a comment")]
        [TestCase("'@Ignore EmptyMethod")]
        [TestCase("")]
        public void EmptyMethod_VariousStatements_ReturnResult(string statement)
        {
            string inputCode =
                $@"Sub Foo()
{statement}
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Label: Foo")]
        [TestCase("Label:\nFoo")]
        [TestCase("Const Bar = 42: Foo")]
        [TestCase("Const Bar = 42\nFoo")]
        [TestCase("Dim bar As Long: Foo")]
        [TestCase("Dim bar As Long\nFoo")]
        [TestCase("Label: Const Bar = 42: Foo")]
        [TestCase("Label:\nConst Bar = 42:\nFoo")]
        [TestCase("Foo 'This is a comment")]
        [TestCase("Foo '@Ignore EmptyMethod")]
        [TestCase("Call Foo: Const Bar = 42: Dim baz As Long")]
        [TestCase("Call Foo\nConst Bar = 42\nDim baz As Long")]
        public void EmptyMethod_VariousStatements_DontReturnResult(string statement)
        {
            string inputCode =
                $@"Sub Qux()
{statement}
End Sub

Sub Foo()
    MsgBox ""?""
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_DeclareStatement_NoResult()
        {
            string inputCode =
                $@"
Private Declare PtrSafe Function GetKeyState Lib ""user32.dll"" (ByVal nVirtKey As Long) As Integer
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string interfaceCode, string concreteCode, int expectedCount)
        {
            var results = InspectionResultsForModules(
                ("IClass1", interfaceCode, ComponentType.ClassModule),
                ("Class1", concreteCode, ComponentType.ClassModule));

            Assert.AreEqual(expectedCount, results.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new EmptyMethodInspection(state);
        }
    }
}
