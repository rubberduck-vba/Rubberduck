using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplementedInterfaceMemberInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ImplementedInterfaceMember_InspectionName()
        {
            const string expectedName = nameof(ImplementedInterfaceMemberInspection);
            var inspection = new ImplementedInterfaceMemberInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void ImplementedInterfaceMember_NotImplemented_DoesNotReturnResult()
        {
            const string interfaceCode =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub
Public Sub DoSomethingElse(ByVal a As Integer)
End Sub";
            const string concreteCode =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
    MsgBox ""?""
End Sub
Public Sub IClass1_DoSomethingElse(ByVal a As Integer)
    MsgBox ""?""
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(("IClass1", interfaceCode), ("Class1", concreteCode), 0);
        }

        [Test]
        [Category("Inspections")]
        public void ImplementedInterfaceMember_Implemented_ReturnsResult()
        {
            const string interfaceCode =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub
Public Sub DoSomethingElse(ByVal a As Integer)
    MsgBox ""?""
End Sub";
            const string concreteCode =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
    MsgBox ""?""
End Sub
Public Sub IClass1_DoSomethingElse(ByVal a As Integer)
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(("IClass1", interfaceCode), ("Class1", concreteCode), 1);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Label:")]
        [TestCase("Rem This is a rem comment")]
        [TestCase("'This is a comment")]
        [TestCase("'@Ignore EmptyMethod")]
        [TestCase("")]
        public void ImplementedInterfaceMember_VariousStatements_DontReturnResult(string statement)
        {
            string interfaceCode =
                $@"Sub Qux()
{statement}
End Sub";

            string concreteCode =
                @"Sub IClass1_Qux()
    MsgBox ""?""
End Sub";

            CheckActualEmptyBlockCountEqualsExpected(("IClass1", interfaceCode), ("Class1", concreteCode), 0);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Label: Foo")]
        [TestCase("Label:\nFoo")]
        [TestCase("Dim bar As Long")]
        [TestCase("Const Bar = 42: Dim baz As Long")]
        [TestCase("Const Bar = 42\nDim baz As Long")]
        [TestCase("Label: Const Bar = 42: Dim baz As Long")]
        [TestCase("Label:\nConst Bar = 42\nDim baz As Long")]
        [TestCase("Const Bar = 42: Foo")]
        [TestCase("Const Bar = 42\nFoo")]
        [TestCase("Dim bar As Long: Foo")]
        [TestCase("Dim bar As Long\nFoo")]
        [TestCase("Label: Const Bar = 42: Foo")]
        [TestCase("Label:\nConst Bar = 42:\nFoo")]
        [TestCase("Foo 'This is a comment")]
        [TestCase("Foo '@Ignore EmptyMethod")]
        [TestCase("Call Foo: Const Bar = 42: Dim baz As Long")]
        [TestCase("Foo\nConst Bar = 42\nDim baz As Long")]
        [TestCase("Const Bar = 42")]
        public void ImplementedInterfaceMember_VariousStatements_ReturnResult(string statement)
        {
            string interfaceCode =
                $@"Sub Qux()
{statement}
End Sub";

            string concreteCode =
                @"Implements IClass1

Sub IClass1_Qux()
End Sub";

            CheckActualEmptyBlockCountEqualsExpected(("IClass1", interfaceCode), ("Class1", concreteCode), 1);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5143
        [TestCase(@"MsgBox ""?""","", 1)]   //No implementers, only the annotation marks interface class
        [TestCase("", "", 0)]   //Annotated only, but no implementers - no result
        [TestCase(@"MsgBox ""?""", "Implements IClass1", 1)] //Annotated and an Implementer yields a single inspection result
        [Category("Inspections")]
        public void ImplementedInterfaceMember_AnnotatedOnly_ReturnsResult(string interfaceBody, string implementsStatement, int expected)
        {
            var interfaceCode =
$@"
'@Interface

Public Sub DoSomething(ByVal a As Integer)
End Sub
Public Sub DoSomethingElse(ByVal a As Integer)
    {interfaceBody}
End Sub";
            var concreteCode =
$@"

{implementsStatement}

Private Sub IClass1_DoSomething(ByVal a As Integer)
    MsgBox ""?""
End Sub
Public Sub IClass1_DoSomethingElse(ByVal a As Integer)
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(("IClass1", interfaceCode), ("Class1", concreteCode), expected);
        }

        private void CheckActualEmptyBlockCountEqualsExpected((string identifier, string code) interfaceDef, (string identifier, string code) implementerDef, int expectedCount)
        {
            var modules = new(string, string, ComponentType)[]
            {
                (interfaceDef.identifier, interfaceDef.code, ComponentType.ClassModule),
                (implementerDef.identifier, implementerDef.code, ComponentType.ClassModule)
            };

            Assert.AreEqual(expectedCount, InspectionResultsForModules(modules).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplementedInterfaceMemberInspection(state);
        }
    }
}
