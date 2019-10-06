using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.Inspections.Abstract;
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
            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 0);
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
            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 1);
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

            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 0);
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

            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 1);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5143
        [TestCase(@"MsgBox ""?""","", 1)]   //No implementers, only the annotation marks interface class
        [TestCase("", "", 0)]   //Annotated only, but no implementers - no result
        [TestCase(@"MsgBox ""?""", "Implements IClass1", 1)] //Annotated and an Implementer yields a single inspection result
        [Category("Inspections")]
        public void ImplementedInterfaceMember_AnnotatedOnly_ReturnsResult(string interfaceBody, string implementsStatement, int expected)
        {
            string interfaceCode =
$@"
'@Interface

Public Sub DoSomething(ByVal a As Integer)
End Sub
Public Sub DoSomethingElse(ByVal a As Integer)
    {interfaceBody}
End Sub";
            string concreteCode =
$@"

{implementsStatement}

Private Sub IClass_DoSomething(ByVal a As Integer)
    MsgBox ""?""
End Sub
Public Sub IClass_DoSomethingElse(ByVal a As Integer)
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, expected);
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string interfaceCode, string concreteCode, int expectedCount)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, interfaceCode)
                .AddComponent("Class1", ComponentType.ClassModule, concreteCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var inspectionResults = InspectionResults(vbe.Object);
            Assert.AreEqual(expectedCount, inspectionResults.Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplementedInterfaceMemberInspection(state);
        }
    }
}
