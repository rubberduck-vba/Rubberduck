using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplementedInterfaceMemberInspectionTests
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

        private void CheckActualEmptyBlockCountEqualsExpected(string interfaceCode, string concreteCode, int expectedCount)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, interfaceCode)
                .AddComponent("Class1", ComponentType.ClassModule, concreteCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplementedInterfaceMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }

        }
    }
}
