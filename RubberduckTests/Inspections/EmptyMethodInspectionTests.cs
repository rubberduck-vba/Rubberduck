using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class EmptyMethodInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void EmptyMethodBlock_InspectionName()
        {
            const string expectedName = nameof(EmptyMethodInspection);
            var inspection = new EmptyMethodInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_DoesNotFireOnImplementedMethod()
        {
            const string inputCode =
                @"Sub Foo()
    MsgBox ""?""
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_FiresOnNotImplementedMethod()
        {
            const string inputCode =
                @"Sub Foo()
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
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
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 1);
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
            CheckActualEmptyBlockCountEqualsExpected(inputCode, 0);
        }

        private void CheckActualEmptyBlockCountEqualsExpected(string inputCode, int expectedCount)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyMethodInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
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

                var inspection = new EmptyMethodInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }

        }
    }
}
