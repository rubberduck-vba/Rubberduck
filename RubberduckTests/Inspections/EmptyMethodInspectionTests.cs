using System.Linq;
using System.Threading;
using NUnit.Framework;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    class EmptyMethodInspectionTests
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
        public void EmptyMethodInConcreteClassMarkedAsInterface_ReturnsResult()
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
    MsgBox ""?""
End Sub";
            CheckActualEmptyBlockCountEqualsExpected(interfaceCode, concreteCode, 1);
        }

        [Test]
        [Category("Inspections")]
        public void EmptyMethod_ReturnsExpectedMessage()
        {
            string inputCode =
                $@"Property Get Bar()
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyMethodInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual("Property Get 'Bar' contains no executable statements.", actualResults.Select(result => result.Description).First());
            }
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
