using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UndeclaredVariableInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsResult()
        {
            const string inputCode =
                @"Sub Test()
    a = 42
    Debug.Print a
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UndeclaredVariableInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfDeclaredLocally()
        {
            const string inputCode =
                @"Sub Test()
    Dim a As Long
    a = 42
    Debug.Print a
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UndeclaredVariableInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfDeclaredModuleScope()
        {
            const string inputCode =
                @"Private a As Long
            
Sub Test()
    a = 42
    Debug.Print a
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UndeclaredVariableInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2525
        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfAnnotated()
        {
            const string inputCode =
                @"Sub Test()
    '@Ignore UndeclaredVariable
    a = 42
    Debug.Print a
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UndeclaredVariableInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }
    }
}
