using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitByRefModifierInspectionTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer, arg2 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByRef()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_SomePassedByRefImplicitly()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer, ByRef arg2 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ParamArray()
        {
            const string inputCode =
                @"Sub Foo(ParamArray arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_MultipleInterfaceImplementations()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string inputCode3 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void InspectionType()
        {
            var inspection = new ImplicitByRefModifierInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [Test]
        [Category("QuickFixes")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitByRefModifierInspection";
            var inspection = new ImplicitByRefModifierInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
