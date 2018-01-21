using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveExplicitByRefModifierQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
                @"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
                @"Sub Foo(Optional ByRef _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
                @"Sub Foo( ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
                @"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementationDiffrentParameterName()
        {
            const string inputCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg2 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_InterfaceImplementationWithMultipleParameters()
        {
            const string inputCode1 =
                @"Sub Foo(ByRef arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, ByRef arg2 as Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(
                    inspectionResults.First(
                        result =>
                            ((VBAParser.ArgContext)result.Context).unrestrictedIdentifier()
                            .identifier()
                            .untypedIdentifier()
                            .identifierValue()
                            .GetText() == "arg1"));

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_QuickFixWorks_PassByRef()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}
