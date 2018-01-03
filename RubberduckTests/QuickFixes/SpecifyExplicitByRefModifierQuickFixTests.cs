using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.Grammar;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SpecifyExplicitByRefModifierQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_ImplicitByRefParameter()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
                @"Sub Foo(Optional arg1 As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
                @"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(Optional _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
                @"Sub Foo(bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo(ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
                @"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
                @"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementationDifferentParameterName()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg2 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_QuickFixWorks_InterfaceImplementationWithMultipleParameters()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(ByRef arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer, arg2 as Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new SpecifyExplicitByRefModifierQuickFix(state).Fix(
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

    }
}
