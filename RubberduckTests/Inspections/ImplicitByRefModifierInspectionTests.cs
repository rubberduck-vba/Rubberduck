using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitByRefModifierInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer, arg2 As Date)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByRef()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_ReturnsResult_SomePassedByRefImplicitly()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer, ByRef arg2 As Date)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ParamArray()
        {
            const string inputCode =
@"Sub Foo(ParamArray arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_QuickFixWorks_ImplicitByRefParameter()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_IgnoredQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
@"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
@"Sub Foo(Optional arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            var project = vbe.Object.VBProjects[0];
            var interfaceComponent = project.VBComponents[0];
            var implementationComponent = project.VBComponents[1];

            Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
            Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(inspectionResults.First());

            var project = vbe.Object.VBProjects[0];
            var interfaceComponent = project.VBComponents[0];
            var implementationComponent = project.VBComponents[1];

            Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
            Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefModifierInspection(state);
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new SpecifyExplicitByRefModifierQuickFix(state, InspectionsHelper.GetLocator()).Fix(
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

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitByRefModifierInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitByRefModifierInspection";
            var inspection = new ImplicitByRefModifierInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
