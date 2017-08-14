using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class RedundantByRefModifierInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) {Severity = CodeInspectionSeverity.Hint};
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer, ByRef arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_DoesNotReturnResult_NoModifier()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_SomePassedByRef()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer, ByRef arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_MultipleInterfaceImplementation()
        {
            const string inputCode1 =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode3 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore RedundantByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_QuickFixWorks_PassByRef()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_IgnoredQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"'@Ignore RedundantByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void RedundantByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
@"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(Optional arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

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

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

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

        [TestMethod]
        [TestCategory("Inspections")]
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

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
            var inspector = InspectionsHelper.GetInspector(inspection);
            var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            new RemoveExplicitByRefModifierQuickFix(state).Fix(
                inspectionResults.First(
                    result =>
                        ((VBAParser.ArgContext) result.Context).unrestrictedIdentifier()
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
            var inspection = new RedundantByRefModifierInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "RedundantByRefModifierInspection";
            var inspection = new RedundantByRefModifierInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
