using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class NonReturningFunctionInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningPropertyGet_ReturnsResult()
        {
            const string inputCode =
@"Property Get Foo() As Boolean
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleFunctions()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function

Function Goo() As String
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_DoesNotReturnResult()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Foo = True
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore NonReturningFunction
Function Foo() As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_MultipleSubs_SomeReturning()
        {
            const string inputCode =
@"Function Foo() As Boolean
    Foo = True
End Function

Function Goo() As String
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_ReturnsResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
@"Function Foo() As Boolean
End Function";
            const string inputCode2 =
@"Implements IClass1

Function IClass1_Foo() As Boolean
End Function";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_QuickFixWorks_Function()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenFunctionNameWithTypeHint_SubNameHasNoTypeHint()
        {
            const string inputCode =
@"Function Foo$()
End Function";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_QuickFixWorks_FunctionReturnsImplicitVariant()
        {
            const string inputCode =
@"Function Foo()
End Function";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_QuickFixWorks_FunctionHasVariable()
        {
            const string inputCode =
@"Function Foo(ByVal b As Boolean) As String
End Function";

            const string expectedCode =
@"Sub Foo(ByVal b As Boolean)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenNonReturningPropertyGetter_QuickFixConvertsToSub()
        {
            const string inputCode =
@"Property Get Foo() As Boolean
End Property";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenNonReturningPropertyGetWithTypeHint_QuickFixDropsTypeHint()
        {
            const string inputCode =
@"Property Get Foo$()
End Property";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenImplicitVariantPropertyGetter_StillConvertsToSub()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            const string expectedCode =
@"Sub Foo()
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void GivenParameterizedPropertyGetter_QuickFixKeepsParameter()
        {
            const string inputCode =
@"Property Get Foo(ByVal b As Boolean) As String
End Property";

            const string expectedCode =
@"Sub Foo(ByVal b As Boolean)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ConvertToProcedureQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void NonReturningFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Function Foo() As Boolean
End Function";

            const string expectedCode =
@"'@Ignore NonReturningFunction
Function Foo() As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new NonReturningFunctionInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new NonReturningFunctionInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        public void InspectionName()
        {
            const string inspectionName = "NonReturningFunctionInspection";
            var inspection = new NonReturningFunctionInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
